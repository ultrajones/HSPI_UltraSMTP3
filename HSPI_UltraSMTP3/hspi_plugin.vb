Imports System.Threading
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Net
Imports System.Data.Common
Imports HomeSeerAPI
Imports Scheduler
Imports System.IO
Imports System.ComponentModel
Imports System.Net.Mail
Imports System.Net.Mime
Imports System.Data.SQLite

Module hspi_plugin

  '
  ' Declare public objects, not required by HomeSeer
  '
  Dim actions As New hsCollection
  Dim triggers As New hsCollection
  Dim conditions As New Hashtable
  Const Pagename = "Events"

  Public SmtpProfiles As New ArrayList

  Public Const IFACE_NAME As String = "UltraSMTP3"

  Public Const LINK_TARGET As String = "hspi_ultrasmtp3/hspi_ultrasmtp3.aspx"
  Public Const LINK_URL As String = "hspi_ultrasmtp3.html"
  Public Const LINK_TEXT As String = "UltraSMTP3"
  Public Const LINK_PAGE_TITLE As String = "UltraSMTP3 HSPI"
  Public Const LINK_HELP As String = "/hspi_ultrasmtp3/UltraSMTP3_HSPI_Users_Guide.pdf"

  Public gBaseCode As String = ""
  Public gIOEnabled As Boolean = True
  Public gImageDir As String = "/images/hspi_ultrasmtp3/"
  Public gHSInitialized As Boolean = False
  Public gINIFile As String = "HSPI_" & UCase(IFACE_NAME) & ".ini"

  Public PickupDirectoryLocation As String = ""
  Public DropDirectoryLocation As String = ""

  Public gMaxAttachmentSize As Integer = 2048
  Public gMaxAttempts As Integer = 10

  Public HSAppPath As String = ""

  Public SMTPQueueThread As Thread

#Region "HSPI - Public Sub/Functions"

#Region "HSPI - Public Sub/Functions"

  ''' <summary>
  ''' Returns number of queued messages
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetQueueCount() As Integer

    Try
      Return Directory.GetFiles(PickupDirectoryLocation, "*.eml").Length()
    Catch pEx As Exception
      Return 0
    End Try
  End Function

  ''' <summary>
  ''' Builds the Smtp Delivery Log SQL query
  ''' </summary>
  ''' <param name="dbTable"></param>
  ''' <param name="dbFields"></param>
  ''' <param name="startDate"></param>
  ''' <param name="endDate"></param>
  ''' <param name="filterField"></param>
  ''' <param name="filterCompare"></param>
  ''' <param name="filterValue"></param>
  ''' <param name="sortOrder"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BuildSmtpDeliveryLogSQL(ByVal dbTable As String, _
                                          ByVal dbFields As String, _
                                          ByVal startDate As String, _
                                          ByVal endDate As String, _
                                          ByVal filterField As String,
                                          ByVal filterCompare As String, _
                                          ByVal filterValue As String,
                                          ByVal sortOrder As String) As String

    Try

      Dim dStartDateTime As Date
      Date.TryParse(startDate, dStartDateTime)

      Dim dEndDateTime As Date
      Date.TryParse(endDate, dEndDateTime)

      Dim ts_start = ConvertDateTimeToEpoch(dStartDateTime)
      Dim ts_end = ConvertDateTimeToEpoch(dEndDateTime)

      Dim dateField As String = "queue_ts"

      '
      ' Build Filter Query
      '
      Dim filterString1 As String = String.Empty
      If filterValue.Trim <> "" Then
        Select Case filterField
          Case "queue_id"
            If IsNumeric(filterValue) = True Then
              filterString1 = String.Format("AND {0} = {1} ", filterField, filterValue)
            End If
          Case Else
            Select Case filterCompare.ToLower
              Case "is"
                filterString1 = String.Format("AND {0} = '{1}' ", filterField, filterValue)
              Case "starts with"
                filterString1 = String.Format("AND {0} LIKE '{1}%' ", filterField, filterValue)
              Case "ends with"
                filterString1 = String.Format("AND {0} LIKE '%{1}' ", filterField, filterValue)
              Case "contains"
                filterString1 = String.Format("AND {0} LIKE '%{1}%' ", filterField, filterValue)
            End Select
        End Select

      End If

      '
      ' Build SQL
      '
      Dim strSQL As String = String.Format("SELECT {0} " _
                                         & "FROM {1} " _
                                         & "WHERE {2} >= {3} " _
                                         & "AND {2} <= {4} " _
                                         & filterString1 _
                                         & "ORDER BY {2} {5}", dbFields, dbTable, dateField, ts_start.ToString, ts_end.ToString, sortOrder)

      WriteMessage(strSQL, MessageType.Debug)

      Return strSQL

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "BuildSmtpDeliveryLogSQL()")
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' Sends an SMTP e-mail
  ''' </summary>
  ''' <param name="ToAddr"></param>
  ''' <param name="Subject"></param>
  ''' <param name="Body"></param>
  ''' <param name="AttachmentPaths"></param>
  ''' <param name="smtp_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SendMail(ByVal ToAddr As String, _
                           ByVal Subject As String, _
                           ByVal Body As String, _
                           Optional ByVal AttachmentPaths() As String = Nothing, _
                           Optional ByVal smtp_id As Integer = 0) As Boolean

    Dim myMessage As MailMessage
    Dim attachment As Mail.Attachment

    Try

      Subject = Subject.Replace("$date", DateTime.Now.ToLongDateString)
      Subject = Subject.Replace("$time", DateTime.Now.ToShortTimeString)

      Dim queue_id As String = Database.InsertQueuedMessage(smtp_id, ToAddr, Subject)

      Dim strMessage As String = String.Format("Generating message '{0}' to '{1}' with subject '{2}'", queue_id, ToAddr, Subject)
      WriteMessage(strMessage, MessageType.Debug)

      '
      ' Process if queue Id was issused
      '
      If queue_id > 0 Then

        myMessage = New MailMessage("nobody@homeseer.com", ToAddr)

        myMessage.Headers.Add("X-QueueId", queue_id)
        myMessage.Headers.Add("X-SmtpId", smtp_id)

        myMessage.Subject = Subject
        myMessage.Body = String.Empty

        If Not AttachmentPaths Is Nothing Then
          If AttachmentPaths.Length > 0 Then
            For index As Integer = 0 To AttachmentPaths.Length - 1
              Dim AttachmentPath As String = AttachmentPaths(index)
              If IO.File.Exists(AttachmentPath) Then
                attachment = New Mail.Attachment(AttachmentPath)
                myMessage.Attachments.Add(attachment)
                WriteMessage(String.Format("Successfully added attachment '{0}' to message '{1}'", AttachmentPath, queue_id), MessageType.Debug)
              Else
                '
                ' Need to indicate an error occured here
                '
                WriteMessage(String.Format("Unable to add attachment '{0}' to message '{1}' because the file does not exist.", AttachmentPath, queue_id), MessageType.Error)
              End If
            Next
          End If
        End If

        Dim Data As Byte() = Encoding.UTF8.GetBytes(ExpandHomeSeerVariables(Body))

        ' Create the body attachment
        Dim contentType As New ContentType()
        contentType.MediaType = MediaTypeNames.Application.Octet
        contentType.Name = "MessageBody.txt"
        Dim BodyAttachment As New Attachment(New MemoryStream(Data), contentType)

        ' Add the attachment
        myMessage.Attachments.Add(BodyAttachment)
        myMessage.Priority = MailPriority.Low

        For Each attachment In myMessage.Attachments
          strMessage = String.Format("Message '{0}' contains attachment '{1}'", queue_id, attachment.Name)
          WriteMessage(strMessage, MessageType.Debug)
        Next

        Dim mySmtpClient As SmtpClient = New SmtpClient()

        mySmtpClient.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory
        mySmtpClient.PickupDirectoryLocation = PickupDirectoryLocation
        mySmtpClient.Send(myMessage)

        myMessage.Dispose()

        If SMTPQueueThread.ThreadState = ThreadState.WaitSleepJoin Then
          SMTPQueueThread.Interrupt()
        End If

        Return True

      Else
        WriteMessage(String.Format("Unable to send Email to {0} because the database returned an invalid Queue Id.", ToAddr), MessageType.Error)

        Return False
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SendMail()")
      Return False
    Finally

    End Try

  End Function

  ''' <summary>
  ''' Gets the Smtp Profiles from the underlying database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSmtpProfiles() As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered GetSmtpProfiles() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("SELECT * FROM tblSmtpProfiles")

      '
      ' Initialize the command object
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      '
      ' Initialize the dataset, then populate it
      '
      Dim MyDS As DataSet = New DataSet

      Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
      MyDA.SelectCommand = MyDbCommand

      SyncLock SyncLockMain
        MyDA.Fill(MyDS)
      End SyncLock

      '
      ' Get our DataTable
      '
      Dim MyDT As DataTable = MyDS.Tables(0)

      '
      ' Get record count
      '
      Dim iRecordCount As Integer = MyDT.Rows.Count

      If iRecordCount > 0 Then
        '
        ' Build field names
        '
        Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
        For iFieldNum As Integer = 0 To iFieldCount
          '
          ' Create the columns
          '
          Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
          Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

          '
          ' Add the columns to the DataTable's Columns collection
          '
          ResultsDT.Columns.Add(MyDataColumn)
        Next

        '
        ' Let's output our records	
        '
        Dim i As Integer = 0
        For i = 0 To iRecordCount - 1
          '
          ' Create the rows
          '
          Dim dr As DataRow
          dr = ResultsDT.NewRow()
          For iFieldNum As Integer = 0 To iFieldCount
            dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
          Next
          ResultsDT.Rows.Add(dr)
        Next

      End If

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "GetSmtpProfiles()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Inserts a new Smtp Profile into the database
  ''' </summary>
  ''' <param name="smtp_server"></param>
  ''' <param name="smtp_port"></param>
  ''' <param name="smtp_ssl"></param>
  ''' <param name="auth_user"></param>
  ''' <param name="auth_pass"></param>
  ''' <param name="mail_from"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InsertSmtpProfile(ByVal smtp_server As String, _
                                    ByVal smtp_port As Integer, _
                                    ByVal smtp_ssl As Integer, _
                                    ByVal auth_user As String, _
                                    ByVal auth_pass As String, _
                                    ByVal mail_from As String) As Integer

    Dim strMessage As String = ""
    Dim smtp_id As Integer = 0

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If smtp_server.Length = 0 Or smtp_port = 0 Or mail_from.Length = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to insert new SMTP profile into the database.")
      End If

      auth_pass = hs.EncryptString(auth_pass, "&Cul8r#1")
      Dim strSQL As String = String.Format("INSERT INTO tblSmtpProfiles (" _
                                           & " smtp_server, smtp_port, smtp_ssl, auth_user, auth_pass, mail_from" _
                                           & ") VALUES (" _
                                           & "'{0}', {1}, {2}, '{3}', '{4}', '{5}' );", _
                                           smtp_server, smtp_port, smtp_ssl, auth_user, auth_pass, mail_from)
      strSQL &= "SELECT last_insert_rowid();"

      Dim dbcmd As DbCommand = DBConnectionMain.CreateCommand()

      dbcmd.Connection = DBConnectionMain
      dbcmd.CommandType = CommandType.Text
      dbcmd.CommandText = strSQL

      SyncLock SyncLockMain
        smtp_id = dbcmd.ExecuteScalar()
      End SyncLock

      dbcmd.Dispose()

    Catch pEx As Exception
      Call ProcessError(pEx, "InsertSmtpProfile()")
    End Try

    Return smtp_id

  End Function

  ''' <summary>
  ''' Updates existing Smtp Profile stored in the database
  ''' </summary>
  ''' <param name="smtp_id"></param>
  ''' <param name="smtp_server"></param>
  ''' <param name="smtp_port"></param>
  ''' <param name="smtp_ssl"></param>
  ''' <param name="auth_user"></param>
  ''' <param name="auth_pass"></param>
  ''' <param name="mail_from"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UpdateSmtpProfile(ByVal smtp_id As Integer, _
                                    ByVal smtp_server As String, _
                                    ByVal smtp_port As Integer, _
                                    ByVal smtp_ssl As Integer, _
                                    ByVal auth_user As String, _
                                    ByVal auth_pass As String, _
                                    ByVal mail_from As String) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If smtp_server.Length = 0 Or smtp_port = 0 Or mail_from.Length = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to save SMTP profile update to database.")
      End If

      Dim strSql As String = ""

      If auth_user.Length > 0 And auth_pass.Length > 0 Then

        auth_pass = hs.EncryptString(auth_pass, "&Cul8r#1")
        strSql = String.Format("UPDATE tblSmtpProfiles SET " _
                            & " smtp_server='{0}', " _
                            & " smtp_port={1}," _
                            & " smtp_ssl={2}," _
                            & " auth_user='{3}'," _
                            & " auth_pass='{4}'," _
                            & " mail_from='{5}'" _
                            & "WHERE smtp_id={6}", smtp_server, smtp_port.ToString, smtp_ssl.ToString, auth_user, auth_pass, mail_from, smtp_id.ToString)

      ElseIf auth_user.Length > 0 Then

        strSql = String.Format("UPDATE tblSmtpProfiles SET " _
                            & " smtp_server='{0}', " _
                            & " smtp_port={1}," _
                            & " smtp_ssl={2}," _
                            & " auth_user='{3}'," _
                            & " mail_from='{4}'" _
                            & "WHERE smtp_id={5}", smtp_server, smtp_port.ToString, smtp_ssl.ToString, auth_user, mail_from, smtp_id.ToString)

      Else

        auth_user = ""
        auth_pass = ""

        strSql = String.Format("UPDATE tblSmtpProfiles SET " _
                            & " smtp_server='{0}', " _
                            & " smtp_port={1}," _
                            & " smtp_ssl={2}," _
                            & " auth_user='{3}'," _
                            & " auth_pass='{4}'," _
                            & " mail_from='{5}'" _
                            & "WHERE smtp_id={6}", smtp_server, smtp_port.ToString, smtp_ssl.ToString, auth_user, auth_pass, mail_from, smtp_id.ToString)

      End If


      '
      ' Build the insert/update/delete query
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSql

      Dim iRecordsAffected As Integer = 0
      SyncLock SyncLockMain
        iRecordsAffected = MyDbCommand.ExecuteNonQuery()
      End SyncLock

      strMessage = "UpdateSmtpProfile() updated " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      MyDbCommand.Dispose()

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "UpdateSmtpProfile()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Removes existing Smtp Profile stored in the database
  ''' </summary>
  ''' <param name="smtp_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeleteSmtpProfile(ByVal smtp_id As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      '
      ' Build the insert/update/delete query
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = String.Format("DELETE FROM tblSmtpProfiles WHERE smtp_id={0}", smtp_id.ToString)

      Dim iRecordsAffected As Integer = 0
      SyncLock SyncLockMain
        iRecordsAffected = MyDbCommand.ExecuteNonQuery()
      End SyncLock

      strMessage = "DeleteSmtpProfile() removed " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      MyDbCommand.Dispose()

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

      Return True

    Catch pEx As Exception
      Call ProcessError(pEx, "DeleteSmtpProfile()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Sends SMTP Email to selected server
  ''' </summary>
  ''' <param name="smtp_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SendSmtpProfileTest(ByVal smtp_id As Integer) As Boolean

    Try

      Call RefreshSmtpProfiles()

      For Each SmtpProfile As SmtpProfile In SmtpProfiles
        If SmtpProfile.SmtpId = smtp_id Then
          SendMail(SmtpProfile.SmtpFrom, "Test Email", "This is a test Email.", Nothing, SmtpProfile.SmtpId)
          Return True
        End If
      Next

      Return False
    Catch pEx As Exception
      Return False
    End Try

  End Function

#End Region

#Region "HSPI - Misc"

  ''' <summary>
  ''' Gets plug-in setting from INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValueDefault"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSetting(ByVal strSection As String, _
                             ByVal strKey As String, _
                             ByVal strValueDefault As String) As String

    Dim strMessage As String = ""

    Try
      strMessage = "Entered GetSetting() function."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Get the ini settings
      '
      Dim strValue As String = hs.GetINISetting(strSection, strKey, strValueDefault, gINIFile)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to decrypt the data
      '
      If strKey = "UserPass" Then
        strValue = hs.DecryptString(strValue, "&Cul8r#1")
      End If

      Return strValue

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetSetting()")
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' Saves plug-in setting to INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValue"></param>
  ''' <remarks></remarks>
  Public Sub SaveSetting(ByVal strSection As String, _
                         ByVal strKey As String, _
                         ByVal strValue As String)

    Dim strMessage As String = ""

    Try
      strMessage = "Entered SaveSetting() subroutine."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to encrypt the data
      '
      If strKey = "UserPass" Then
        If strValue.Length = 0 Then Exit Sub
        strValue = hs.EncryptString(strValue, "&Cul8r#1")
      End If

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Save selected settings to global variables
      '
      If strSection = "Options" And strKey = "MaxDeliveryAttempts" Then
        If IsNumeric(strValue) Then
          gMaxAttempts = CInt(Val(strValue))
        End If
      End If

      '
      ' Save the settings
      '
      hs.SaveINISetting(strSection, strKey, strValue, gINIFile)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SaveSetting()")
    End Try

  End Sub

#End Region

#End Region

#Region "SMTP Threads"

  Dim SMTPQueue As New Specialized.StringDictionary

  ''' <summary>
  ''' Mail queue thread
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ProcessSMTPQueue()

    Dim bAbortThread As Boolean = False

    Dim DirInfo As New DirectoryInfo(PickupDirectoryLocation)

    Try
      WriteMessage("The ProcessSMTPQueue thread has started ...", MessageType.Debug)

      While bAbortThread = False

        Try
          '
          ' Get all the pending .EML files needing to be sent
          '
          Dim aryFi As IO.FileInfo() = DirInfo.GetFiles("*.eml")

          '
          ' Process each message found
          '
          For Each fi As IO.FileInfo In aryFi

            Try

              Dim strFileSize As String = (Math.Round(fi.Length / 1024)).ToString()

              WriteMessage(String.Format("File Name: {0}", fi.Name), MessageType.Debug)
              WriteMessage(String.Format("File Full Name: {0}", fi.FullName), MessageType.Debug)
              WriteMessage(String.Format("File Size (KB): {0}", strFileSize), MessageType.Debug)
              WriteMessage(String.Format("File Extension: {0}", fi.Extension), MessageType.Debug)
              WriteMessage(String.Format("Last Accessed: {0}", fi.LastAccessTime), MessageType.Debug)

              Dim MailMessage As New MailMessage()
              Dim MailSmtpId As Integer = 0
              Dim QueueId As Integer = 0

              Dim sr As StreamReader = fi.OpenText()
              Dim s As String = String.Empty
              Dim strPeakLine As String = String.Empty

              Dim MailBody As New StringBuilder

              Dim ContentTransferEncoding As String = ""
              Dim ContentType As String = ""

              Dim AttachmentData As New StringBuilder
              Dim AttachmentName As String = ""
              Dim AttachmentBoundry As String = ""

              Dim bHeaderProcessed As Boolean = False
              Dim bReadAttachments As Boolean = False

              '
              ' Begin parsing the e-mail
              '
              While sr.EndOfStream = False
                '
                ' Process Line
                '
                If strPeakLine.Length = 0 Then
                  s = sr.ReadLine()
                Else
                  s = strPeakLine
                  strPeakLine = String.Empty
                End If

                '
                ' Process the Message Headers
                '
                If bHeaderProcessed = False And Regex.IsMatch(s, "^[A-Za-z0-9-]+: ") = True Then

                  If Regex.IsMatch(s, "Content-Type: multipart/mixed; boundary=", RegexOptions.IgnoreCase) = True Then

                    Dim regexPattern As String = "multipart/mixed; boundary=(?<boundry>(.+))"
                    AttachmentBoundry = Regex.Match(s, regexPattern, RegexOptions.IgnoreCase).Groups("boundry").ToString()
                    bReadAttachments = True

                  ElseIf Regex.IsMatch(s, "Content-Type: multipart/mixed;", RegexOptions.IgnoreCase) = True Then
                    s = sr.ReadLine()

                    Dim regexPattern As String = "\sboundary=(?<boundry>(.+))"
                    AttachmentBoundry = Regex.Match(s, regexPattern, RegexOptions.IgnoreCase).Groups("boundry").ToString()
                    bReadAttachments = True

                  Else

                    Dim arrHeaders() As String = s.Split(":", 2, StringSplitOptions.None)
                    Dim strHeader As String = arrHeaders(0).ToLower
                    Dim strHeaderValue As String = arrHeaders(1).Trim
                    Select Case strHeader
                      Case "Subject".ToLower
                        strPeakLine = sr.ReadLine
                        If Regex.IsMatch(strPeakLine, "^\s.+") = True Then
                          strHeaderValue &= strPeakLine
                          strPeakLine = String.Empty
                        End If
                        MailMessage.Subject = strHeaderValue
                      Case "X-Sender".ToLower
                        'MailMessage.From = New MailAddress(strHeaderValue)
                      Case "X-Receiver".ToLower
                        MailMessage.To.Add(New MailAddress(strHeaderValue))
                      Case "X-QueueId".ToLower
                        QueueId = Val(arrHeaders(1))
                        MailMessage.Headers.Add(strHeader, strHeaderValue)
                      Case "X-SmtpId".ToLower
                        MailSmtpId = Val(arrHeaders(1))
                        MailMessage.Headers.Add(strHeader, strHeaderValue)
                    End Select
                  End If

                Else
                  If Regex.IsMatch(s, "^$") = True Then
                    bHeaderProcessed = True
                  End If
                End If

                '
                ' Process the Message Body
                '
                If bHeaderProcessed = True Then
                  If bReadAttachments = True Then
                    If Regex.IsMatch(s, "^----") = True Then
                      '
                      ' Begin/End processing of attachment
                      '
                      If AttachmentData.Length > 0 Then
                        Select Case ContentTransferEncoding.ToLower
                          Case "Content-Transfer-Encoding: base64".ToLower
                            '
                            ' Process base64 encoded content
                            '
                            Dim regexPattern As String = "name=(?<filename>([^\s]+))"
                            If Regex.IsMatch(ContentType, regexPattern, RegexOptions.IgnoreCase) = True Then
                              Try
                                AttachmentName = Regex.Match(ContentType, regexPattern).Groups("filename").ToString()
                                Dim data As Byte() = Convert.FromBase64String(AttachmentData.ToString.Trim)

                                If Regex.IsMatch(AttachmentName, "MessageBody.txt", RegexOptions.IgnoreCase) = True Then
                                  MailBody.Append(System.Text.Encoding.UTF8.GetString(data))
                                Else
                                  ' Save the data to a memory stream
                                  Dim ms As New MemoryStream(data)

                                  ' Create the attachment from a stream. Be sure to name the data with a file and 
                                  ' media type that is respective of the data
                                  MailMessage.Attachments.Add(New Attachment(ms, AttachmentName))
                                End If
                              Catch pEx As Exception
                                '
                                ' Process Exception
                                '
                                ProcessError(pEx, "ProcessSMTPQueue()")
                              End Try

                            End If

                          Case "Content-Transfer-Encoding: quoted-printable".ToLower
                            '
                            ' Process quoted printable content
                            '
                        End Select

                      End If

                      ContentTransferEncoding = ""
                      ContentType = ""

                      AttachmentName = String.Empty
                      AttachmentData.Length = 0

                    ElseIf Regex.IsMatch(s, "^Content-Type:", RegexOptions.IgnoreCase) = True Then
                      '
                      ' This is a MIME header
                      '
                      ContentType = s

                      If Regex.IsMatch(s, ";$") = True Then
                        s = sr.ReadLine
                        ContentType += s
                      End If

                    ElseIf Regex.IsMatch(s, "^Content-Transfer-Encoding:", RegexOptions.IgnoreCase) = True Then
                      '
                      ' This is a MIME header
                      '
                      ContentTransferEncoding = s

                    ElseIf Regex.IsMatch(s, "^Content", RegexOptions.IgnoreCase) = True Then
                      '
                      ' This is some other content description that we don't care about
                      '
                    ElseIf Regex.IsMatch(s, "^\r\n") = True Then
                      '
                      ' This is an emtpy line
                      '
                    ElseIf s.Length = 0 Then
                      '
                      ' This is an emtpy line
                      '
                    Else
                      '
                      ' This is raw encoded data
                      '
                      AttachmentData.Append(s)
                    End If

                  Else
                    '
                    ' This is the body of the message (which we ignore)
                    '
                  End If
                End If

              End While

              If MailBody.Length > 0 Then
                MailMessage.Body = MailBody.ToString
              End If

              '
              ' End message parsing
              '
              sr.Close()
              sr.Dispose()

              '
              ' Determine if the body of the message contains HTML
              '
              If MailMessage.Body.Contains("<html") = True Then
                MailMessage.IsBodyHtml = True
              End If

              '
              ' Attempt to send the message
              '
              Dim bMessageSent As Boolean = SendSMTPMessage(QueueId, MailSmtpId, MailMessage)

              '
              ' Process the results
              '
              Dim iAttempts As Integer = Database.GetDeliveryAttempts(QueueId)
              Dim iFileAction As Integer = FileAction.Keep
              Dim iDeliveryResult As Integer = DeliveryResult.Failure

              If bMessageSent = True Then
                '
                ' Delivery successful
                '
                Dim strMessage As String = String.Format("Email Queue Id {0} successfully delivered to {1}.", QueueId.ToString.PadLeft(6, "0"), MailMessage.To.ToString)
                WriteMessage(strMessage, MessageType.Informational)

                iDeliveryResult = DeliveryResult.Success
                iFileAction = FileAction.Delete
              ElseIf iAttempts >= gMaxAttempts Then
                '
                ' Delivery exceeded maximum retry count
                '
                Dim strMessage As String = String.Format("Email Queue Id {0} to {1} exceeded {2} retries; moving to drop queue.", QueueId.ToString.PadLeft(6, "0"), MailMessage.To.ToString, iAttempts)
                WriteMessage(strMessage, MessageType.Warning)

                iDeliveryResult = DeliveryResult.Failure
                iFileAction = FileAction.Move
              ElseIf iAttempts = -1 Then
                '
                ' Delivery caused unknown exception
                '
                Dim strMessage As String = String.Format("Email Queue Id {0} to {1} not found in the database; moving to drop queue.", QueueId.ToString.PadLeft(6, "0"), MailMessage.To.ToString)
                WriteMessage(strMessage, MessageType.Error)

                iDeliveryResult = DeliveryResult.Failure
                iFileAction = FileAction.Move
              Else
                '
                ' Delivery was deferred
                '
                Dim strMessage As String = String.Format("Email Queue Id {0} to {1} was deferred.", QueueId.ToString.PadLeft(6, "0"), MailMessage.To.ToString)
                WriteMessage(strMessage, MessageType.Error)

                iDeliveryResult = DeliveryResult.Deferred
                iFileAction = FileAction.Keep
              End If

              MailMessage.Dispose()

              '
              ' Determine what action to take on queued message
              '
              Select Case iFileAction
                Case FileAction.Delete
                  '
                  ' Delivery was successful, so delete the queued file
                  '
                  fi.Delete()
                Case FileAction.Move
                  '
                  ' Delivery failed, so move queued message to drop queue
                  '
                  Dim strFileNameDest As String = String.Format("{0}{1}", DropDirectoryLocation, fi.Name)

                  fi.CopyTo(strFileNameDest)
                  fi.Delete()

                  '
                  ' Update the queued message
                  '
                  Database.UpdateQueuedMessage(QueueId, "GeneralFailure", "Moved to drop queue", 0)

                Case FileAction.Keep
              End Select

              '
              ' Determine what HomeSeer triggers to check
              '
              Select Case iDeliveryResult
                Case DeliveryResult.Success
                  '
                  ' Check Delivery Trigger
                  '
                  Dim strTriggerName As String = GetEnumDescription(SMTPTriggers.EmailDeliveryStatus)
                  Dim strTrigger As String = String.Format("{0},{1}", strTriggerName, "Success")
                  CheckTrigger(IFACE_NAME, SMTPTriggers.EmailDeliveryStatus, -1, strTrigger)

                Case DeliveryResult.Deferred
                  '
                  ' Check Delivery Trigger
                  '
                  Dim strTriggerName As String = GetEnumDescription(SMTPTriggers.EmailDeliveryStatus)
                  Dim strTrigger As String = String.Format("{0},{1}", strTriggerName, "Deferral")
                  CheckTrigger(IFACE_NAME, SMTPTriggers.EmailDeliveryStatus, -1, strTrigger)

                Case DeliveryResult.Failure
                  '
                  ' Check Delivery Trigger
                  '
                  Dim strTriggerName As String = GetEnumDescription(SMTPTriggers.EmailDeliveryStatus)
                  Dim strTrigger As String = String.Format("{0},{1}", strTriggerName, "Failure")
                  CheckTrigger(IFACE_NAME, SMTPTriggers.EmailDeliveryStatus, -1, strTrigger)

              End Select

            Catch pEx As Exception
              '
              ' Process Exception
              '
              ProcessError(pEx, "ProcessSMTPQueue()")
            End Try

          Next

          '
          ' Give up some time
          '
          Thread.Sleep(60 * 1000)

        Catch pEx As ThreadInterruptedException
          '
          ' Thread sleep was interrupted
          '
          WriteMessage("ProcessSMTPQueue was interrupted.", MessageType.Debug)

        Catch pEx As Exception
          '
          ' Process Exception
          '
          Call ProcessError(pEx, "ProcessSMTPQueue()")
        End Try

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("ProcessSMTPQueue thread received abort request, terminating normally."), MessageType.Informational)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "ProcessSMTPQueue()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("ProcessSMTPQueue thread terminated."), MessageType.Debug)

    End Try

  End Sub

  ''' <summary>
  ''' Convert from Byte() to Base64
  ''' </summary>
  ''' <param name="data"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ToBase64(ByVal data() As Byte) As String
    If data Is Nothing Then Throw New ArgumentNullException("data")
    Return Convert.ToBase64String(data)
  End Function

  ''' <summary>
  ''' Convert From Base64 to Byte()
  ''' </summary>
  ''' <param name="base64"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function FromBase64(ByVal base64 As String) As Byte()
    If base64 Is Nothing Then Throw New ArgumentNullException("base64")
    Return Convert.FromBase64String(base64)
  End Function

  ''' <summary>
  ''' Send the SMTP message
  ''' </summary>
  ''' <param name="QueueID"></param>
  ''' <param name="MailSmtpId"></param>
  ''' <param name="MailMessage"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function SendSMTPMessage(ByVal QueueID As Integer, ByVal MailSmtpId As Integer, ByRef MailMessage As MailMessage) As Boolean

    Dim bMessageSent As Boolean = False

    Try
      '
      ' Refresh the list of SMTP profiles
      '
      RefreshSmtpProfiles()

      If SmtpProfiles.Count = 0 Then
        Database.UpdateQueuedMessage(QueueID, "GeneralFailure", "No SMTP servers are defined.", 0)
        Database.InsertSmtpLog(QueueID, 0, "", "GeneralFailure", "No SMTP servers are defined.", MailMessage.From.ToString, MailMessage.To.ToString, MailMessage.Subject)
        Return bMessageSent
      End If

      '
      ' Attempt to send the message
      '
      For Each SmtpProfile As SmtpProfile In SmtpProfiles.Clone

        Try

          '
          ' Skip if we are to use a particular Smtp Id
          '
          If MailSmtpId > 0 And SmtpProfile.SmtpId <> MailSmtpId Then Continue For

          MailMessage.Sender = New MailAddress(SmtpProfile.SmtpFrom)
          MailMessage.From = New MailAddress(SmtpProfile.SmtpFrom)

          Dim mySmtpClient As SmtpClient = New SmtpClient(SmtpProfile.SmtpServer, SmtpProfile.SmtpPort)

          mySmtpClient.UseDefaultCredentials = False
          mySmtpClient.EnableSsl = SmtpProfile.SmtpSsl

          Dim bEnableAuth As Boolean = False
          If SmtpProfile.AuthUser.Length > 0 And SmtpProfile.AuthPass.Length > 0 Then
            bEnableAuth = True
            mySmtpClient.Credentials = New NetworkCredential(SmtpProfile.AuthUser, SmtpProfile.AuthPass)
          End If

          '
          ' Write debug message
          '
          WriteMessage(String.Format("Email Queue Id {0} delivery attempted to STMP server {1}:{2} [ SSL={3}; Auth={4} ]", _
                                      QueueID.ToString.PadLeft(6, "0"), _
                                      SmtpProfile.SmtpServer, _
                                      SmtpProfile.SmtpPort, _
                                      mySmtpClient.EnableSsl.ToString, _
                                      bEnableAuth.ToString), MessageType.Debug)

          '
          ' Attempt to send the message
          '
          mySmtpClient.Send(MailMessage)

          '
          ' Record success
          '
          Database.InsertSmtpLog(QueueID, SmtpProfile.SmtpId, SmtpProfile.SmtpServer, "Delivered", "Success", MailMessage.From.ToString, MailMessage.To.ToString, MailMessage.Subject)

          bMessageSent = True
          Exit For

          'Catch pEx As SmtpFailedRecipientsException
          '  For i As Integer = 0 To pEx.InnerExceptions.Length - 1
          '    Dim status As SmtpStatusCode = pEx.InnerExceptions(i).StatusCode
          '     strRecipient as String = pEx.InnerExceptions(i).FailedRecipient
          '    If status = SmtpStatusCode.MailboxBusy Or status = SmtpStatusCode.MailboxUnavailable Then
          '
          '    End If
          '  Next

        Catch pEx As SmtpException
          '
          ' Process the error
          '
          Database.InsertSmtpLog(QueueID, SmtpProfile.SmtpId, SmtpProfile.SmtpServer, pEx.StatusCode.ToString, pEx.Message, MailMessage.From.ToString, MailMessage.To.ToString, MailMessage.Subject)

        Catch pEx As Exception
          '
          ' Record failure
          '
          Database.InsertSmtpLog(QueueID, SmtpProfile.SmtpId, SmtpProfile.SmtpServer, "GeneralFailure", pEx.Message, MailMessage.From.ToString, MailMessage.To.ToString, MailMessage.Subject)

          '
          ' Process the error
          '
          Call ProcessError(pEx, "SendSMTPMessage()")
        End Try

      Next

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SendSMTPMessage()")
    End Try

    Return bMessageSent

  End Function

  ''' <summary>
  ''' Refresh the available SMTP profiles
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub RefreshSmtpProfiles()

    Try

      Dim strSQL As String = String.Format("SELECT smtp_id, smtp_server, smtp_port, smtp_ssl, auth_user, auth_pass, mail_from FROM tblSmtpProfiles")

      '
      ' Initialize the command object
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      SmtpProfiles.Clear()

      SyncLock SyncLockMain
        Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

        '
        ' Process the resutls
        '
        While dtrResults.Read()
          Dim SmtpProfile As New SmtpProfile()

          With SmtpProfile
            .SmtpId = dtrResults("smtp_id")
            .SmtpServer = dtrResults("smtp_server")
            .SmtpPort = dtrResults("smtp_port")
            .SmtpSsl = IIf(dtrResults("smtp_ssl") = 1, True, False)
            .AuthUser = dtrResults("auth_user")
            .AuthPass = hs.DecryptString(dtrResults("auth_pass"), "&Cul8r#1")
            .SmtpFrom = dtrResults("mail_from")
          End With
          SmtpProfiles.Add(SmtpProfile)

        End While

        dtrResults.Close()
      End SyncLock

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "RefreshSmtpProfiles()")
    End Try

  End Sub

  ''' <summary>
  ''' Expands HomeSeer attributes contained within a string
  ''' </summary>
  ''' <param name="strMessageBody"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ExpandHomeSeerVariables(ByVal strMessageBody As String) As String

    Try
      '
      ' Call HomeSeer ReplaceVariables
      '
      strMessageBody = hs.ReplaceVariables(strMessageBody)

      Dim colMatches As MatchCollection

      '
      ' Expand HomeSeer device strings
      '
      colMatches = Regex.Matches(strMessageBody, "\$\$(?<ATTR>D.[ACR]):(?<ADDR>[^:]{1,32}):")
      If colMatches.Count > 0 Then
        For Each objMatch As Match In colMatches
          Dim strDeviceAttr As String = objMatch.Groups("ATTR").Value
          Dim strDeviceAddr As String = objMatch.Groups("ADDR").Value

          Dim strHSDeviceAttribute As String = GetHSDeviceAttribute(strDeviceAddr, strDeviceAttr)

          WriteMessage("DC = " & strDeviceAddr, MessageType.Debug)
          WriteMessage("Attr = " & strHSDeviceAttribute, MessageType.Debug)

          Dim strFind As String = String.Format("$${0}:{1}", strDeviceAttr, strDeviceAddr)
          WriteMessage("Find = " & strFind, MessageType.Debug)

          strMessageBody = Regex.Replace(strMessageBody, Regex.Escape(strFind) & "\b", strHSDeviceAttribute)
          WriteMessage("Line = " & strMessageBody, MessageType.Debug)
        Next
      End If

      '
      ' Expand HomeSeer variable strings
      '
      colMatches = Regex.Matches(strMessageBody, "\$\$GetVar:\[(?<VAR>.+?)\]")
      If colMatches.Count > 0 Then
        For Each objMatch As Match In colMatches
          Dim strVariableName As String = objMatch.Groups("VAR").Value

          WriteMessage("VAR = " & strVariableName, MessageType.Debug)

          Dim strFind As String = String.Format("\$\$GetVar:\[{0}\]", strVariableName)
          WriteMessage("Find = " & strFind, MessageType.Debug)

          Dim strVariableValue As String = " {???} "

          Try
            Dim objObject As Object = hs.GetVar(strVariableName)
            If TypeOf objObject Is String Then
              strVariableValue = CStr(objObject)
            End If
            objObject = Nothing

          Catch pEx As Exception

          End Try

          strMessageBody = Regex.Replace(strMessageBody, strFind, strVariableValue)
          WriteMessage("Line = " & strMessageBody, MessageType.Debug)
        Next
      End If

      'This message was sent on $date at $time. Math.Divide(114, 100)Math.Multiply(114, 2)Math.Add(114, 0)Math.Subtract(114, 1)
      Dim Operations As String() = {"Add", "Subtract", "Multiply", "Divide"}

      '
      ' Expand math variables
      '
      For Each strOperation As String In Operations

        Dim strRegexPattern As String = String.Format("\$\$Math\.{0}\((?<Expression>-?\d+,\s?-?\d+)\)", strOperation)
        colMatches = Regex.Matches(strMessageBody, strRegexPattern)

        If colMatches.Count > 0 Then
          For Each objMatch As Match In colMatches
            Dim strExpression As String = objMatch.Groups("Expression").Value

            Dim updatedOperation As String = "+"
            Select Case strOperation
              Case "Divide"
                updatedOperation = "/"
              Case "Multiply"
                updatedOperation = "*"
              Case "Add"
                updatedOperation = "+"
              Case "Subtract"
                updatedOperation = "-"
            End Select

            Dim strFind As String = String.Format("\$\$Math\.{0}\({1}\)", strOperation, strExpression)
            WriteMessage("Find = " & strFind, MessageType.Debug)

            Dim updatedExpression As String = strExpression.Replace(",", updatedOperation)
            WriteMessage("Expression = " & updatedExpression, MessageType.Debug)

            Dim strResult As String = CStr(New DataTable().Compute(updatedExpression, ""))
            WriteMessage("Result = " & strResult, MessageType.Debug)

            strMessageBody = Regex.Replace(strMessageBody, strFind, strResult)
            WriteMessage("Line = " & strMessageBody, MessageType.Debug)

          Next
        End If

      Next

      '
      ' Get HomeSeer log data
      '
      If Regex.IsMatch(strMessageBody, "\$\$HS\.Logs") = True Then
        strMessageBody = strMessageBody.Replace("$$HS.Logs", GetExistingHomeSeerLogData(10, "html"))
      End If

      '
      ' Replace misc
      '
      strMessageBody = strMessageBody.Replace("$$HS.Sunset", hs.Sunset())
      strMessageBody = strMessageBody.Replace("$$HS.Sunrise", hs.Sunrise())
      strMessageBody = strMessageBody.Replace("$$HS.SystemUptime", hs.SystemUpTime())

      If strMessageBody.Contains("$$HS.WANIP") Then
        Try
          hs.SetRemoteTimeout(10)
          strMessageBody = strMessageBody.Replace("$$HS.WANIP", hs.WANIP())
        Catch pEx As Exception

        End Try
      End If
      strMessageBody = strMessageBody.Replace("$$HS.LANIP", hs.LANIP())
      strMessageBody = strMessageBody.Replace("$$HS.GetLastRemoteIP", hs.GetLastRemoteIP())
      strMessageBody = strMessageBody.Replace("$$HS.HSMemoryUsed", hs.HSMemoryUsed())

      '
      ' Update the date and time
      '
      strMessageBody = strMessageBody.Replace("$date", DateTime.Now.ToString("MMMM dd, yyyy"))
      strMessageBody = strMessageBody.Replace("$time", DateTime.Now.ToString("hh:mm tt"))

      '
      ' Remove leading and trailing spaces
      '
      strMessageBody = strMessageBody.Trim

      '
      ' Prevent characters that cause message from being displayed
      '
      'Dim strPattern As String = "[^-A-Za-z0-9~.,:;'|<>\\\/\s!@#$%&*()_+={}\[\]]"
      'strMessageBody = System.Text.RegularExpressions.Regex.Replace(strMessageBody, strPattern, " ")

      Return strMessageBody

    Catch pEx As Exception
      Call ProcessError(pEx, "ExpandHomeSeerVariables")
      Return strMessageBody
    End Try

  End Function

  ''' <summary>
  ''' This function is called to get the HomeSeer log data we didn't get
  ''' </summary>
  ''' <param name="iLines"></param>
  ''' <param name="strFormat"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetExistingHomeSeerLogData(ByVal iLines As Integer, ByVal strFormat As String) As String

    Dim HSLog As New StringBuilder()

    Try
      '
      ' Get the pending HomeSeer log entries
      '
      Dim strLog As String = hs.LogGet
      Dim lines() As String = strLog.Split(Chr(10))

      Dim iStart As Integer = 0
      Dim iEnd As Integer = 0

      iEnd = lines.GetLength(0) - 1
      iStart = iEnd - iLines

      If iStart < 0 Then iStart = 0

      For i As Integer = iStart To iEnd

        Dim colMatches As MatchCollection = Regex.Matches(lines(i), "^(?<date>.+)~!~(?<type>.+)~!~(?<msg>.+)$")

        If colMatches.Count > 0 Then

          Dim EventDate As String = colMatches.Item(0).Groups("date").Value
          Dim EventType As String = colMatches.Item(0).Groups("type").Value
          Dim EventMsg As String = colMatches.Item(0).Groups("msg").Value

          ' Dim LogLine As String = String.Format("<table class=""hslog""><tr><td>{0, -10}</td><td>{1, -10}</td><td>{2, -10}</td></tr></table>", EventDate, EventType, EventMsg)

          HSLog.Append(lines(i))

        End If

      Next

    Catch pEx As Exception
      '
      ' Ignore any erorrs here 
      '
    End Try

    Return HSLog.ToString

  End Function

  ''' <summary>
  ''' Get a HomeSeer device attribute
  ''' </summary>
  ''' <param name="DeviceAddress"></param>
  ''' <param name="DeviceAttribute"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetHSDeviceAttribute(ByVal DeviceAddress As String, ByVal DeviceAttribute As String) As String

    Dim dv As Scheduler.Classes.DeviceClass
    Dim dvRef As Long = 0

    Dim DeviceAttrValue As String = " {???} "

    Try

      If DeviceAttribute.EndsWith("C") Then
        '
        ' Locate device by code
        '
        dv = hspi_devices.LocateDeviceByCode(DeviceAddress)

      ElseIf DeviceAttribute.EndsWith("A") Then
        '
        ' Locate device by address
        '
        dv = hspi_devices.LocateDeviceByAddr(DeviceAddress)

      ElseIf DeviceAttribute.EndsWith("R") Then
        '
        ' Locate device by ref
        '
        dv = hspi_devices.LocateDeviceByRef(DeviceAddress)

      Else
        '
        ' No idea what the replacment variable is
        '
        dv = Nothing

      End If

      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        If DeviceAttribute.StartsWith("DN") Then
          DeviceAttrValue = dv.Name(Nothing) & ""
        ElseIf DeviceAttribute.StartsWith("DL") Then
          DeviceAttrValue = dv.Location(Nothing)
        ElseIf DeviceAttribute.StartsWith("Dl") Then
          DeviceAttrValue = dv.Location2(Nothing)
        ElseIf DeviceAttribute.StartsWith("DT") Then
          DeviceAttrValue = dv.Device_Type_String(Nothing)
        ElseIf DeviceAttribute.StartsWith("Dt") Then
          DeviceAttrValue = dv.Device_Type_String(Nothing)
        ElseIf DeviceAttribute.StartsWith("DC") Then
          DeviceAttrValue = dv.Last_Change(Nothing).ToString
        End If
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "GetHSDeviceAttribute")
    End Try

    Return DeviceAttrValue

  End Function

  Public Enum FileAction
    Keep = 0
    Delete = 1
    Move = 2
  End Enum

  Public Enum DeliveryResult
    Success = 0
    Deferred = 1
    Failure = 2
  End Enum

#End Region

#Region "UltraSMTP3 Actions/Triggers/Conditions"

#Region "Trigger Proerties"

  ''' <summary>
  ''' Defines the valid triggers for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetTriggers()
    Dim o As Object = Nothing
    If triggers.Count = 0 Then
      triggers.Add(o, "Email Delivery Status")           ' 1
    End If
  End Sub

  ''' <summary>
  ''' Lets HomeSeer know our plug-in has triggers
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasTriggers() As Boolean
    Get
      SetTriggers()
      Return IIf(triggers.Count > 0, True, False)
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerCount() As Integer
    SetTriggers()
    Return triggers.Count
  End Function

  ''' <summary>
  ''' Returns the subtrigger count
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer
    Get
      Dim trigger As trigger
      If ValidTrig(TriggerNumber) Then
        trigger = triggers(TriggerNumber - 1)
        If Not (trigger Is Nothing) Then
          Return 0
        Else
          Return 0
        End If
      Else
        Return 0
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String
    Get
      If Not ValidTrig(TriggerNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, triggers.Keys(TriggerNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the subtrigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <param name="SubTriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String
    Get
      Dim trigger As trigger
      If ValidSubTrig(TriggerNumber, SubTriggerNumber) Then
        Return ""
      Else
        Return ""
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is valid
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidTrig(ByVal TrigIn As Integer) As Boolean
    SetTriggers()
    If TrigIn > 0 AndAlso TrigIn <= triggers.Count Then
      Return True
    End If
    Return False
  End Function

  ''' <summary>
  ''' Determines if the trigger is a valid subtrigger
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <param name="SubTrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidSubTrig(ByVal TrigIn As Integer, ByVal SubTrigIn As Integer) As Boolean
    Return False
  End Function

  ''' <summary>
  ''' Tell HomeSeer which triggers have conditions
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean
    Get
      Select Case TriggerNumber
        Case 0
          Return True   ' Render trigger as IF / AND IF
        Case Else
          Return False  ' Render trigger as IF / OR IF
      End Select
    End Get
  End Property

  ''' <summary>
  ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  
  ''' Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Set(ByVal value As Boolean)

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      ' TriggerCondition(sKey) = value

    End Set
    Get

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Return False

    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is a condition
  ''' </summary>
  ''' <param name="sKey"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property TriggerCondition(sKey As String) As Boolean
    Get

      If conditions.ContainsKey(sKey) = True Then
        Return conditions(sKey)
      Else
        Return False
      End If

    End Get
    Set(value As Boolean)

      If conditions.ContainsKey(sKey) = False Then
        conditions.Add(sKey, value)
      Else
        conditions(sKey) = value
      End If

    End Set
  End Property

  ''' <summary>
  ''' Called when HomeSeer wants to check if a condition is true
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Return False
  End Function

#End Region

#Region "Trigger Interface"

  ''' <summary>
  ''' Builds the Trigger UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = TrigInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    Else 'new event, so clean out the trigger object
      trigger = New trigger
    End If

    Select Case TrigInfo.TANumber
      Case SMTPTriggers.EmailDeliveryStatus
        Dim triggerName As String = GetEnumName(SMTPTriggers.EmailDeliveryStatus)

        Dim ActionSelected As String = trigger.Item("DeliveryStatus")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "DeliveryStatus", UID, sUnique)

        Dim jqDSN As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqDSN.autoPostBack = True

        jqDSN.AddItem("(Select Delivery Status)", "", (ActionSelected = ""))
        Dim Actions As String() = {"Success", "Deferral", "Failure"}
        For Each strAction As String In Actions
          Dim strOptionValue As String = strAction
          Dim strOptionName As String = strOptionValue
          jqDSN.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqDSN.Build)

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Process changes to the trigger from the HomeSeer events page
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection, _
                                       ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As String = TrigInfo.UID.ToString
    Dim TANumber As Integer = TrigInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = TrigInfo.DataIn
    Ret.TrigActInfo = TrigInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    ' DeSerializeObject
    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If
    trigger.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case SMTPTriggers.EmailDeliveryStatus
          Dim triggerName As String = GetEnumName(SMTPTriggers.EmailDeliveryStatus)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, triggerName & "DeliveryStatus_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("DeliveryStatus") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(trigger, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Trigger not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Trigger UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret

  End Function

  ''' <summary>
  ''' Determines if a trigger is properly configured
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Get
      Dim Configured As Boolean = True
      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Select Case TrigInfo.TANumber
        Case SMTPTriggers.EmailDeliveryStatus
          If trigger.Item("DeliveryStatus") = "" Then Configured = False

      End Select

      Return Configured
    End Get
  End Property

  ''' <summary>
  ''' Formats the trigger for display
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim stb As New StringBuilder

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Select Case TrigInfo.TANumber
      Case SMTPTriggers.EmailDeliveryStatus
        If trigger.uid <= 0 Then
          stb.Append("Trigger has not been properly configured.")
        Else
          Dim strTriggerName As String = GetEnumDescription(SMTPTriggers.EmailDeliveryStatus)
          Dim strDeliveryStatus As String = trigger.Item("DeliveryStatus")

          stb.AppendFormat("{0} is <font class='event_Txt_Option'>{1}</font>", strTriggerName, strDeliveryStatus)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Checks to see if trigger should fire
  ''' </summary>
  ''' <param name="Plug_Name"></param>
  ''' <param name="TrigID"></param>
  ''' <param name="SubTrig"></param>
  ''' <param name="strTrigger"></param>
  ''' <remarks></remarks>
  Private Sub CheckTrigger(Plug_Name As String, TrigID As Integer, SubTrig As Integer, strTrigger As String)

    Try
      '
      ' Check HomeSeer Triggers
      '
      If Plug_Name.Contains(":") = False Then Plug_Name &= ":"
      Dim TrigsToCheck() As IAllRemoteAPI.strTrigActInfo = callback.TriggerMatches(Plug_Name, TrigID, SubTrig)

      Try

        For Each TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo In TrigsToCheck
          Dim UID As String = TrigInfo.UID.ToString

          If Not (TrigInfo.DataIn Is Nothing) Then

            Dim trigger As New trigger
            DeSerializeObject(TrigInfo.DataIn, trigger)

            Select Case TrigID

              Case SMTPTriggers.EmailDeliveryStatus
                Dim strTriggerName As String = GetEnumDescription(SMTPTriggers.EmailDeliveryStatus)
                Dim strDeliveryStatus As String = trigger.Item("DeliveryStatus")

                Dim strTriggerCheck As String = String.Format("{0},{1}", strTriggerName, strDeliveryStatus)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

            End Select

          End If

        Next

      Catch pEx As Exception

      End Try

    Catch pEx As Exception

    End Try

  End Sub

#End Region

#Region "Action Properties"

  ''' <summary>
  ''' Defines the valid actions for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetActions()
    Dim o As Object = Nothing
    If actions.Count = 0 Then
      actions.Add(o, "Send an Email")           ' 1
    End If
  End Sub

  ''' <summary>
  ''' Returns the action count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ActionCount() As Integer
    SetActions()
    Return actions.Count
  End Function

  ''' <summary>
  ''' Returns the action name
  ''' </summary>
  ''' <param name="ActionNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String
    Get
      If Not ValidAction(ActionNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, actions.Keys(ActionNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if an action is valid
  ''' </summary>
  ''' <param name="ActionIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidAction(ByVal ActionIn As Integer) As Boolean
    SetActions()
    If ActionIn > 0 AndAlso ActionIn <= actions.Count Then
      Return True
    End If
    Return False
  End Function

#End Region

#Region "Action Interface"

  ''' <summary>
  ''' Builds the Action UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks>This function is called from the HomeSeer event page when an event is in edit mode.</remarks>
  Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = ActInfo.UID.ToString
    Dim stb As New StringBuilder
    Dim stbSlider1 As New StringBuilder
    Dim stbSlider2 As New StringBuilder

    Try

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      End If

      Select Case ActInfo.TANumber
        Case SMTPActions.SendEmail

          Dim sbTemplate As New StringBuilder

          Dim strDefaultTo As String = hs.GetINISetting("Settings", "gSMTPTo", "")
          Dim strDefaultSubject As String = hs.GetINISetting("Settings", "gSMTPDefSubj", "UltraSMTP3 Notification")
          Dim strDefaultBody As String = hs.GetINISetting("Settings", "gSMTPDefMess", "")

          sbTemplate.Append(strDefaultBody)

          Dim actionName As String = GetEnumName(SMTPActions.SendEmail)

          '
          ' Start Recipient
          '
          Dim ActionSelected As String = IIf(action.Item("EmailRecipient").Length = 0, strDefaultTo, action.Item("EmailRecipient"))
          Dim actionId As String = String.Format("{0}{1}_{2}_{3}", actionName, "EmailRecipient", UID, sUnique)

          Dim jqEmailRecipient As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 50, False)
          stb.Append("<table border='1'>")
          stb.Append("<tr>")
          stb.Append(" <td style='text-align:right'>")
          stb.Append("To:")
          stb.Append(New clsJQuery.jqToolTip("Enter the Email address of the intended recipient.").build)
          stb.Append(" </td>")
          stb.Append(" <td colspan='2'>")
          stb.Append(jqEmailRecipient.Build)
          stb.Append(" </td>")
          stb.Append("</tr>")

          '
          ' Start Subject
          '
          ActionSelected = IIf(action.Item("EmailSubject").Length = 0, strDefaultSubject, action.Item("EmailSubject"))
          actionId = String.Format("{0}{1}_{2}_{3}", actionName, "EmailSubject", UID, sUnique)

          Dim jqEmailSubject As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 100, False)

          stb.Append("<tr>")
          stb.Append(" <td style='text-align:right'>")
          stb.Append("Subject:")
          stb.Append(New clsJQuery.jqToolTip("Enter the subject of the message.").build)
          stb.Append(" </td>")
          stb.Append(" <td colspan='2'>")
          stb.Append(jqEmailSubject.Build)
          stb.Append(" </td>")
          stb.Append("</tr>")

          '
          ' selReplacmentVariable
          '
          Dim List1 As New SortedList
          List1.Add("$date", "Replacement is the current date in long format, e.g.: Jan 1, 1970")
          List1.Add("$time", "Replacement is the current time in 12 hour format, e.g. 2:00 PM")
          List1.Add("$$GLOBALVAR:name:", "Replacement is the value of the global variable indicated by name.")
          List1.Add("$$COUNTER:name:", "Replacement is the value of the counter indicated by name.")
          List1.Add("$$TIMER:name:", "Replacement is the value of the timer indicated by name.")
          List1.Add("$$HS.Logs", "Replacement is the contents of the last 10 HomeSeer log lines.")
          List1.Add("$$HS.WANIP", "Replacement is the name and IP address of your HomeSeer WAN interface.")
          List1.Add("$$HS.LANIP", "Replacement is the IP address of your HomeSeer local LAN interface.")
          List1.Add("$$HS.GetLastRemoteIP", "Replacement is the IP address of the last system that accessed your HomeSeer.")
          List1.Add("$$HS.HSMemoryUsed", "Replacement is the value of the memory used by HomeSeer.")
          List1.Add("$$HS.SystemUptime", "Replacement is the system uptime since last restart.")
          List1.Add("$$HS.Sunrise", "Replacement is the time of sunrise.")
          List1.Add("$$HS.Sunset", "Replacement is the time of sunset.")

          Dim selReplacmentVariable As New clsJQuery.jqDropList("selReplacmentVariable", Pagename, False)
          selReplacmentVariable.id = "selReplacmentVariable"
          selReplacmentVariable.AddItem("Select the HomeSeer Replacment Variable", "", False)
          For Each key As String In List1.Keys
            selReplacmentVariable.AddItem(List1(key), key, False)
          Next

          '
          ' selDeviceList
          '
          Dim DeviceList As SortedList = GetDeviceList()
          Dim selDeviceList As New clsJQuery.jqDropList("selDeviceRef", Pagename, False)
          selDeviceList.id = "selDeviceRef"

          selDeviceList.AddItem("Select the HomeSeer Device", "", False)
          For Each key As String In DeviceList.Keys
            Dim value As String = DeviceList(key)
            selDeviceList.AddItem(key, value, False)
          Next

          '
          ' selDeviceProperty
          '
          Dim List2 As New SortedList
          List2.Add("$$DVR:ref:", "Replacement is the VALUE of the selected device")
          List2.Add("$$DSR:ref:", "Replacement is the STATUS of the selected device")
          List2.Add("$$DTR:ref:", "Replacement is the STRING of the selected device")
          List2.Add("$$DNR:ref:", "Replacement is the NAME of the selected device")
          List2.Add("$$DLR:ref:", "Replacement is the LOCATION1 of the selected device")
          List2.Add("$$DlR:ref:", "Replacement is the LOCATION2 of the selected device")
          List2.Add("$$DtR:ref:", "Replacement is the TYPE of the selected device")
          List2.Add("$$DCR:ref:", "Replacement is the LASTCHANGE of the selected device")

          Dim selDeviceProperty As New clsJQuery.jqDropList("selDeviceProperty", Pagename, False)
          selDeviceProperty.id = "selDeviceProperty"

          selDeviceProperty.AddItem("Select the Device Property Replacment Variable", "", True)
          For Each key As String In List2.Keys
            selDeviceProperty.AddItem(List2(key), key, False)
          Next

          stbSlider1.AppendLine("<div>")
          stbSlider1.AppendLine("<ol>")
          stbSlider1.AppendLine("<li>Place your cursor into the message text area where you want to insert a replacment variable.</li>")
          stbSlider1.AppendLine("<li>Select the HomeSeer Device from the dropdown list.</li>")
          stbSlider1.AppendLine("<li>Select the Replacment Variable from the dropdown list.</li>")
          stbSlider1.AppendLine("<li>Click the ""Insert Replacment Variable Into Message"" button.</li>")
          stbSlider1.AppendLine("</ol>")

          stbSlider1.AppendLine("<table cellspacing='1' width='100%'>")
          stbSlider1.AppendLine(" <tr>")
          stbSlider1.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selDeviceList.Build, vbCrLf)
          stbSlider1.AppendLine(" </tr>")
          stbSlider1.AppendLine(" <tr>")
          stbSlider1.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selDeviceProperty.Build, vbCrLf)
          stbSlider1.AppendLine(" </tr>")
          stbSlider1.AppendLine(" <tr>")
          stbSlider1.AppendLine("<td><button id=""btnInsert1"">Insert Replacment Variable Into Message</button></td>")
          stbSlider1.AppendLine(" </tr>")
          stbSlider1.AppendLine("</table>")
          stbSlider1.AppendLine("</div>")

          stbSlider2.AppendLine("<div>")
          stbSlider2.AppendLine("<ol>")
          stbSlider2.AppendLine("<li>Place your cursor into the message text area where you want to insert a replacment variable.</li>")
          stbSlider2.AppendLine("<li>Select the Replacment Variable from the dropdown list.</li>")
          stbSlider2.AppendLine("<li>Click the ""Insert Replacment Variable Into Message"" button.</li>")
          stbSlider2.AppendLine("</ol>")

          stbSlider2.AppendLine("<table cellspacing='1' width='100%'>")
          stbSlider2.AppendLine(" <tr>")
          stbSlider2.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selReplacmentVariable.Build, vbCrLf)
          stbSlider2.AppendLine(" </tr>")
          stbSlider2.AppendLine(" <tr>")
          stbSlider2.AppendLine("<td><button id=""btnInsert2"">Insert Replacment Variable Into Message</button></td>")
          stbSlider2.AppendLine(" </tr>")
          stbSlider2.AppendLine("</table>")
          stbSlider2.AppendLine("</div>")

          Dim jqSlidingTab1 As New clsJQuery.jqSlidingTab("myslide1_ID", Pagename, False)
          jqSlidingTab1.initiallyOpen = False
          jqSlidingTab1.tab.AddContent(stbSlider1.ToString)
          jqSlidingTab1.tab.name = "myslide1_name"
          jqSlidingTab1.tab.tabName.Unselected = "Show Device Replacment Variables"
          jqSlidingTab1.tab.tabName.Selected = "Hide Device Replacment Variables"

          Dim jqSlidingTab2 As New clsJQuery.jqSlidingTab("myslide2_ID", Pagename, False)
          jqSlidingTab2.initiallyOpen = False
          jqSlidingTab2.tab.AddContent(stbSlider2.ToString)
          jqSlidingTab2.tab.name = "myslide2_name"
          jqSlidingTab2.tab.tabName.Unselected = "Show HomeSeer Replacment Variables"
          jqSlidingTab2.tab.tabName.Selected = "Hide HomeSeer Replacment Variables"

          '
          ' Start Body
          '
          ActionSelected = IIf(action.Item("EmailBody").Length = 0, sbTemplate.ToString, action.Item("EmailBody"))
          actionId = String.Format("{0}{1}_{2}_{3}", actionName, "EmailBody", UID, sUnique)

          stb.AppendLine("<tr>")
          stb.AppendLine(" <td style='text-align:right'>")
          stb.AppendLine("Message:")
          stb.AppendLine(New clsJQuery.jqToolTip("Enter the body of the message.  See the helpfile for a list of supported replacment variables.").build)
          stb.AppendLine(" </td>")
          stb.AppendLine(" <td>")
          stb.AppendLine(jqSlidingTab1.Build())
          stb.AppendLine(jqSlidingTab2.Build())
          stb.AppendFormat("<textarea style='display: table-cell;' cols='75' rows='25' id='txtMessageBody' class='txtDropTarget' name='{0}'>{1}</textarea>", actionId, ActionSelected)
          stb.AppendLine(" </td>")
          stb.Append("</tr>")

          '
          ' Start Attachment
          '
          ActionSelected = IIf(action.Item("EmailAttachment").Length = 0, "", action.Item("EmailAttachment"))
          actionId = String.Format("{0}{1}_{2}_{3}", actionName, "EmailAttachment", UID, sUnique)

          'Dim fs As New clsJQuery.jqLocalFileSelector(actionId, Pagename, True)
          'fs.label = "HomeSeer File System"
          'fs.path = hs.GetAppPath
          'fs.AddExtension("*.*")

          Dim jqEmailAttachment As New clsJQuery.jqTextBox(actionId, "text", ActionSelected, Pagename, 75, True)

          stb.Append("<tr>")
          stb.Append(" <td style='text-align:right'>")
          stb.Append("Attachment&nbsp;Path:")
          stb.Append(New clsJQuery.jqToolTip("Enter the full path to the attachment or directory/extension wildcard.  See the helpfile for a list of supported extension wildcard characters.").build)
          stb.Append(" </td>")
          stb.Append(" <td colspan='2'>")
          'stb.Append(fs.Build)
          stb.Append(jqEmailAttachment.Build)
          stb.Append(" </td>")
          stb.Append("</tr>")

          '
          ' Start Submit Button
          '
          Dim jqButton As New clsJQuery.jqButton("btnSave", "Save Message", Pagename, True)

          stb.Append("<tr>")
          stb.Append(" <td>")
          stb.Append(" </td>")
          stb.Append(" <td colspan='2'>")
          stb.Append(jqButton.Build)
          stb.Append(" </td>")
          stb.Append("</tr>")
          stb.Append("</table>")

          stb.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrasmtp3/js/hspi_ultrasmtp3_tools.js""></script>")

      End Select

    Catch pEx As Exception
      WriteMessage(pEx.Message, MessageType.Error)
    End Try

    Return stb.ToString

  End Function

  ''' <summary>
  ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection, _
                                      ByVal ActInfo As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As Integer = ActInfo.UID
    Dim TANumber As Integer = ActInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = ActInfo.DataIn
    Ret.TrigActInfo = ActInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    '
    ' DeSerializeObject
    '
    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If
    action.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case SMTPActions.SendEmail
          Dim actionName As String = GetEnumName(SMTPActions.SendEmail)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "EmailRecipient_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("EmailRecipient") = ActionValue

              Case InStr(sKey, actionName & "EmailSubject_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("EmailSubject") = ActionValue

              Case InStr(sKey, actionName & "EmailBody_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("EmailBody") = ActionValue

              Case InStr(sKey, actionName & "EmailAttachment_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("EmailAttachment") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(action, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Action not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Action UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret
  End Function

  ''' <summary>
  ''' Determines if our action is proplery configured
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return TRUE if the given action is configured properly</returns>
  ''' <remarks>There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.</remarks>
  Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim Configured As Boolean = True
    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case SMTPActions.SendEmail
        If action.Item("EmailRecipient") = "" Then Configured = False
        If action.Item("EmailSubject") = "" Then Configured = False
        If action.Item("EmailBody") = "" Then Configured = False
        'If action.Item("EmailAttachment") = "" Then Configured = False

    End Select

    Return Configured

  End Function

  ''' <summary>
  ''' After the action has been configured, this function is called in your plugin to display the configured action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return text that describes the given action.</returns>
  ''' <remarks></remarks>
  Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String
    Dim stb As New StringBuilder

    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case SMTPActions.SendEmail
        If action.uid <= 0 Then
          stb.Append("Action has not been properly configured.")
        Else
          Dim strActionName = GetEnumDescription(SMTPActions.SendEmail)

          Dim strEmailRecipient As String = action.Item("EmailRecipient")
          Dim EmailSubject As String = action.Item("EmailSubject")
          Dim strEmailBody As String = action.Item("EmailBody")
          Dim strEmailAttachment As String = action.Item("EmailAttachment")

          If strEmailAttachment.Length = 0 Then
            strEmailAttachment = "No attachment specified"
          End If

          stb.AppendFormat("{0} To <font class='event_Txt_Selection'>{1}</font> " & _
                            "with the subject <font class='event_Txt_Selection'>{2}</font><br>" & _
                            "Attachment Path: <font class='event_Txt_Selection'>{3}</font>",
                            strActionName, _
                            strEmailRecipient, _
                            EmailSubject, _
                            strEmailAttachment)

        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Handles the HomeSeer Event Action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = ActInfo.UID.ToString

    Try

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      Else
        Return False
      End If

      Select Case ActInfo.TANumber
        Case SMTPActions.SendEmail
          Dim strEmailRecipient As String = action.Item("EmailRecipient")
          Dim EmailSubject As String = action.Item("EmailSubject")
          Dim strEmailBody As String = action.Item("EmailBody")
          Dim strEmailAttachment As String = action.Item("EmailAttachment")

          Dim MailAttachments() As String = Nothing
          If strEmailAttachment.Length > 0 Then

            Dim regexPattern As String = "^(?<path>(.*\\))(?<filename>([^\\]*))"
            Dim DirPath As String = Regex.Match(strEmailAttachment, regexPattern).Groups("path").ToString()
            Dim FileName As String = Regex.Match(strEmailAttachment, regexPattern).Groups("filename").ToString()

            If FileName.Length = 0 And Directory.Exists(DirPath) = True Then
              '
              ' We are dealing with a directory
              '
            ElseIf File.Exists(strEmailAttachment) = True Then
              '
              ' We are dealing with a single attachment
              '
              ReDim MailAttachments(0)
              MailAttachments(0) = strEmailAttachment
            ElseIf FileName.Length > 0 And Directory.Exists(DirPath) Then
              '
              ' Determine if the filename contains wildcard characters
              '
              Dim DirectoryInfo As New IO.DirectoryInfo(DirPath)
              Dim DirectoryFiles As IO.FileInfo() = DirectoryInfo.GetFiles(FileName)

              Dim iSizeTotal As Long = 0
              For Each FileInfo As IO.FileInfo In DirectoryFiles
                Dim iFileIndex As Integer = 0
                iSizeTotal += Math.Round(FileInfo.Length / 1024)
                If Not MailAttachments Is Nothing Then
                  iFileIndex = MailAttachments.Length
                End If

                If iSizeTotal >= (gMaxAttachmentSize) Then
                  WriteMessage(String.Format("Unable to attach {0} because attachment would exceeds {1} KB.", FileInfo.Name, gMaxAttachmentSize.ToString), MessageType.Warning)
                Else

                  WriteMessage(String.Format("Attaching {0}, total attachment size is {1} KB.", FileInfo.Name, iSizeTotal), MessageType.Debug)
                  ReDim Preserve MailAttachments(iFileIndex)
                  MailAttachments(iFileIndex) = FileInfo.FullName
                End If

              Next

            End If

          End If

          '
          ' Send the e-mail message
          '
          SendMail(strEmailRecipient, EmailSubject, strEmailBody, MailAttachments, 0)

      End Select

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      hs.WriteLog(IFACE_NAME, "Error executing action: " & pEx.Message)
    End Try

    Return True

  End Function

#End Region

#End Region

End Module

Public Enum SMTPTriggers
  <Description("Email Delivery Status")> _
  EmailDeliveryStatus = 1
End Enum

Public Enum SMTPActions
  <Description("Send and Email")> _
  SendEmail = 1
End Enum