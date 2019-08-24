Imports System.Data.Common
Imports System.Threading
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Data.SQLite

Module Database

  Public DBConnectionMain As SQLite.SQLiteConnection   ' Our main database connection
  Public DBConnectionTemp As SQLite.SQLiteConnection  ' Our temp database connection

  Public gDBInsertSuccess As ULong = 0            ' Tracks DB insert success
  Public gDBInsertFailure As ULong = 0            ' Tracks DB insert success

  Public bDBInitialized As Boolean = False        ' Indicates if database successfully initialized

  Public SyncLockMain As New Object
  Public SyncLockTemp As New Object

#Region "Database Initilization"

  ''' <summary>
  ''' Initializes the database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitializeMainDatabase() As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered InitializeMainDatabase() function.", MessageType.Debug)

    Try
      '
      ' Close database if it's open
      '
      If Not DBConnectionMain Is Nothing Then
        If CloseDBConn(DBConnectionMain) = False Then
          Throw New Exception("An existing database connection could not be closed.")
        End If
      End If

      '
      ' Create the database directory if it does not exist
      '
      Dim databaseDir As String = FixPath(String.Format("{0}\Data\{1}\", hs.GetAppPath(), IFACE_NAME.ToLower))
      If Directory.Exists(databaseDir) = False Then
        Directory.CreateDirectory(databaseDir)
      End If

      '
      ' Determine the database filename
      '
      Dim strDataSource As String = FixPath(String.Format("{0}\Data\{1}\{1}.db3", hs.GetAppPath(), IFACE_NAME.ToLower))

      '
      ' Determine the database provider factory and connection string
      '
      Dim strConnectionString As String = String.Format("Data Source={0}; Version=3;", strDataSource)

      '
      ' Attempt to open the database connection
      '
      bSuccess = OpenDBConn(DBConnectionMain, strConnectionString)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      bSuccess = False
      Call ProcessError(pEx, "InitializeDatabase()")
    End Try

    Return bSuccess

  End Function

  '------------------------------------------------------------------------------------
  'Purpose: Initializes the temporary database
  'Inputs:  None
  'Outputs: True or False indicating if database was initialized
  '------------------------------------------------------------------------------------
  Public Function InitializeTempDatabase() As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    strMessage = "Entered InitializeChannelDatabase() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Close database if it's open
      '
      If Not DBConnectionTemp Is Nothing Then
        If CloseDBConn(DBConnectionTemp) = False Then
          Throw New Exception("An existing database connection could not be closed.")
        End If
      End If

      '
      ' Determine the database filename
      '
      Dim dtNow As DateTime = DateTime.Now
      Dim iHour As Integer = dtNow.Hour
      Dim strDBDate As String = iHour.ToString.PadLeft(2, "0")
      Dim strDataSource As String = FixPath(String.Format("{0}\Data\{1}\{1}_{2}.db3", hs.GetAppPath(), IFACE_NAME.ToLower, strDBDate))

      '
      ' Determine the database provider factory and connection string
      '
      Dim strConnectionString As String = String.Format("Data Source={0}; Version=3; Journal Mode=Off;", strDataSource)

      '
      ' Attempt to open the database connection
      '
      bSuccess = OpenDBConn(DBConnectionTemp, strConnectionString)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      bSuccess = False
      Call ProcessError(pEx, "InitializeTempDatabase()")
    End Try

    Return bSuccess

  End Function

  ''' <summary>
  ''' Opens a connection to the database
  ''' </summary>
  ''' <param name="DBConnection"></param>
  ''' <param name="strConnectionString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function OpenDBConn(ByRef DBConnection As SQLite.SQLiteConnection, _
                              ByVal strConnectionString As String) As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered OpenDBConn() function.", MessageType.Debug)

    Try
      DBConnection = New SQLite.SQLiteConnection()
      DBConnection.ConnectionString = strConnectionString
      DBConnection.Open()

      '
      ' Run database vacuum
      '
      WriteMessage("Running SQLite database vacuum.", MessageType.Debug)
      Using MyDbCommand As DbCommand = DBConnection.CreateCommand()

        MyDbCommand.Connection = DBConnection
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = "VACUUM"
        MyDbCommand.ExecuteNonQuery()

        MyDbCommand.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "OpenDBConn()")
    End Try

    '
    ' Determine database connection status
    '
    bSuccess = DBConnection.State = ConnectionState.Open

    '
    ' Record connection state to HomeSeer log
    '
    If bSuccess = True Then
      strMessage = "Database initialization complete."
      Call WriteMessage(strMessage, MessageType.Debug)
    Else
      strMessage = "Database initialization failed using [" & strConnectionString & "]."
      Call WriteMessage(strMessage, MessageType.Debug)
    End If

    Return bSuccess

  End Function

  ''' <summary>
  ''' Closes database connection
  ''' </summary>
  ''' <param name="DBConnection"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CloseDBConn(ByRef DBConnection As SQLite.SQLiteConnection) As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered CloseDBConn() function.", MessageType.Debug)

    Try
      '
      ' Attempt to the database
      '
      If DBConnection.State <> ConnectionState.Closed Then
        DBConnection.Close()
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CloseDBConn()")
    End Try

    '
    ' Determine database connection status
    '
    bSuccess = DBConnection.State = ConnectionState.Closed

    '
    ' Record connection state to HomeSeer log
    '
    If bSuccess = True Then
      strMessage = "Database connection closed successfuly."
      Call WriteMessage(strMessage, MessageType.Debug)
    Else
      strMessage = "Unable to close database; Try restarting HomeSeer."
      Call WriteMessage(strMessage, MessageType.Debug)
    End If

    Return bSuccess

  End Function

  ''' <summary>
  ''' Checks to ensure a table exists
  ''' </summary>
  ''' <param name="strTableName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CheckDatabaseTable(ByVal strTableName As String) As Boolean

    Dim strMessage As String = ""
    Dim bSuccess As Boolean = False

    Try
      '
      ' Build SQL delete statement
      '
      If Regex.IsMatch(strTableName, "tblSmtpProfiles") = True Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          '
          ' Create the table
          '
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Debug)

          Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

            Dim sqlQueue As New Queue

            sqlQueue.Enqueue("CREATE TABLE tblSmtpProfiles(" _
                            & "smtp_id INTEGER PRIMARY KEY," _
                            & "smtp_server varchar(255) NOT NULL," _
                            & "smtp_port INTEGER NOT NULL," _
                            & "smtp_ssl INTEGER NOT NULL," _
                            & "auth_user varchar(255)," _
                            & "auth_pass varchar(255)," _
                            & "mail_from varchar(255)" _
                          & ")")

            sqlQueue.Enqueue("CREATE UNIQUE INDEX idxUNIQUE ON tblSmtpProfiles(smtp_server, smtp_port)")

            While sqlQueue.Count > 0
              Dim strSQL As String = sqlQueue.Dequeue

              dbcmd.Connection = DBConnectionMain
              dbcmd.CommandType = CommandType.Text
              dbcmd.CommandText = strSQL

              Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
              If iRecordsAffected <> 1 Then
                'Throw New Exception("Database schemea update failed due to error.")
              End If

            End While

            dbcmd.Dispose()
          End Using

        End If

      ElseIf String.Compare(strTableName, "tblSmtpQueue", True) = 0 Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Informational)

          Dim dbcmd As DbCommand = DBConnectionMain.CreateCommand()

          Dim sqlQueue As New Queue

          sqlQueue.Enqueue("CREATE TABLE tblSmtpQueue(" _
                          & "queue_id INTEGER PRIMARY KEY," _
                          & "queue_ts INTEGER NOT NULL," _
                          & "delivered_ts INTEGER NOT NULL," _
                          & "attempts INTEGER NOT NULL," _
                          & "smtp_id INTEGER NOT NULL," _
                          & "last_status varchar(255) NOT NULL," _
                          & "last_result varchar(255)," _
                          & "mail_to varchar(255) NOT NULL," _
                          & "mail_subject varchar(255)" _
                      & ")")

          sqlQueue.Enqueue("CREATE INDEX idxQUEUEID    ON tblSmtpQueue(queue_id)")
          sqlQueue.Enqueue("CREATE INDEX idxQUEUETS    ON tblSmtpQueue(queue_ts)")

          While sqlQueue.Count > 0
            Dim strSQL As String = sqlQueue.Dequeue

            dbcmd.Connection = DBConnectionMain
            dbcmd.CommandType = CommandType.Text
            dbcmd.CommandText = strSQL

            Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
            If iRecordsAffected <> 1 Then
              'Throw New Exception("Create table failed due to error.")
            End If

            dbcmd.Dispose()
          End While
        End If

      ElseIf String.Compare(strTableName, "tblSmtpLog", True) = 0 Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Informational)

          Dim dbcmd As DbCommand = DBConnectionMain.CreateCommand()

          Dim sqlQueue As New Queue

          sqlQueue.Enqueue("CREATE TABLE tblSmtpLog(" _
                          & "log_id INTEGER PRIMARY KEY," _
                          & "log_ts INTEGER NOT NULL," _
                          & "queue_id INTEGER NOT NULL," _
                          & "smtp_id INTEGER NOT NULL," _
                          & "smtp_server varchar(255) NOT NULL," _
                          & "smtp_status varchar(255) NOT NULL," _
                          & "smtp_result varchar(255)," _
                          & "mail_from varchar(255) NOT NULL," _
                          & "mail_to varchar(255) NOT NULL," _
                          & "mail_subject varchar(255)" _
                      & ")")

          sqlQueue.Enqueue("CREATE INDEX idxLOGID ON tblSmtpLog(log_id)")
          sqlQueue.Enqueue("CREATE INDEX idxLOGTS ON tblSmtpLog(log_ts)")
          sqlQueue.Enqueue("CREATE INDEX idxMSGID ON tblSmtpLog(queue_id)")

          While sqlQueue.Count > 0
            Dim strSQL As String = sqlQueue.Dequeue

            dbcmd.Connection = DBConnectionMain
            dbcmd.CommandType = CommandType.Text
            dbcmd.CommandText = strSQL

            Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
            If iRecordsAffected <> 1 Then
              'Throw New Exception("Create table failed due to error.")
            End If

            dbcmd.Dispose()
          End While
        End If

      Else
        Throw New Exception(strTableName & " not currently supported.")
      End If

      bSuccess = True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CheckDatabaseTables()")
    End Try

    Return bSuccess

  End Function

  ''' <summary>
  ''' Returns the size of the selected database
  ''' </summary>
  ''' <param name="databaseName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetDatabaseSize(ByVal databaseName As String)

    Try

      Select Case databaseName
        Case "DBConnectionMain"
          '
          ' Determine the database filename
          '
          Dim strDataSource As String = FixPath(String.Format("{0}\Data\{1}\{1}.db3", hs.GetAppPath(), IFACE_NAME.ToLower))
          Dim file As New FileInfo(strDataSource)
          Return FormatFileSize(file.Length)

      End Select

    Catch pEx As Exception

    End Try

    Return "Unknown"

  End Function

  ''' <summary>
  ''' Converts filesize to String
  ''' </summary>
  ''' <param name="FileSizeBytes"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function FormatFileSize(ByVal FileSizeBytes As Long) As String

    Try

      Dim sizeTypes() As String = {"B", "KB", "MB", "GB"}
      Dim Len As Decimal = FileSizeBytes
      Dim sizeType As Integer = 0

      Do While Len > 1024
        Len = Decimal.Round(Len / 1024, 2)
        sizeType += 1
        If sizeType >= sizeTypes.Length - 1 Then Exit Do
      Loop

      Dim fileSize As String = String.Format("{0} {1}", Len.ToString("N0"), sizeTypes(sizeType))
      Return fileSize

    Catch pEx As Exception

    End Try

    Return FileSizeBytes.ToString

  End Function

#End Region

#Region "Database Queries"

  ''' <summary>
  ''' Execute Raw SQL
  ''' </summary>
  ''' <param name="strSQL"></param>
  ''' <param name="iRecordCount"></param>
  ''' <param name="iPageSize"></param>
  ''' <param name="iPageCount"></param>
  ''' <param name="iPageCur"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ExecuteSQL(ByVal strSQL As String, _
                             ByRef iRecordCount As Integer, _
                             ByVal iPageSize As Integer, _
                             ByRef iPageCount As Integer, _
                             ByRef iPageCur As Integer) As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered ExecuteSQL() function."
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

      '
      ' Determine Requested database action
      '
      If strSQL.StartsWith("SELECT", StringComparison.CurrentCultureIgnoreCase) Then

        '
        ' Populate the DataSet
        '

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
        iRecordCount = MyDT.Rows.Count

        If iRecordCount > 0 Then
          '
          ' Determine the number of pages available
          '
          iPageSize = IIf(iPageSize <= 0, 1, iPageSize)
          iPageCount = iRecordCount \ iPageSize
          If iRecordCount Mod iPageSize > 0 Then
            iPageCount += 1
          End If

          '
          ' Find starting and ending record
          '
          Dim nStart As Integer = iPageSize * (iPageCur - 1)
          Dim nEnd As Integer = nStart + iPageSize - 1
          If nEnd > iRecordCount - 1 Then
            nEnd = iRecordCount - 1
          End If

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
          For i = nStart To nEnd
            'Add some rows
            Dim dr As DataRow
            dr = ResultsDT.NewRow()
            For iFieldNum As Integer = 0 To iFieldCount
              dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
            Next
            ResultsDT.Rows.Add(dr)
          Next

          '
          ' Make sure current page count is valid
          '
          If iPageCur > iPageCount Then iPageCur = iPageCount

        Else
          '
          ' Query succeeded, but returned 0 records
          '
          strMessage = "Your query executed and returned 0 record(s)."
          Call WriteMessage(strMessage, MessageType.Debug)

        End If

      Else
        '
        ' Execute query (does not return recordset)
        '
        Try
          '
          ' Build the insert/update/delete query
          '
          Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

            MyDbCommand.Connection = DBConnectionMain
            MyDbCommand.CommandType = CommandType.Text
            MyDbCommand.CommandText = strSQL

            Dim iRecordsAffected As Integer = 0
            SyncLock SyncLockMain
              iRecordsAffected = MyDbCommand.ExecuteNonQuery()
            End SyncLock

            strMessage = "Your query executed and affected " & iRecordsAffected & " row(s)."
            Call WriteMessage(strMessage, MessageType.Debug)

            MyDbCommand.Dispose()

          End Using

        Catch pEx As Common.DbException
          '
          ' Process Database Error
          '
          strMessage = "Your query failed for the following reason(s):  "
          strMessage &= "[Error Source: " & pEx.Source & "] " _
                      & "[Error Number: " & pEx.ErrorCode & "] " _
                      & "[Error Desciption:  " & pEx.Message & "] "
          Call WriteMessage(strMessage, MessageType.Error)
        End Try

      End If

      Call WriteMessage("SQL: " & strSQL, MessageType.Debug)
      Call WriteMessage("Record Count: " & iRecordCount, MessageType.Debug)
      Call WriteMessage("Page Count: " & iPageCount, MessageType.Debug)
      Call WriteMessage("Page Current: " & iPageCur, MessageType.Debug)

      '
      ' Populate the table name
      '
      ResultsDT.TableName = String.Format("Table1:{0}:{1}:{2}", iRecordCount, iPageCount, iPageCur)

    Catch pEx As Exception
      '
      ' Error:  Query error
      '
      strMessage = "Your query failed for the following reason:  " _
                  & Err.Source & " function/subroutine:  [" _
                  & Err.Number & " - " & Err.Description _
                  & "]"

      Call ProcessError(pEx, "ExecuteHSLogSQL()")

    End Try

    Return ResultsDT

  End Function

#End Region

#Region "SMTP Logs"

  ''' <summary>
  ''' Insert queued message into database
  ''' </summary>
  ''' <param name="smtp_id"></param>
  ''' <param name="mail_to"></param>
  ''' <param name="mail_subject"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InsertQueuedMessage(ByVal smtp_id As Integer, _
                                      ByVal mail_to As String, _
                                      ByVal mail_subject As String) As Integer

    Dim strMessage As String = ""
    Dim strSQL As String = ""
    Dim iQueueId As Integer = 0

    '
    ' Ensure database is loaded before attempting to use it
    '
    Select Case DBConnectionMain.State
      Case ConnectionState.Broken, ConnectionState.Closed
        strMessage = "Unable to complete database transaction because the database " _
                   & "connection has not been initialized."
        Throw New System.Exception(strMessage)
    End Select

    Try
      '
      ' Update the timestamp
      '
      Dim queue_ts As Double = Database.ConvertDateTimeToEpoch(DateTime.Now)

      strSQL = "INSERT INTO tblSmtpQueue ( queue_ts, delivered_ts, attempts, smtp_id, last_status, last_result, mail_to, mail_subject ) VALUES " _
               & "(" & String.Format("{0},{1},{2},{3},'{4}','{5}','{6}','{7}'", queue_ts, 0, 0, smtp_id, "Pending", "Queued", mail_to, mail_subject) & ");"
      strSQL &= "SELECT last_insert_rowid();"

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          iQueueId = dbcmd.ExecuteScalar()
        End SyncLock

        If iQueueId > 0 Then
          strMessage = String.Format("E-mail Queue Id {0} to {1} queued for delivery.", iQueueId.ToString.PadLeft(6, "0"), mail_to)
          WriteMessage(strMessage, MessageType.Informational)
        End If

        dbcmd.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process error
      '
      strMessage = "InsertQueuedMessage() Reports Error: [" & pEx.ToString & "], " & "Failed on SQL: " & strSQL & "."
      Call WriteMessage(strMessage, MessageType.Error)
    End Try

    Return iQueueId

  End Function

  ''' <summary>
  ''' Update queued message in database
  ''' </summary>
  ''' <param name="queue_id"></param>
  ''' <param name="last_status"></param>
  ''' <param name="last_result"></param>
  ''' <param name="delivered_ts"></param>
  ''' <remarks></remarks>
  Public Sub UpdateQueuedMessage(ByVal queue_id As Integer, _
                                 ByVal last_status As String, _
                                 ByVal last_result As String, _
                                 ByVal delivered_ts As Double)

    Dim strMessage As String = ""
    Dim strSQL As String = ""
    Dim iRecordsAffected As Integer = 0

    '
    ' Ensure database is loaded before attempting to use it
    '
    If bDBInitialized = False Then Exit Sub

    Try
      '
      ' Update the timestamp
      '
      strSQL = String.Format("UPDATE tblSmtpQueue " _
                            & "SET delivered_ts={0}, " _
                            & " attempts=attempts + 1, " _
                            & " last_status='{1}', " _
                            & " last_result='{2}' " _
                            & "WHERE queue_id={3}", delivered_ts, last_status, last_result, queue_id)

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          iRecordsAffected = dbcmd.ExecuteNonQuery()
        End SyncLock

        dbcmd.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process error
      '
      strMessage = "UpdateQueuedMessage() Reports Error: [" & pEx.ToString & "], " & "Failed on SQL: " & strSQL & "."
      Call WriteMessage(strMessage, MessageType.Error)
    End Try

  End Sub

  ''' <summary>
  ''' Gets the delivery attempts for a message
  ''' </summary>
  ''' <param name="queue_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetDeliveryAttempts(ByVal queue_id As Integer) As Integer

    Dim strMessage As String = ""
    Dim iAttempts As Integer = 0

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim uuid As String = ""
      Dim addr As String = ""
      Dim irbutton As String = ""
      Dim irdata As String = ""

      Dim strSQL As String = String.Format("SELECT attempts FROM tblSmtpQueue WHERE queue_id={0}", queue_id)

      '
      ' Execute the data reader
      '
      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = dbcmd.ExecuteReader()

          '
          ' Process the resutls
          '
          While dtrResults.Read()
            iAttempts = dtrResults("attempts")
          End While

          dtrResults.Close()
        End SyncLock

        dbcmd.Dispose()

      End Using

      Return iAttempts
    Catch pEx As Exception
      Call ProcessError(pEx, "GetDeliveryAttempts()")
      Return -1
    End Try

  End Function

  ''' <summary>
  ''' Insert Smtp log data into database
  ''' </summary>
  ''' <param name="queue_id"></param>
  ''' <param name="smtp_id"></param>
  ''' <param name="smtp_server"></param>
  ''' <param name="smtp_status"></param>
  ''' <param name="smtp_result"></param>
  ''' <param name="mail_from"></param>
  ''' <param name="mail_to"></param>
  ''' <param name="mail_subject"></param>
  ''' <remarks></remarks>
  Public Sub InsertSmtpLog(ByVal queue_id As Integer, _
                           ByVal smtp_id As Integer, _
                           ByVal smtp_server As String, _
                           ByVal smtp_status As String, _
                           ByVal smtp_result As String, _
                           ByVal mail_from As String, _
                           ByVal mail_to As String, _
                           ByVal mail_subject As String)

    Dim strMessage As String = ""
    Dim strSQL As String = ""
    Dim iRecordsAffected As Integer = 0

    Try
      '
      ' Ensure database is loaded before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      '
      '
      ' Update the timestamp
      '
      Dim log_ts As Double = Database.ConvertDateTimeToEpoch(DateTime.Now)
      Dim delivered_ts As Double = IIf(smtp_status = "Delivered", log_ts, 0)

      '
      ' Update the queued message
      '
      Database.UpdateQueuedMessage(queue_id, smtp_status, smtp_result, delivered_ts)

      strSQL = "INSERT INTO tblSmtpLog ( log_ts, queue_id, smtp_id, smtp_server, smtp_status, smtp_result, mail_from, mail_to, mail_subject ) VALUES " _
               & "(" & String.Format("{0},{1},{2},'{3}','{4}','{5}','{6}','{7}','{8}'", log_ts, queue_id, smtp_id, smtp_server, smtp_status, smtp_result, mail_from, mail_to, mail_subject) & ")"

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          iRecordsAffected = dbcmd.ExecuteNonQuery()
        End SyncLock

        dbcmd.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process error
      '
      strMessage = "InsertSmtpLog() Reports Error: [" & pEx.ToString & "], " & "Failed on SQL: " & strSQL & "."
      Call WriteMessage(strMessage, MessageType.Error)
    End Try

  End Sub

#End Region

#Region "Database Date Formatting"

  ''' <summary>
  ''' dateTime as DateTime
  ''' </summary>
  ''' <param name="dateTime"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ConvertDateTimeToEpoch(ByVal dateTime As DateTime) As Long

    Dim baseTicks As Long = 621355968000000000
    Dim tickResolution As Long = 10000000

    Return (dateTime.ToUniversalTime.Ticks - baseTicks) / tickResolution

  End Function

  ''' <summary>
  ''' Converts Epoch to datetime
  ''' </summary>
  ''' <param name="epochTicks"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ConvertEpochToDateTime(ByVal epochTicks As Long) As DateTime

    '
    ' Create a new DateTime value based on the Unix Epoch
    '
    Dim converted As New DateTime(1970, 1, 1, 0, 0, 0, 0)

    '
    ' Return the value in string format
    '
    Return converted.AddSeconds(epochTicks).ToLocalTime

  End Function

  ''' <summary>
  ''' Converts date to format recognized by all regions
  ''' </summary>
  ''' <param name="strDate"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function RegionalizeDate(ByVal strDate As String) As String

    Dim ci As CultureInfo = New CultureInfo(Thread.CurrentThread.CurrentUICulture.ToString())
    Dim TheDate As New DateTime

    Try
      '
      ' Try to parse the date provided
      '
      TheDate = Date.Parse(strDate)
    Catch pEx As Exception
      '
      ' Let's just return the current date
      '
      TheDate = Date.Parse(Date.Now)
    End Try

    Return TheDate.ToString("F", ci)

  End Function

#End Region

End Module

