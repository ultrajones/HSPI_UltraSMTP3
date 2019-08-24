Module hspi_logging

  Public gPluginPath As String = System.IO.Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory)
  Public gLogFile As String = gPluginPath & "\hspi_" & LCase(IFACE_NAME) & "_debug.log"
  Public gLogLevel As Integer = LogLevel.Debug

#Region "Logging - Enums"

  ''' <summary>
  ''' Used to specify the log level for logging to the HomeSeer log
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum LogLevel As Integer
    Emergency = 0
    Alert = 1
    Critical = 2
    [Error] = 3
    Warning = 4
    Notice = 5
    Informational = 6
    Trace = 7
    Debug = 8
  End Enum

  ''' <summary>
  ''' Used to specify the mesasge type for logging to the HomeSeer log
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum MessageType
    Emergency = 0
    Alert = 1
    Critical = 2
    [Error] = 3
    Warning = 4
    Notice = 5
    Informational = 6
    Trace = 7
    Debug = 8
  End Enum

#End Region

#Region "Logging - Functions and Routines"

  ''' <summary>
  ''' Writes a message to the HomeSeer log
  ''' </summary>
  ''' <param name="Message"></param>
  ''' <param name="MessageType"></param>
  ''' <param name="LogColor"></param>
  ''' <param name="LogPriority"></param>
  ''' <param name="LogSource"></param>
  ''' <param name="LogErrorCode"></param>
  ''' <remarks></remarks>
  Public Sub WriteMessage(ByVal Message As String,
                          Optional ByVal MessageType As MessageType = MessageType.Informational,
                          Optional ByVal LogColor As String = "#000000",
                          Optional ByVal LogPriority As Integer = 0,
                          Optional ByVal LogSource As String = IFACE_NAME,
                          Optional ByVal LogErrorCode As Integer = 0)

    Try

      Dim MessageLogLevel As Integer = DirectCast([Enum].Parse(GetType(MessageType), MessageType), MessageType)

      '
      ' Ignore message based on selected logging level
      '
      If MessageLogLevel > gLogLevel Then Exit Sub

      Dim LogType As String = CType(MessageType, MessageType).ToString()

      If gHSInitialized = True Then
        '
        ' Don't write debug logs to the HomeSeer logs
        '
        If MessageLogLevel < MessageType.Debug Then
          hs.WriteLogDetail(LogType, Message, LogColor, LogPriority, LogSource, LogErrorCode)
        End If

      End If

      '
      ' Write all messages to console and disk if in debug mode
      '
      If gLogLevel >= LogLevel.Debug Then
        Console.WriteLine(Message)
        Call WriteToDisk(LogType, Message)
      End If

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' Writes a log message to disk
  ''' </summary>
  ''' <param name="strLogType"></param>
  ''' <param name="strMessage"></param>
  ''' <param name="bAppend"></param>
  ''' <remarks></remarks>
  Public Sub WriteToDisk(ByVal strLogType As String,
                         ByVal strMessage As String,
                         Optional ByVal bAppend As Boolean = True)

    Try

      Using objFile As New System.IO.StreamWriter(gLogFile, bAppend)
        objFile.WriteLine("{0}...{1}~~!~~{2}", Now(), strLogType, strMessage)
        objFile.Close()
      End Using

    Catch pEx As Exception
      '
      ' Process error condition
      '
    End Try

  End Sub

  ''' <summary>
  ''' Writes a program exception to the HomeSeer log
  ''' </summary>
  ''' <param name="pEx"></param>
  ''' <param name="Source"></param>
  ''' <remarks></remarks>
  Public Sub ProcessError(ByRef pEx As Exception, ByVal Source As String)

    Try

      Call WriteMessage(pEx.GetType().FullName, MessageType.Error, Nothing, 1, Source, 0)
      Call WriteMessage(pEx.ToString, MessageType.Error, Nothing, 1, Source, 0)

    Catch ex As Exception

    End Try

  End Sub

#End Region

End Module
