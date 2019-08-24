Imports Scheduler
Imports HomeSeerAPI
Imports HSCF.Communication.Scs.Communication.EndPoints.Tcp
Imports HSCF.Communication.ScsServices.Client
Imports HSCF.Communication.ScsServices.Service

Module hspi_instance

  Public WithEvents client As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IHSApplication)
  Dim WithEvents clientCallback As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IAppCallbackAPI)

  Public InstanceFriendlyName As String = ""
  Public Instance As String = ""

  Private host As HomeSeerAPI.IHSApplication
  Private gAppAPI As HSPI

  Private sIp As String = "127.0.0.1"

  Public AllInstances As New SortedList

  Public Class InstanceHolder

    Public hspi As HSPI
    Public client As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IHSApplication)
    Public clientCallback As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IAppCallbackAPI)
    Public host As HomeSeerAPI.IHSApplication

  End Class

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <remarks></remarks>
  Sub Main()

    Dim argv As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
    argv = My.Application.CommandLineArgs

    '
    ' Read arguments from the command line
    '
    For Each sCmd As String In argv
      Dim ch(0) As String
      ch(0) = "="
      Dim parts() As String = sCmd.Split(ch, StringSplitOptions.None)
      Select Case parts(0).ToLower
        Case "server" : sIp = parts(1)
        Case "instance"
          Try
            instance = parts(1)
          Catch ex As Exception
            instance = ""
          End Try
      End Select
    Next

    gAppAPI = New HSPI

    WriteMessage(String.Format("{0} HSPI starting...", IFACE_NAME, MessageType.Informational))

    client = ScsServiceClientBuilder.CreateClient(Of IHSApplication)(New ScsTcpEndPoint(sIp, 10400), gAppAPI)
    clientCallback = ScsServiceClientBuilder.CreateClient(Of IAppCallbackAPI)(New ScsTcpEndPoint(sIp, 10400), gAppAPI)

    Dim Attempts As Integer = 1
    Dim bConnected As Boolean = False

    Do

      Try

        WriteMessage(String.Format("Connecting to HomeSeer server at {0}, attempt {1}...", sIp, Attempts), MessageType.Informational)

        client.Connect()
        clientCallback.Connect()

        host = client.ServiceProxy

        Dim APIVersion As Double = host.APIVersion  ' Will cause an error if not really connected

        callback = clientCallback.ServiceProxy
        APIVersion = callback.APIVersion            ' Will cause an error if not really connected

        bConnected = True

      Catch pEx As Exception

        WriteMessage(String.Format("Connection to HomeSeer at {0} attempt {1} failed due to error:  {2}", sIp, Attempts, pEx.Message), MessageType.Warning)
        If pEx.Message.ToLower.Contains("timeout occured.") Then
          Attempts += 1
        Else
          bConnected = False
          Exit Do
        End If

      End Try

    Loop While bConnected = False And Attempts < 6

    '
    ' Check to see if we are connnected
    '
    If bConnected = False Then

      If client IsNot Nothing Then
        client.Dispose()
        client = Nothing
      End If

      If clientCallback IsNot Nothing Then
        clientCallback.Dispose()
        clientCallback = Nothing
      End If

      wait(4)
      Return

    End If

    Try
      '
      ' Connect to HS so it can register a callback to us
      '
      host.Connect(IFACE_NAME, Instance)

      '
      ' Create the user object that is the real plugin, accessed from the pluginAPI wrapper
      '
      callback = callback
      hs = host

      gAppAPI.OurInstanceFriendlyName = Instance
      WriteMessage(String.Format("Connection to HomeSeer at {0}, attempt {1} succeeded.", sIp, Attempts), MessageType.Informational)

      Do
        Threading.Thread.Sleep(10)
      Loop While client.CommunicationState = HSCF.Communication.Scs.Communication.CommunicationStates.Connected And Not bShutDown

      If Not bShutDown Then
        gAppAPI.ShutdownIO()
        WriteMessage(String.Format("Connection to HomeSeer at {0} was lost, exiting.", sIp), MessageType.Critical)
      Else
        WriteMessage(String.Format("Shutting down...", IFACE_NAME, MessageType.Informational))
      End If

      '
      ' Disconnect from server for good here
      '
      client.Disconnect()
      clientCallback.Disconnect()

    Catch pEx As Exception
      WriteMessage(String.Format("Unable to connect to HomeSeer at {0} failed due to error: {1}", sIp, pEx.Message), MessageType.Critical)
    Finally
      wait(2)
      End
    End Try

  End Sub

  Public Function AddInstance(InstanceName As String) As String

    '
    ' Write the debug message
    '
    WriteMessage("Entered AddInstance() function.", MessageType.Debug)

    WriteMessage(String.Format("Attempting to add new plug-in instance name {0}.", InstanceName), MessageType.Informational)

    If AllInstances.Contains(InstanceName) Then
      Return "Cannot add duplicate instance"
    End If

    Dim PlugAPI As HSPI = New HSPI
    PlugAPI.instance = InstanceName

    Dim lhost As HomeSeerAPI.IHSApplication
    Dim lclient As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IHSApplication)
    Dim lclientCallback As HSCF.Communication.ScsServices.Client.IScsServiceClient(Of IAppCallbackAPI)

    lclient = ScsServiceClientBuilder.CreateClient(Of IHSApplication)(New ScsTcpEndPoint(sIp, 10400), PlugAPI)
    lclientCallback = ScsServiceClientBuilder.CreateClient(Of IAppCallbackAPI)(New ScsTcpEndPoint(sIp, 10400), PlugAPI)

    Try

      lclient.Connect()
      lclientCallback.Connect()
      lhost = lclient.ServiceProxy

    Catch pEx As Exception
      WriteMessage(String.Format("Cannot start instance {0}:{1}", InstanceName, pEx.Message), MessageType.Critical)
      Return pEx.Message
    End Try

    Try

      lhost.Connect(IFACE_NAME, InstanceName)

      ' everything is ok, save instance
      Dim ih As New InstanceHolder
      ih.client = lclient
      ih.clientCallback = lclientCallback
      ih.host = lhost
      ih.hspi = PlugAPI
      AllInstances.Add(InstanceName, ih)

    Catch pEx As Exception
      WriteMessage(String.Format("Error connecting instance {0}:{1}", InstanceName, pEx.Message), MessageType.Critical)
    End Try

    Return ""
  End Function

  ''' <summary>
  ''' Removes a plug-in instance
  ''' </summary>
  ''' <param name="InstanceName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function RemoveInstance(InstanceName As String) As String

    If Not AllInstances.Contains(InstanceName) Then
      WriteMessage(String.Format("Error removing instance {0} because it does not exist.", InstanceName), MessageType.Critical)
      Return "Instance does not exist"
    End If

    Try

      For Each DE As DictionaryEntry In AllInstances
        Dim key As String = DE.Key
        If key.ToLower = InstanceName.ToLower Then

          Dim ih As InstanceHolder = DE.Value
          ih.hspi.ShutdownIO()
          ih.client.Disconnect()
          ih.clientCallback.Disconnect()

          AllInstances.Remove(key)
          Exit For
        End If
      Next

    Catch pEx As Exception
      WriteMessage(String.Format("Error removing instance {0}:{1}", InstanceName, pEx.Message), MessageType.Critical)
    End Try

    Return ""

  End Function

  ''' <summary>
  ''' Subroutined called when the client disconnect from the server
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Private Sub client_Disconnected(ByVal sender As Object, ByVal e As System.EventArgs) Handles client.Disconnected

    WriteMessage("Disconnected from server - client", MessageType.Informational)

  End Sub

  ''' <summary>
  ''' Sleeps for the number of seconds specified
  ''' </summary>
  ''' <param name="secs"></param>
  ''' <remarks></remarks>
  Private Sub wait(ByVal secs As Integer)

    Threading.Thread.Sleep(secs * 1000)

  End Sub

End Module

