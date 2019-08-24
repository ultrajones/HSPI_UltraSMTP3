Imports System.Threading
Imports System.Globalization
Imports System.ComponentModel
Imports System.Reflection

Module hspi_main

  Public hs As HomeSeerAPI.IHSApplication
  Public callback As HomeSeerAPI.IAppCallbackAPI

  Public InterfaceVersion As Integer

  Public bShutDown As Boolean = False

  Public nfi As NumberFormatInfo = New CultureInfo("en-US", False).NumberFormat

  <System.Runtime.CompilerServices.Extension> _
  Public Function StringToEnum(Of EnumType)(enumValue As [String]) As EnumType
    Return DirectCast([Enum].Parse(GetType(EnumType), enumValue), EnumType)
  End Function

  <System.Runtime.CompilerServices.Extension> _
  Public Function GetEnumName(eff As [Enum]) As [String]
    Return [Enum].GetName(eff.[GetType](), eff)
  End Function

  Public Function GetEnumDescription(en As [Enum]) As String
    Dim type As Type = en.[GetType]()

    Dim memInfo As MemberInfo() = type.GetMember(en.ToString())

    If memInfo IsNot Nothing AndAlso memInfo.Length > 0 Then
      Dim attrs As Object() = memInfo(0).GetCustomAttributes(GetType(DescriptionAttribute), False)

      If attrs IsNot Nothing AndAlso attrs.Length > 0 Then
        Return DirectCast(attrs(0), DescriptionAttribute).Description
      End If
    End If

    Return en.ToString()
  End Function

  ''' <summary>
  ''' Plug-in Connection Thread
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub PluginConnectionThread()

    Dim bAbortThread As Boolean = False

    Try

      While bAbortThread = False
        '
        ' Call the Refresh Routine
        '
        WriteMessage("Thread running refresh routine ...", MessageType.Debug)

        '
        ' Give up some time
        '
        Thread.Sleep(1000 * 60)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("ConnectionThread thread received abort request, terminating normally."), MessageType.Debug)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "ECMConnectionThread()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("ConnectionThread terminated."), MessageType.Debug)

    End Try

  End Sub

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the list of users authorized to access the web page
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function WEBUserRolesAuthorized() As Integer

    Dim USERS_AUTHORIZED As Integer = USER_ADMIN

    Try
      Dim AuthorizedRoles As String = hs.GetINISetting("WEBUsers", "AuthorizedRoles", USER_ADMIN, gINIFile)
      If IsNumeric(AuthorizedRoles) Then
        USERS_AUTHORIZED = CInt(AuthorizedRoles)
      End If

    Catch pEx As Exception
      '
      ' Ignore this error
      '
      USERS_AUTHORIZED = USER_ADMIN
    End Try

    Return USERS_AUTHORIZED

  End Function

  '----------------------------------------------------------------------
  'Purpose: Determine if logged in user is authorized to view the web page
  'Inputs:  LoggedInUser as String
  'Outputs: Boolean (True indicates user is authorized)
  '----------------------------------------------------------------------
  ''' <summary>
  ''' Determine if logged in user is authorized to view the web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <param name="USER_ROLES_AUTHORIZED"></param>
  ''' <returns>Boolean (True indicates user is authorized)</returns>
  ''' <remarks></remarks>
  Public Function WEBUserIsAuthorized(ByVal LoggedInUser As String, _
                                      ByVal USER_ROLES_AUTHORIZED As Integer) As Boolean

    Dim bAuthorized As Boolean = False

    Try
      '
      ' Obtain the list of users from HomeSeer
      '      
      Dim Users As String = hs.GetUsers()
      Dim UserPairs() As String = Users.Split(",")

      For Each UserPair As String In UserPairs
        Dim User() As String = UserPair.Split("|")

        Dim UserName As String = User(0)
        Dim UserRights As Integer = CInt(User(1))

        If String.Compare(LoggedInUser, User(0)) = 0 Then
          If (UserRights And USER_GUEST) = USER_GUEST Then
            bAuthorized = USER_GUEST And USER_ROLES_AUTHORIZED
          ElseIf (UserRights And USER_ADMIN) = USER_ADMIN Then
            bAuthorized = (USER_ADMIN And USER_ROLES_AUTHORIZED)
          ElseIf (UserRights And USER_LOCAL) = USER_LOCAL Then
            bAuthorized = USER_LOCAL And USER_ROLES_AUTHORIZED
          ElseIf (UserRights And USER_NORMAL) = USER_NORMAL Then
            bAuthorized = USER_NORMAL And USER_ROLES_AUTHORIZED
          End If
        End If
      Next

    Catch pEx As Exception
      '
      ' Lets be safe and return false
      '
      bAuthorized = False
    End Try

    Return bAuthorized

  End Function

#End Region

End Module
