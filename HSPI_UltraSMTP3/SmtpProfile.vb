Public Class SmtpProfile

  Private m_SmtpId As Integer = 0
  Private m_SmtpServer As String = ""
  Private m_SmtpPort As Integer = 25
  Private m_SmtpSsl As Boolean = False
  Private m_AuthUser As String = ""
  Private m_AuthPass As String = ""
  Private m_SmtpFrom As String = ""

  Public Property SmtpId() As Integer
    Get
      Return m_SmtpId
    End Get
    Set(ByVal value As Integer)
      m_SmtpId = value
    End Set
  End Property

  Public Property SmtpSsl() As Boolean
    Get
      Return m_SmtpSsl
    End Get
    Set(ByVal value As Boolean)
      m_SmtpSsl = value
    End Set
  End Property

  Public Property SmtpServer() As String
    Get
      Return m_SmtpServer
    End Get
    Set(ByVal value As String)
      m_SmtpServer = value
    End Set
  End Property

  Public Property SmtpPort() As Integer
    Get
      Return m_SmtpPort
    End Get
    Set(ByVal value As Integer)
      m_SmtpPort = value
    End Set
  End Property

  Public Property AuthUser() As String
    Get
      Return m_AuthUser
    End Get
    Set(ByVal value As String)
      m_AuthUser = value
    End Set
  End Property

  Public Property AuthPass() As String
    Get
      Return m_AuthPass
    End Get
    Set(ByVal value As String)
      m_AuthPass = value
    End Set
  End Property

  Public Property SmtpFrom() As String
    Get
      Return m_SmtpFrom
    End Get
    Set(ByVal value As String)
      m_SmtpFrom = value
    End Set
  End Property

End Class

