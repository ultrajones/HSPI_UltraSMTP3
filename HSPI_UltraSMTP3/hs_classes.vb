Module hs_classes

  <Serializable()> _
  Public Class hsCollection

    Inherits Dictionary(Of String, Object)
    Dim KeyIndex As New Collection

    Public Sub New()
      MyBase.New()
    End Sub

    Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
      MyBase.New(info, context)
    End Sub

    Public Overloads Sub Add(value As Object, Key As String)
      If Not MyBase.ContainsKey(Key) Then
        MyBase.Add(Key, value)
        KeyIndex.Add(Key, Key)
      Else
        MyBase.Item(Key) = value
      End If
    End Sub

    Public Overloads Sub Remove(Key As String)
      On Error Resume Next
      MyBase.Remove(Key)
      KeyIndex.Remove(Key)
    End Sub

    Public Overloads Sub Remove(Index As Integer)
      MyBase.Remove(KeyIndex(Index))
      KeyIndex.Remove(Index)
    End Sub

    Public Overloads ReadOnly Property Keys(ByVal index As Integer) As Object
      Get
        Dim i As Integer
        Dim key As String = Nothing
        For Each key In MyBase.Keys
          If i = index Then
            Exit For
          Else
            i += 1
          End If
        Next
        Return key
      End Get
    End Property

    Default Public Overloads Property Item(ByVal index As Integer) As Object
      Get
        Return MyBase.Item(KeyIndex(index))
      End Get
      Set(ByVal value As Object)
        MyBase.Item(KeyIndex(index)) = value
      End Set
    End Property

    Default Public Overloads Property Item(ByVal Key As String) As Object
      Get
        On Error Resume Next
        Return MyBase.Item(Key)
      End Get
      Set(ByVal value As Object)
        If Not MyBase.ContainsKey(Key) Then
          Add(value, Key)
        Else
          MyBase.Item(Key) = value
        End If
      End Set
    End Property

  End Class

  <Serializable()> _
  Public Class action

    Private _uid As Integer = 0
    Private params As New Specialized.StringDictionary

    Sub New()
      MyBase.New()
    End Sub

    Public Property uid As Integer
      Get
        Return _uid
      End Get
      Set(value As Integer)
        _uid = value
      End Set
    End Property

    Public Property Item(ByVal Key As String) As String
      Get
        On Error Resume Next
        If Not params.ContainsKey(Key) Then
          Return ""
        Else
          Return params(Key)
        End If
      End Get
      Set(ByVal value As String)
        On Error Resume Next
        If Not params.ContainsKey(Key) Then
          params.Add(Key, value)
        Else
          params(Key) = value
        End If
      End Set
    End Property

  End Class

  <Serializable()> _
  Public Class trigger

    Private _uid As Integer = 0
    Private params As New Specialized.StringDictionary
    Private mvarCondition As Boolean = False

    Sub New()
      MyBase.New()
    End Sub

    Public Property uid As Integer
      Get
        Return _uid
      End Get
      Set(value As Integer)
        _uid = value
      End Set
    End Property

    Public Property Condition As Boolean
      Get
        Return mvarCondition
      End Get
      Set(value As Boolean)
        mvarCondition = value
      End Set
    End Property

    Public Property Item(ByVal Key As String) As String
      Get
        On Error Resume Next
        If Not params.ContainsKey(Key) Then
          Return ""
        Else
          Return params(Key)
        End If
      End Get
      Set(ByVal value As String)
        On Error Resume Next
        If Not params.ContainsKey(Key) Then
          params.Add(Key, value)
        Else
          params(Key) = value
        End If
      End Set
    End Property

  End Class

End Module
