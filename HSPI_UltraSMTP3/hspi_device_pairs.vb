Public Class hspi_device_pairs

  Private _value As Integer = 0
  Private _status As String = ""
  Private _image As String = ""
  Private _type As Integer = 0

  Sub New(ByVal Value As Integer, ByVal Status As String, ByVal Image As String, ByVal Type As Integer)
    _value = Value
    _status = Status
    _image = Image
    _type = Type
  End Sub

  Public ReadOnly Property Value As String
    Get
      Return _value.ToString
    End Get
  End Property

  Public ReadOnly Property Status As String
    Get
      Return _status
    End Get
  End Property

  Public ReadOnly Property Image As String
    Get
      Return _image
    End Get
  End Property

  Public ReadOnly Property Type As Integer
    Get
      Return _type
    End Get
  End Property

  Protected Overrides Sub Finalize()
    MyBase.Finalize()
  End Sub
End Class
