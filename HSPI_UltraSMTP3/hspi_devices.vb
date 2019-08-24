Imports HomeSeerAPI
Imports Scheduler
Imports HSCF
Imports HSCF.Communication.ScsServices.Service

Module hspi_devices

  Public Function GetDeviceList() As SortedList

    Dim DevEnum As Scheduler.Classes.clsDeviceEnumeration
    Dim dv As Scheduler.Classes.DeviceClass
    Dim DeviceList As New SortedList

    Try
      '
      ' Go through devices to see if we have one assigned to our plug-in
      '
      DevEnum = hs.GetDeviceEnumerator

      If Not DevEnum Is Nothing Then

        Do While Not DevEnum.Finished
          dv = DevEnum.GetNext
          If dv Is Nothing Then Continue Do
          If dv.Interface(Nothing) IsNot Nothing Then

            Dim dv_ref As Integer = dv.Ref(Nothing)
            Dim dv_name As String = dv.Name(Nothing)
            Dim dv_type As String = dv.Device_Type_String(Nothing)
            Dim dv_loc1 As String = dv.Location(Nothing)
            Dim dv_loc2 As String = dv.Location2(Nothing)

            'Dim key As String = dv_ref.ToString
            'Dim value As String = String.Format("{0} {1} {2}", dv_loc1, dv_loc2, dv_name)

            Dim key As String = String.Format("{0} {1} {2}", dv_loc1, dv_loc2, dv_name)
            Dim value As String = dv_ref.ToString

            If DeviceList.ContainsKey(key) = False Then
              DeviceList.Add(key, value)
            End If

          End If
        Loop
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "GetDeviceList")
    End Try

    Return DeviceList

  End Function

  ''' <summary>
  ''' Locates device by device by ref
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByRef(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      Long.TryParse(strDeviceAddr, dev_ref)
      If dev_ref > 0 Then
        objDevice = hs.GetDeviceByRef(dev_ref)
        If Not objDevice Is Nothing Then
          Return objDevice
        End If
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByRef")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Locates device by device by code
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByCode(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.DeviceExistsCode(strDeviceAddr)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByCode")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Locates device by device address
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByAddr(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.DeviceExistsAddress(strDeviceAddr, False)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByAddr")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Sets the HomeSeer string and device values
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_value"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceValue(ByVal dv_addr As String,
                            ByVal dv_value As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_value), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        If IsNumeric(dv_value) Then

          Dim dblDeviceValue As Double = Double.Parse(hs.DeviceValueEx(dv_ref))
          Dim dblSensorValue As Double = Double.Parse(dv_value)

          If dblDeviceValue <> dblSensorValue Then
            hs.SetDeviceValueByRef(dv_ref, dblSensorValue, True)
          End If

        End If

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceValue()")

    End Try

  End Sub

  ''' <summary>
  ''' Sets the HomeSeer device string
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_string"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceString(ByVal dv_addr As String,
                             ByVal dv_string As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_string), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        hs.SetDeviceString(dv_ref, dv_string, True)

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceString()")

    End Try

  End Sub

End Module
