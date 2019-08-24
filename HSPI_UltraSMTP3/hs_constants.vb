Module hs_constants

  ' User access levels

  Public Const USER_GUEST As Integer = 1    ' user can view web pages only, cannot make changes
  Public Const USER_ADMIN As Integer = 2    ' user can make changes 
  Public Const USER_LOCAL As Integer = 4    ' this user is used when logging in on a local subnet
  Public Const USER_NORMAL As Integer = 8   ' Not guest, not admin, just NORMAL! 

  ' Plug-In Capabilities

  Public Const CA_IO As Integer = 4        ' supports I/O, must be defined for ALL plugins
  Public Const CA_SEC As Integer = 8       ' security, currently not supporte
  Public Const CA_THERM As Integer = 16    ' Indicates a thermostat plug-in.
  Public Const CA_MUSIC As Integer = 32    ' Music API, currently not supported 

  ' Device MISC bit settings

  Public Const MISC_NO_LOG As Integer = 8                 ' no logging to event log for this device
  Public Const MISC_STATUS_ONLY As Integer = &H10       ' device cannot be controlled
  Public Const MISC_HIDDEN As Integer = &H20            ' device is hidden from views
  Public Const MISC_INCLUDE_PF As Integer = &H80        ' if set, device's state is restored if power fail enabled
  Public Const MISC_SHOW_VALUES As Integer = &H100      ' set=display value options in win gui and web status
  Public Const MISC_AUTO_VC As Integer = &H200          ' set=create a voice command for this device
  Public Const MISC_VC_CONFIRM As Integer = &H400       ' set=confirm voice command
  Public Const MISC_SETSTATUS_NOTIFY As Integer = &H4000  ' if set, SetDeviceStatus calls plugin SetIO
  Public Const MISC_SETVALUE_NOTIFY As Integer = &H8000 ' if set, SetDeviceValue calls plugin SetIO 
  Public Const MISC_NO_STATUS_TRIG As Integer = &H20000 ' if set, the device will not appear in the device status

  'HSEvent Callback Types used with RegisterEventCB

  Public Const EV_TYPE_X10 As Integer = 1
  Public Const EV_TYPE_LOG As Integer = 2
  Public Const EV_TYPE_STATUS_CHANGE As Integer = 4
  Public Const EV_TYPE_AUDIO As Integer = 8
  Public Const EV_TYPE_X10_TRANSMIT As Integer = &H10S
  Public Const EV_TYPE_CONFIG_CHANGE As Integer = &H20S
  Public Const EV_TYPE_STRING_CHANGE As Integer = &H40S
  Public Const EV_TYPE_SPEAKER_CONNECT As Integer = &H80S
  Public Const EV_TYPE_CALLER_ID As Integer = &H100
  Public Const EV_TYPE_ZWAVE As Integer = &H200
  Public Const EV_TYPE_VALUE_CHANGE As Integer = &H400
  Public Const EV_TYPE_STATUS_ONLY_CHANGE As Integer = &H800
  Public Const EV_TYPE_GENERIC As Integer = &H8000

  ' Phone LINEStatus Values

  Public Const LINE_IDLE As Integer = 0
  Public Const LINE_OFFERING As Integer = 1
  Public Const LINE_RINGING As Integer = 2
  Public Const LINE_CONNECTED As Integer = 3
  Public Const LINE_INACTIVE As Integer = 4
  Public Const LINE_BUSY As Integer = 5
  Public Const LINE_INUSE As Integer = 6
  Public Const LINE_TIMEOUT As Integer = 7 ' for calling

  ''' <summary>
  ''' This Enum holds values referencing individual bits in an integer which indicate different characteristics of a device.
  ''' </summary>
  ''' <remarks></remarks>
  Enum dvMISC As UInteger
    NO_LOG = 8                      ' No logging to the log for this device
    STATUS_ONLY = &H10              ' Device cannot be controlled
    HIDDEN = &H20                   ' Device is hidden from the device utility page when Hide Marked is used.
    INCLUDE_POWERFAIL = &H80        ' The device's state is restored if power failure recovery is enabled
    SHOW_VALUES = &H100             ' If not set, device control options will not be displayed.
    AUTO_VOICE_COMMAND = &H200      ' When set, this device is included in the voice recognition context for device commands.
    VOICE_COMMAND_CONFIRM = &H400   ' When set, voice commands for this device are confirmed.
    NO_STATUS_TRIGGER = &H20000     ' If set, the device status values will not appear in the device change trigger.
    CONTROL_POPUP = &H100000        ' The controls for this device should appear in a popup window on the device utility page.
  End Enum

  ''' <summary>
  ''' This eNum is used as the return for the Parent and Child properties of the DeviceClass object, and are as follows:
  ''' </summary>
  ''' <remarks></remarks>
  Enum eRelationship As Integer
    Not_Set = 0
    Indeterminate = 1   ' Could not be determined
    Parent_Root = 2
    Standalone = 3
    Child = 4
  End Enum

  ''' <summary>
  '''  This Enum is used when the Device_Type API is set to Script, and the Device_Type type is set to one of the script run values (See eDeviceType_Script).  
  ''' This Enum is one of the parameters passed to the script that is run when the device changes, and it indicates what changed to cause the script to be run.
  ''' </summary>
  ''' <remarks></remarks>
  Enum DeviceScriptChange As Integer
    DevValue = 1        ' The device value changed.
    DevString = 2       ' The device string changed.
    Both = 3            ' Both the device value and string changed.
  End Enum

End Module
