Imports System.Text
Imports System.Web
Imports Scheduler
Imports System.Collections.Specialized
Imports System.Text.RegularExpressions

Public Class hspi_webconfig
  Inherits clsPageBuilder

  Public hspiref As HSPI

  Dim TimerEnabled As Boolean

  ''' <summary>
  ''' Initializes new webconfig
  ''' </summary>
  ''' <param name="pagename"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal pagename As String)
    MyBase.New(pagename)
  End Sub

#Region "Page Building"

  ''' <summary>
  ''' Web pages that use the clsPageBuilder class and registered with hs.RegisterLink and hs.RegisterConfigLink will then be called through this function. 
  ''' A complete page needs to be created and returned.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="queryString"></param>
  ''' <param name="instance"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, instance As String) As String

    Try

      Dim stb As New StringBuilder

      '
      ' Called from the start of your page to reset all internal data structures in the clsPageBuilder class, such as menus.
      '
      Me.reset()

      '
      ' Determine if user is authorized to access the web page
      '
      Dim LoggedInUser As String = hs.WEBLoggedInUser()
      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      '
      ' Handle any queries like mode=something
      '
      Dim parts As Collections.Specialized.NameValueCollection = Nothing
      If (queryString <> "") Then
        parts = HttpUtility.ParseQueryString(queryString)
      End If

      Dim Header As New StringBuilder
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrasmtp3/css/jquery.dataTables.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrasmtp3/css/dataTables.tableTools.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrasmtp3/css/dataTables.editor.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrasmtp3/css/jquery.dataTables_themeroller.css"" rel=""stylesheet"" />")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrasmtp3/js/jquery.dataTables.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrasmtp3/js/dataTables.tableTools.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrasmtp3/js/dataTables.editor.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrasmtp3/js/hspi_ultrasmtp3_profiles.js""></script>")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrasmtp3/js/hspi_ultrasmtp3_logs.js""></script>")

      Me.AddHeader(Header.ToString)

      Dim pageTile As String = String.Format("{0} {1}", pageName, instance).TrimEnd
      stb.Append(hs.GetPageHeader(pageName, pageTile, "", "", False, False))

      '
      ' Start the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

      '
      ' A message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
      '
      stb.Append(clsPageBuilder.DivStart("divErrorMessage", "class='errormessage'"))
      stb.Append(clsPageBuilder.DivEnd)

      Me.RefreshIntervalMilliSeconds = 3000
      stb.Append(Me.AddAjaxHandlerPost("id=timer", pageName))

      If WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED) = False Then
        '
        ' Current user not authorized
        '
        stb.Append(WebUserNotUnauthorized(LoggedInUser))
      Else
        '
        ' Specific page starts here
        '
        stb.Append(BuildContent)
      End If

      '
      ' End the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivEnd)

      '
      ' Add the body html to the page
      '
      Me.AddBody(stb.ToString)

      '
      ' Return the full page
      '
      Return Me.BuildPage()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "GetPagePlugin")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the HTML content
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildContent() As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table border='0' cellpadding='0' cellspacing='0' width='1000'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td width='1000' align='center' style='color:#FF0000; font-size:14pt; height:30px;'><strong><div id='divMessage'>&nbsp;</div></strong></td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", BuildTabs())
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildContent")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the jQuery Tabss
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BuildTabs() As String

    Try

      Dim stb As New StringBuilder
      Dim tabs As clsJQuery.jqTabs = New clsJQuery.jqTabs("oTabs", Me.PageName)
      Dim tab As New clsJQuery.Tab

      tabs.postOnTabClick = True

      tab.tabTitle = "Status"
      tab.tabDIVID = "tabStatus"
      tab.tabContent = "<div id='divStatus'>" & BuildTabStatus() & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Options"
      tab.tabDIVID = "tabOptions"
      tab.tabContent = "<div id='divOptions'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "SMTP Profiles"
      tab.tabDIVID = "tabAccounts"
      tab.tabContent = "<div id='divAccounts'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Delivery Logs"
      tab.tabDIVID = "tabDeliveryLogs"
      tab.tabContent = "<div id='divDeliveryLogs'></div>"
      tabs.tabs.Add(tab)

      Return tabs.Build

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Status Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabStatus(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine(clsPageBuilder.FormStart("frmStatus", "frmStatus", "Post"))

      stb.AppendLine("<div>")
      stb.AppendLine("<table>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Plug-In Status </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Name:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", IFACE_NAME)
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Status:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", "OK")
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Version:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", HSPI.Version)
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> SMTP Status </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Queue&nbsp;Count:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", Convert.ToInt32(GetQueueCount()).ToString("N0"))
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Database&nbsp;Size:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", GetDatabaseSize("DBConnectionMain"))
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table>")
      stb.AppendLine("</div>")

      stb.AppendLine(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divStatus", stb.ToString)

      Return stb.ToString
    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabStatus")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabOptions(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder
      Dim itemValues As Array
      Dim itemNames As Array

      stb.Append(clsPageBuilder.FormStart("frmOptions", "frmOptions", "Post"))

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      '
      ' SMTP Delivery Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>SMTP Delivery Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' SMTP Delivery Options (Maximum Delivery Attempts)
      '   
      Dim selMaxDeliveryAttempts As New clsJQuery.jqDropList("selMaxDeliveryAttempts", Me.PageName, False)
      selMaxDeliveryAttempts.id = "selMaxDeliveryAttempts"
      selMaxDeliveryAttempts.toolTip = "Specify the maximum number of delivery attempts before marking a message undeliverable."

      Dim strRetention As String = GetSetting("Options", "MaxDeliveryAttempts", "10")
      For index As Integer = 2 To 60 Step 1
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0}", index.ToString)
        selMaxDeliveryAttempts.AddItem(desc, value, index.ToString = strRetention)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'style='width: 20%'>Maximum&nbsp;Delivery&nbsp;Attempts:</td>")
      stb.AppendFormat("  <td class='tablecell'>{0} Time(s)</td>{1}", selMaxDeliveryAttempts.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' SMTP Delivery Options (Queue Runner)
      '   
      Dim selQueueRunner As New clsJQuery.jqDropList("selQueueRunner", Me.PageName, False)
      selQueueRunner.id = "selQueueRunner"
      selQueueRunner.toolTip = "Specify how often to retry deferred messasges."

      Dim txtQueueRunner As String = GetSetting("Options", "QueueRunnerFrequency", "2")
      For index As Integer = 1 To 30 Step 1
        Dim value As String = index.ToString
        Dim desc As String = String.Format("{0}", index.ToString)
        selQueueRunner.AddItem(desc, value, index.ToString = txtQueueRunner)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Queue&nbsp;Runner&nbsp;Frequency:</td>")
      stb.AppendFormat("  <td class='tablecell'>{0} Minute(s)</td>{1}", selQueueRunner.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Web Page Access
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Web Page Access</td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Authorized User Roles</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", BuildWebPageAccessCheckBoxes, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Application Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Application Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' Application Logging Level
      '
      Dim selLogLevel As New clsJQuery.jqDropList("selLogLevel", Me.PageName, False)
      selLogLevel.id = "selLogLevel"
      selLogLevel.toolTip = "Specifies the plug-in logging level."

      itemValues = System.Enum.GetValues(GetType(LogLevel))
      itemNames = System.Enum.GetNames(GetType(LogLevel))

      For i As Integer = 0 To itemNames.Length - 1
        Dim itemSelected As Boolean = IIf(gLogLevel = itemValues(i), True, False)
        selLogLevel.AddItem(itemNames(i), itemValues(i), itemSelected)
      Next
      selLogLevel.autoPostBack = True

      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", "Logging&nbsp;Level", vbCrLf)
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selLogLevel.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divOptions", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabOptions")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Accounts Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabAccounts(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmAccounts", "frmAccounts", "Post"))

      stb.AppendLine("<div id='divAccountsTable'>")
      stb.AppendLine(BuildAccountsTable())
      stb.AppendLine("</div>")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divAccounts", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabNetCamDevices")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Accounts table
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildAccountsTable(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table id='table_accounts' class='display compact' style='width:100%'>")
      stb.AppendFormat("<caption class='tableheader'>{0}</caption>{1}", "SMTP Accounts", vbCrLf)
      stb.AppendLine(" <thead>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>SMTP Server</th>")
      stb.AppendLine("   <th>SMTP Port</th>")
      stb.AppendLine("   <th>Use SSL</th>")
      stb.AppendLine("   <th>Sender E-mail Address</th>")
      stb.AppendLine("   <th>User Id</th>")
      stb.AppendLine("   <th>User Password</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </thead>")
      stb.AppendLine("<tbody>")

      Using MyDataTable As DataTable = GetSmtpProfiles()

        If MyDataTable.Columns.Contains("smtp_id") Then

          For Each row As DataRow In MyDataTable.Rows
            Dim smtp_id As Integer = row("smtp_id")
            Dim smtp_server As String = row("smtp_server")
            Dim smtp_port As String = row("smtp_port")
            Dim smtp_ssl As String = row("smtp_ssl")
            Dim auth_user As String = row("auth_user")
            Dim auth_pass As String = row("auth_pass")
            Dim mail_from As String = row("mail_from")

            stb.AppendFormat("  <tr id='{0}'>", smtp_id.ToString)
            stb.AppendFormat("   <td>{0}</td>{1}", smtp_server, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", smtp_port, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", smtp_ssl, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", mail_from, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", auth_user, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", "", vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", "", vbCrLf)
            stb.AppendLine("  </tr>")

          Next
        End If

      End Using

      stb.AppendLine("</tbody>")
      stb.AppendLine(" <tfoot>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>SMTP Server</th>")
      stb.AppendLine("   <th>SMTP Port</th>")
      stb.AppendLine("   <th>Use SSL</th>")
      stb.AppendLine("   <th>Sender E-mail Address</th>")
      stb.AppendLine("   <th>User Id</th>")
      stb.AppendLine("   <th>User Password</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </tfoot>")
      stb.AppendLine("</table>")

      Dim strInfo As String = "Edit the SMTP Accounts using the action links."
      Dim strHint As String = "The Sender E-mail Address must match the SMTP Account."

      stb.AppendLine(" <div>&nbsp;</div>")
      stb.AppendLine(" <p>")
      stb.AppendFormat("<img alt='Info' src='/images/hspi_ultrasmtp3/ico_info.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strInfo)
      stb.AppendFormat("<img alt='Hint' src='/images/hspi_ultrasmtp3/ico_hint.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strHint)
      stb.AppendLine(" </p>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildAccountsTable")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Delivery Logs Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabDeliveryLogs(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmDeliveryLogs", "frmDeliveryLogs", "Post"))

      stb.AppendLine("<div id='accordion-log'>")
      stb.AppendLine(" <span class='accordion-toggle'>Delivery Log Search</span>")
      stb.AppendLine(" <div class='accordion-content default'>")
      stb.AppendLine("  <table border='1' width='100%' style='border-collapse: collapse'>")
      stb.AppendLine("   <tbody>")
      stb.AppendLine("    <tr>")
      stb.AppendLine("     <td>")

      '
      ' Build Filter Field
      '
      Dim selFilterField As New clsJQuery.jqDropList("selFilterField", Me.PageName, True)
      selFilterField.id = "selFilterField"
      selFilterField.toolTip = ""

      Dim strFilterField As String = GetSetting("WebPage", "FilterField", "queue_id")
      selFilterField.AddItem("Queue Id", "queue_id", "queue_id" = strFilterField)
      stb.AppendLine(selFilterField.Build())

      '
      ' Build Filter Compare
      '
      Dim selFilterCompare As New clsJQuery.jqDropList("selFilterCompare", Me.PageName, True)
      selFilterCompare.id = "selFilterCompare"
      selFilterCompare.toolTip = ""

      Dim strFilterCompare As String = GetSetting("WebPage", "FilterCompare", "is")
      selFilterCompare.AddItem("Is", "is", "is" = strFilterCompare)
      selFilterCompare.AddItem("Starts With", "starts with", "starts with" = strFilterCompare)
      selFilterCompare.AddItem("Ends With", "ends with", "ends with" = strFilterCompare)
      selFilterCompare.AddItem("Contains", "contains", "contains" = strFilterCompare)
      stb.AppendLine(selFilterCompare.Build())

      '
      ' Build Filter Value
      '
      Dim tbFilterValue As New clsJQuery.jqTextBox("txtFilterValue", "text", "", PageName, 30, True)
      tbFilterValue.id = "txtFilterValue"
      tbFilterValue.toolTip = ""
      tbFilterValue.editable = True
      stb.AppendLine(tbFilterValue.Build())

      '
      ' Build Sort Order
      '
      Dim selSortOrder As New clsJQuery.jqDropList("selSortOrder", Me.PageName, True)
      selSortOrder.id = "selSortOrder"
      selSortOrder.toolTip = ""

      Dim strSortOrder As String = GetSetting("Webpage", "SortOrder", "desc")
      selSortOrder.AddItem("Sort Ascending", "asc", "asc" = strSortOrder)
      selSortOrder.AddItem("Sort Descending", "desc", "desc" = strSortOrder)
      stb.AppendLine(selSortOrder.Build())

      Dim btnFilter As New clsJQuery.jqButton("btnFilter", "Filter", Me.PageName, True)
      stb.AppendLine(btnFilter.Build())

      stb.AppendLine("     </td>")
      stb.AppendLine("    </tr>")

      Dim tbBeginDate As New clsJQuery.jqTextBox("txtBeginDate", "text", Date.Now.ToLongDateString, PageName, 30, True)
      tbBeginDate.id = "txtBeginDate"
      tbBeginDate.toolTip = ""
      tbBeginDate.editable = True

      Dim selBeginTime As New clsJQuery.jqDropList("selBeginTime", Me.PageName, True)
      selBeginTime.id = "selBeginTime"
      selBeginTime.toolTip = ""
      For hour As Integer = 0 To 23
        For min As Integer = 0 To 45 Step 15
          Dim time As String = String.Format("{0}:{1}", hour.ToString.PadLeft(2, "0"), min.ToString.PadLeft(2, "0"))
          selBeginTime.AddItem(time, time, time = "00:00")
        Next
      Next

      Dim tbEndDate As New clsJQuery.jqTextBox("txtEndDate", "text", Date.Now.ToLongDateString, PageName, 30, True)
      tbEndDate.id = "txtEndDate"
      tbEndDate.toolTip = ""
      tbEndDate.editable = True

      Dim selEndTime As New clsJQuery.jqDropList("selEndTime", Me.PageName, False)
      selEndTime.id = "selEndTime"
      selEndTime.toolTip = ""
      For hour As Integer = 0 To 23
        For min As Integer = 0 To 45 Step 15
          Dim time As String = String.Format("{0}:{1}", hour.ToString.PadLeft(2, "0"), min.ToString.PadLeft(2, "0"))
          selEndTime.AddItem(time, time, time = "00:00")
        Next
        selEndTime.AddItem("23:59", "23:59", True)
      Next

      stb.AppendLine("    <tr>")
      stb.AppendLine("     <td>")
      stb.AppendLine("    <strong>From:</strong>&nbsp;" & tbBeginDate.Build & "&nbsp;at&nbsp;" & selBeginTime.Build & "&nbsp;")
      stb.AppendLine("    <strong>To:</strong>&nbsp;" & tbEndDate.Build & "&nbsp;at&nbsp;" & selEndTime.Build & "&nbsp;")
      stb.AppendLine("     </td>")
      stb.AppendLine("    </tr>")
      stb.AppendLine("   </tbody>")
      stb.AppendLine(" </table>")

      stb.AppendLine(" </div>")
      stb.AppendLine("</div>")

      stb.Append(clsPageBuilder.FormEnd())

      stb.AppendLine("<div class='clear'><br /></div>")
      stb.AppendLine("<div id='divDeliveryLogTable'>")
      stb.AppendLine(BuildDeliveryLogTable(strFilterField, strFilterCompare, "", strSortOrder, Date.Now.ToLongDateString, "00:00", Date.Now.ToLongDateString, "23:59"))
      stb.AppendLine("</div>")

      If Rebuilding Then Me.divToUpdate.Add("divDeliveryLogs", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabLogs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Delivery Log Table
  ''' </summary>
  ''' <param name="filterField"></param>
  ''' <param name="filterCompare"></param>
  ''' <param name="filterValue"></param>
  ''' <param name="sortOrder"></param>
  ''' <param name="beginDate"></param>
  ''' <param name="beginTime"></param>
  ''' <param name="endDate"></param>
  ''' <param name="endTime"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildDeliveryLogTable(ByVal filterField As String, _
                                 ByVal filterCompare As String, _
                                 ByVal filterValue As String, _
                                 ByVal sortOrder As String, _
                                 ByVal beginDate As String, _
                                 ByVal beginTime As String, _
                                 ByVal endDate As String, _
                                 ByVal endTime As String) As String

    Try

      Dim stb As New StringBuilder

      Dim pageSize As Integer = 5000
      Dim pageCur As Integer = 1
      Dim pageCount As Integer = 1
      Dim recordCount As Integer = 0

      sortOrder = IIf(sortOrder.Length = 0, "asc", sortOrder)

      '
      ' Calculate the beginDate
      '            
      beginTime = IIf(beginTime = "", "00:00:01", beginTime & ":00")
      If IsDate(beginDate) = False Then
        beginDate = Date.Now.ToLongDateString
      Else
        beginDate = Date.Parse(beginDate).ToLongDateString
      End If
      beginDate = String.Format("{0} {1}", beginDate, beginTime)

      '
      ' Calculate the endDate
      '  
      endTime = IIf(endTime = "", "23:59:59", endTime & ":59")
      If IsDate(endDate) = False Then
        endDate = Date.Now.ToLongDateString
      Else
        endDate = Date.Parse(endDate).ToLongDateString
      End If
      endDate = String.Format("{0} {1}", endDate, endTime)

      'Dim dbFields As String = "log_id, datetime(log_ts, 'unixepoch', 'localtime') as log_ts, queue_id, smtp_id, smtp_server, smtp_status, smtp_result, mail_from, mail_to, mail_subject"

      Dim dbFields As String = "datetime(queue_ts, 'unixepoch', 'localtime') as queue_ts, " _
                             & "datetime(delivered_ts, 'unixepoch', 'localtime') as delivered_ts, " _
                             & "attempts, " _
                             & "queue_id, " _
                             & "last_status, " _
                             & "last_result, " _
                             & "mail_to, " _
                             & "mail_subject"

      Dim strSQL As String = BuildSmtpDeliveryLogSQL("tblSmtpQueue", dbFields, beginDate, endDate, filterField, filterCompare, filterValue, sortOrder)

      stb.AppendLine("<table id='table_log' class='display compact' style='width:100%'>")
      stb.AppendLine(" <thead>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Queue Id</th>")
      stb.AppendLine("   <th>Date Queued</th>")
      stb.AppendLine("   <th>Date Sent</th>")
      stb.AppendLine("   <th>Attempts</th>")
      stb.AppendLine("   <th>Mail To</th>")
      stb.AppendLine("   <th>Mail Subject</th>")
      stb.AppendLine("   <th>Delivery Status</th>")
      stb.AppendLine("   <th>Delivery Result</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </thead>")
      stb.AppendLine("<tbody>")

      Using MyDataTable As DataTable = ExecuteSQL(strSQL, recordCount, pageSize, pageCount, pageCur)

        Dim iRowIndex As Integer = 1
        If MyDataTable.Columns.Contains("queue_id") Then

          For Each row As DataRow In MyDataTable.Rows
            Dim queue_ts As String = row("queue_ts")
            Dim delivered_ts As String = row("delivered_ts")
            Dim queue_id As Integer = row("queue_id")
            Dim attempts As Integer = row("attempts")

            Dim mail_subject As String = row("mail_subject")
            Dim mail_to As String = row("mail_to")
            Dim last_status As String = row("last_status")
            Dim last_result As String = row("last_result")

            stb.AppendLine("  <tr>")
            stb.AppendFormat("   <td>{0}</td>{1}", queue_id.ToString, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", queue_ts, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", delivered_ts, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", attempts, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", mail_to, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", mail_subject, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", last_status, vbCrLf)
            stb.AppendFormat("   <td>{0}</td>{1}", last_result, vbCrLf)
            stb.AppendLine("  </tr>")

          Next
        End If

      End Using

      stb.AppendLine("</tbody>")
      stb.AppendLine(" <tfoot>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Queue Id</th>")
      stb.AppendLine("   <th>Date Queued</th>")
      stb.AppendLine("   <th>Date Sent</th>")
      stb.AppendLine("   <th>Attempts</th>")
      stb.AppendLine("   <th>Mail To</th>")
      stb.AppendLine("   <th>Mail Subject</th>")
      stb.AppendLine("   <th>Delivery Status</th>")
      stb.AppendLine("   <th>Delivery Result</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </tfoot>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildHSLogTable")
      Return "error - " & Err.Description
    End Try

  End Function


  ''' <summary>
  ''' Build the Web Page Access Checkbox List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildWebPageAccessCheckBoxes()

    Try

      Dim stb As New StringBuilder

      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      Dim cb1 As New clsJQuery.jqCheckBox("chkWebPageAccess_Guest", "Guest", Me.PageName, True, True)
      Dim cb2 As New clsJQuery.jqCheckBox("chkWebPageAccess_Admin", "Admin", Me.PageName, True, True)
      Dim cb3 As New clsJQuery.jqCheckBox("chkWebPageAccess_Normal", "Normal", Me.PageName, True, True)
      Dim cb4 As New clsJQuery.jqCheckBox("chkWebPageAccess_Local", "Local", Me.PageName, True, True)

      cb1.id = "WebPageAccess_Guest"
      cb1.checked = CBool(USER_ROLES_AUTHORIZED And USER_GUEST)

      cb2.id = "WebPageAccess_Admin"
      cb2.checked = CBool(USER_ROLES_AUTHORIZED And USER_ADMIN)
      cb2.enabled = False

      cb3.id = "WebPageAccess_Normal"
      cb3.checked = CBool(USER_ROLES_AUTHORIZED And USER_NORMAL)

      cb4.id = "WebPageAccess_Local"
      cb4.checked = CBool(USER_ROLES_AUTHORIZED And USER_LOCAL)

      stb.Append(clsPageBuilder.FormStart("frmWebPageAccess", "frmWebPageAccess", "Post"))

      stb.Append(cb1.Build())
      stb.Append(cb2.Build())
      stb.Append(cb3.Build())
      stb.Append(cb4.Build())

      stb.Append(clsPageBuilder.FormEnd())

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildWebPageAccessCheckBoxes")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

#Region "Page Processing"

  ''' <summary>
  ''' Post a message to this web page
  ''' </summary>
  ''' <param name="sMessage"></param>
  ''' <remarks></remarks>
  Sub PostMessage(ByVal sMessage As String)

    Try

      Me.divToUpdate.Add("divMessage", sMessage)

      Me.pageCommands.Add("starttimer", "")

      TimerEnabled = True

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' When a user clicks on any controls on one of your web pages, this function is then called with the post data. You can then parse the data and process as needed.
  ''' </summary>
  ''' <param name="page">The name of the page as registered with hs.RegisterLink or hs.RegisterConfigLink</param>
  ''' <param name="data">The post data</param>
  ''' <param name="user">The name of logged in user</param>
  ''' <param name="userRights">The rights of logged in user</param>
  ''' <returns>Any serialized data that needs to be passed back to the web page, generated by the clsPageBuilder class</returns>
  ''' <remarks></remarks>
  Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String

    Try

      WriteMessage("Entered postBackProc() function.", MessageType.Debug)

      Dim postData As NameValueCollection = HttpUtility.ParseQueryString(data)

      '
      ' Write debug to console
      '
      If gLogLevel >= MessageType.Debug Then
        For Each keyName As String In postData.AllKeys
          Console.WriteLine(String.Format("{0}={1}", keyName, postData(keyName)))
        Next
      End If

      '
      ' Process actions
      '
      Select Case postData("editor_action")
        Case "smtpaccount-create"
          Dim smtp_server As String = postData("data[smtp_server]").Trim
          Dim smtp_port As String = postData("data[smtp_port]").Trim
          Dim smtp_ssl As String = postData("data[smtp_ssl]").Trim
          Dim mail_from As String = postData("data[mail_from]").Trim
          Dim auth_user As String = postData("data[auth_user]").Trim
          Dim auth_pass As String = postData("data[auth_pass]").Trim

          If smtp_server.Length = 0 Then
            Return DatatableFieldError("smtp_server", "The SMTP server field is blank.  This is a required field.")
          ElseIf smtp_port.Length = 0 Then
            Return DatatableFieldError("smtp_port", "The SMTP port field is blank.  This is a required field.")
          ElseIf Regex.IsMatch(smtp_port, "^\d+$") = False Then
            Return DatatableFieldError("smtp_port", "The SMTP TCP port must be numeric.")
          ElseIf smtp_ssl.Length = 0 Then
            Return DatatableFieldError("smtp_ssl", "The SMTP use TLS field is blank.  This is a required field.")
          ElseIf Regex.IsMatch(smtp_ssl, "^\d+$") = False Then
            Return DatatableFieldError("smtp_ssl", "The SMTP use SSL field must be numeric.")
          ElseIf mail_from.Length = 0 Then
            Return DatatableFieldError("mail_from", "The SMTP mail from field is blank.  This is a required field.")
          End If

          Dim smtp_id As Integer = hspi_plugin.InsertSmtpProfile(smtp_server, smtp_port, smtp_ssl, auth_user, auth_pass, mail_from)
          If smtp_id = 0 Then
            Return DatatableError("Unable to add new SMTP account profile due to an unexpected error.")
          Else
            Return DatatableRowAccount(smtp_id, smtp_server, smtp_port, smtp_ssl, mail_from, auth_user, auth_pass)
          End If

        Case "smtpaccount-edit"
          Dim smtp_id As String = postData("id").Trim
          Dim smtp_server As String = postData("data[smtp_server]").Trim
          Dim smtp_port As String = postData("data[smtp_port]").Trim
          Dim smtp_ssl As String = postData("data[smtp_ssl]").Trim
          Dim mail_from As String = postData("data[mail_from]").Trim
          Dim auth_user As String = postData("data[auth_user]").Trim
          Dim auth_pass As String = postData("data[auth_pass]").Trim

          If smtp_server.Length = 0 Then
            Return DatatableFieldError("smtp_server", "The SMTP server field is blank.  This is a required field.")
          ElseIf smtp_port.Length = 0 Then
            Return DatatableFieldError("smtp_port", "The SMTP port field is blank.  This is a required field.")
          ElseIf Regex.IsMatch(smtp_port, "^\d+$") = False Then
            Return DatatableFieldError("smtp_port", "The SMTP TCP port must be numeric.")
          ElseIf smtp_ssl.Length = 0 Then
            Return DatatableFieldError("smtp_ssl", "The SMTP use TLS field is blank.  This is a required field.")
          ElseIf Regex.IsMatch(smtp_ssl, "^\d+$") = False Then
            Return DatatableFieldError("smtp_ssl", "The SMTP use SSL field must be numeric.")
          ElseIf mail_from.Length = 0 Then
            Return DatatableFieldError("mail_from", "The SMTP mail from field is blank.  This is a required field.")
          End If

          Dim bSuccess As Boolean = hspi_plugin.UpdateSmtpProfile(smtp_id, smtp_server, smtp_port, smtp_ssl, auth_user, auth_pass, mail_from)
          If smtp_id = 0 Then
            Return DatatableError("Unable to edit SMTP account profile due to an unexpected error.")
          Else
            Return DatatableRowAccount(smtp_id, smtp_server, smtp_port, smtp_ssl, mail_from, auth_user, auth_pass)
          End If

        Case "smtpaccount-remove"
          Dim smtp_id As String = postData("id[]")
          Dim bSuccess As Boolean = hspi_plugin.DeleteSmtpProfile(smtp_id)

          If bSuccess = False Then
            Return DatatableError("Unable to delete the SMTP account profile to an unexpected error.")
          Else
            BuildTabAccounts(True)
            Me.pageCommands.Add("executefunction", "reDrawAccounts()")
            Return "{ }"
          End If

      End Select

      '
      ' Process the post data
      '
      Select Case postData("id")
        Case "tabStatus"
          BuildTabStatus(True)

        Case "tabOptions"
          BuildTabOptions(True)

        Case "tabAccounts"
          BuildTabAccounts(True)
          Me.pageCommands.Add("executefunction", "reDrawAccounts()")

        Case "tabDeliveryLogs"
          BuildTabDeliveryLogs(True)
          Me.pageCommands.Add("executefunction", "reDrawDeliveryLogs()")

        Case "btnFilter"
          Dim filterField As String = postData("selFilterField")
          Dim filterCompare As String = postData("selFilterCompare")
          Dim filterValue As String = postData("txtFilterValue")

          Dim sortOrder As String = postData("selSortOrder")
          Dim beginDate As String = postData("txtBeginDate")
          Dim beginTime As String = postData("selBeginTime")
          Dim endDate As String = postData("txtEndDate")
          Dim endTime As String = postData("selEndTime")

          Me.divToUpdate.Add("divDeliveryLogTable", BuildDeliveryLogTable(filterField, filterCompare, filterValue, sortOrder, beginDate, beginTime, endDate, endTime))
          Me.pageCommands.Add("executefunction", "reDrawDeliveryLogs()")


        Case "selMaxDeliveryAttempts"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "MaxDeliveryAttempts", value)

          PostMessage("The maximum delivery attempts option has been updated.")

        Case "selQueueRunner"
          Dim value As String = postData(postData("id"))
          SaveSetting("Options", "QueueRunnerFrequency", value)

          PostMessage("The queue runner frequency option has been updated.")

        Case "selLogLevel"
          gLogLevel = Int32.Parse(postData("selLogLevel"))
          hs.SaveINISetting("Options", "LogLevel", gLogLevel.ToString, gINIFile)

          PostMessage("The application logging level has been updated.")

        Case "WebPageAccess_Guest"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Guest") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_GUEST
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_GUEST
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Normal"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Normal") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_NORMAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_NORMAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Local"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Local") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_LOCAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_LOCAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "timer" ' This stops the timer and clears the message
          If TimerEnabled Then 'this handles the initial timer post that occurs immediately upon enabling the timer.
            TimerEnabled = False
          Else
            Me.pageCommands.Add("stoptimer", "")
            Me.divToUpdate.Add("divMessage", "&nbsp;")
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "postBackProc")
    End Try

    Return MyBase.postBackProc(page, data, user, userRights)

  End Function

  ''' <summary>
  ''' Returns the Datatable Error JSON
  ''' </summary>
  ''' <param name="errorString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DatatableError(ByVal errorString As String) As String

    Try
      Return String.Format("{{ ""error"": ""{0}"" }}", errorString)
    Catch pEx As Exception
      Return String.Format("{{ ""error"": ""{0}"" }}", pEx.Message)
    End Try

  End Function

  ''' <summary>
  ''' Returns the Datatable Field Error JSON
  ''' </summary>
  ''' <param name="fieldName"></param>
  ''' <param name="fieldError"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DatatableFieldError(fieldName As String, fieldError As String) As String

    Try
      Return String.Format("{{ ""fieldErrors"": [ {{""name"": ""{0}"",""status"": ""{1}""}} ] }}", fieldName, fieldError)
    Catch pEx As Exception
      Return String.Format("{{ ""fieldErrors"": [ {{""name"": ""{0}"",""status"": ""{1}""}} ] }}", fieldName, pEx.Message)
    End Try

  End Function

  ''' <summary>
  ''' Returns the Datatable Row JSON
  ''' </summary>
  ''' <param name="smtp_id"></param>
  ''' <param name="smtp_server"></param>
  ''' <param name="smtp_port"></param>
  ''' <param name="smtp_ssl"></param>
  ''' <param name="mail_from"></param>
  ''' <param name="auth_user"></param>
  ''' <param name="auth_pass"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DatatableRowAccount(ByVal smtp_id As String, _
                                       ByVal smtp_server As String, _
                                       ByVal smtp_port As String,
                                       ByVal smtp_ssl As String, _
                                       ByVal mail_from As String, _
                                       ByVal auth_user As String, _
                                       ByVal auth_pass As String) As String

    Try

      Dim sb As New StringBuilder
      sb.AppendLine("{")
      sb.AppendLine(" ""row"": { ")

      sb.AppendFormat(" ""{0}"": {1}, ", "DT_RowId", smtp_id)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "smtp_server", smtp_server)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "smtp_port", smtp_port)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "smtp_ssl", smtp_ssl)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "mail_from", mail_from)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "auth_user", auth_user)
      sb.AppendFormat(" ""{0}"": ""{1}"" ", "auth_pass", auth_pass)

      sb.AppendLine(" }")
      sb.AppendLine("}")

      Return sb.ToString

    Catch pEx As Exception
      Return "{ }"
    End Try

  End Function


#End Region

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the HTML Not Authorized web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function WebUserNotUnauthorized(LoggedInUser As String) As String

    Try

      Dim sb As New StringBuilder

      sb.AppendLine("<table border='0' cellpadding='2' cellspacing='2' width='575px'>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td nowrap>")
      sb.AppendLine("     <h4>The Web Page You Were Trying To Access Is Restricted To Authorized Users ONLY</h4>")
      sb.AppendLine("   </td>")
      sb.AppendLine("  </tr>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td>")
      sb.AppendLine("     <p>This page is displayed if the credentials passed to the web server do not match the ")
      sb.AppendLine("      credentials required to access this web page.</p>")
      sb.AppendFormat("     <p>If you know the <b>{0}</b> user should have access,", LoggedInUser)
      sb.AppendFormat("      then ask your <b>HomeSeer Administrator</b> to check the <b>{0}</b> plug-in options", IFACE_NAME)
      sb.AppendFormat("      page to make sure the roles assigned to the <b>{0}</b> user allow access to this", LoggedInUser)
      sb.AppendLine("        web page.</p>")
      sb.AppendLine("  </td>")
      sb.AppendLine(" </tr>")
      sb.AppendLine("</table>")

      Return sb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "WebUserNotUnauthorized")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

End Class

