<%@ Page Language="VB" MasterPageFile="PageMaster.master" AutoEventWireup="false" EnableEventValidation="false" viewStateEncryptionMode="Auto" ASPCOMPAT="TRUE" %>
<%@ MasterType VirtualPath="PageMaster.master" %>
<%@ Register TagPrefix="obout" Namespace="Obout.Grid" Assembly="obout_Grid_NET" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.Page" %> 
<%@ Import Namespace="Webapps.Utils" %>
<%@ Import Namespace="Ionic.Zip" %>

<script language="VB" runat="server">

    Dim OwnerDataSet As DataSet = Nothing
    Dim SessionVariable_Prefix As String = "UserList_"
    '************* Error logging Section ***********************
    Dim pageName As String
    Dim errLocation As String
    Dim errString As String



    '********************** Bread & Crumb *****************
    Protected Sub SiteMapPath1_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SiteMapNodeItemEventArgs)
        If e.Item.ItemType = SiteMapNodeItemType.Root OrElse (e.Item.ItemType = SiteMapNodeItemType.PathSeparator AndAlso e.Item.ItemIndex = 1) Then
            e.Item.Visible = False
        End If
    End Sub

    '************* SQL Section ***********************

    Dim strCommandBase As String = " SELECT * FROM [v_w_GetUserList] WHERE 1=1 "
    'Dim strCommandOrderBy As String = strCommandGroupBy & "  ORDER BY  ud.UserID desc"
    Dim strCommandOrderBy As String = "  ORDER BY  UserID desc"

    '************* Seach Section ***********************
    'All common variables and session vaiables should use prefix unique to the page (example CR_xxxx)
    'Search Base SQL Statement.  Can be the same as strCommandBase
    'Search Order by string.  Can be the same as strCommandOrderBy
    'Display search criteria
    Dim strDisplayItemName As String = "Users"
    Dim strDisplaySearchCriteria As String = "Displaying: All " + strDisplayItemName
    '---- Search Parameter #1 
    Dim SearchParam1ColumnName As String = "UserID"  ' name the textbox as tbSearch1, and the dropdown box as ddSearch1
    Dim SearchParam1ColumnNameAlt2 As String = "LastName"  ' name the textbox as tbSearch1, and the dropdown box as ddSearch1
    Dim SearchParam1ColumnNameAlt3 As String = "Email"  ' name the textbox as tbSearch1, and the dropdown box as ddSearch1
    Dim SearchParam1ColumnNameDefault As String = "UserID"  ' name the textbox as tbSearch1, and the dropdown box as ddSearch1
    Dim SearchParam1Value As String = "" ' use session to store Session("ClaimRegister_searchParam1Value")

    Protected Sub Page_init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        If Not CustomRoles.RolesForPageLoad() Then
            CustomRoles.TransferIfNotInRole(True)
            Response.End()
            Exit Sub
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        pageName = Request.RawUrl.ToString
        Try
            Master.SetCurrentMenuItem = System.IO.Path.GetFileName(Request.RawUrl.ToString)
        Catch ex As Exception
            Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
            Exit Sub
        End Try
        ' IF ACCESS IS GOOD DO NOTHING, OTHERWISE TRANSFER
        If CustomRoles.RolesForPageLoad() Then

            Dim strPreviousSearch As String = Session(SessionVariable_Prefix & "IsSearch")
            Dim strGridPageIndex As String = Session(SessionVariable_Prefix & "GridPageIndex")
            Dim strRecordPerPage As String = Session(SessionVariable_Prefix & "RecordsPerPage")
            If Not IsCallback Then
                RestoreSetColumnSortGrouping()
            Else
                StoreSetColumnSortGrouping()
            End If
            If Not IsPostBack Then  'Perform on FIRST page load only

                If strPreviousSearch <> "TRUE" Then 'need to read from session
                    SetDefaultSession()
                End If
                If strGridPageIndex = Nothing Then   'Set grid Page Index (for sticky pages)
                    Session(SessionVariable_Prefix & "GridPageIndex") = grid1.CurrentPageIndex
                Else
                    grid1.CurrentPageIndex = Int32.Parse(strGridPageIndex)
                End If
                If strRecordPerPage = Nothing Then   'Set grid Page Index (for sticky pages)
                    Session(SessionVariable_Prefix & "RecordsPerPage") = grid1.PageSize
                Else
                    grid1.CurrentPageIndex = Int32.Parse(strGridPageIndex)
                    grid1.PageSize = Int32.Parse(strRecordPerPage)
                End If
                CreateGrid()
                SetFromSession()
                If CustomRoles.IsInRole("R_Admin_Administrator") Then
                    dvAdd.Visible = True
                Else
                    dvAdd.Visible = False
                End If
            Else
                CreateGrid()
                Session(SessionVariable_Prefix & "GridPageIndex") = grid1.CurrentPageIndex    'Set the Session variable to the current grid page index (sticky grid pages)
                Session(SessionVariable_Prefix & "RecordsPerPage") = grid1.PageSize
            End If
            'the not-misc-roles version of misc roles
            Dim ds As DataSet = Nothing
            CustomRoles.GetData(ds)

        Else
            CustomRoles.TransferIfNotInRole()
            Exit Sub
        End If
    End Sub



    Protected Sub CreateGrid()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim myComm As New SqlCommand()
        Dim myReader As SqlDataReader = Nothing
        Dim searchStr As String = Session(SessionVariable_Prefix & "searchParam1Value")
        Try
            myComm.Connection = myConn
            myComm.CommandText = strCommandBase
            '   Dim strFilter As String = CommonUtilsv2.GetSupplierRecordsFilterByUser(Session("User"))
            '  If String.IsNullOrEmpty(strFilter) Then

            '   Else
            '       myComm.CommandText = myComm.CommandText & " AND (" & strFilter & ") "
            '  End If  
            'Dim UsereMail As String = CommonUtilsv2.GetCurrentUserEamil(Session("User"))
            'If (Not UsereMail.Contains("@jmsassoc.com")) Then
            '    myComm.CommandText = myComm.CommandText & " and Email  NOT LIKE '%@jmsassoc.com%' "
            'End If

            If CustomRoles.IsInRole("R_Admin_Administrator") = False Then
                myComm.CommandText = myComm.CommandText & " and UserID in (select User_ID from v_w_ClientUsers ) "
            End If

            If Session(SessionVariable_Prefix & "IsSearch") <> "TRUE" Then
                myComm.CommandText = myComm.CommandText & strCommandOrderBy
                Session(SessionVariable_Prefix & "IsSearch") = "FALSE"
                strDisplaySearchCriteria = "Displaying: All " + strDisplayItemName
                Session(SessionVariable_Prefix & "SearchLabel") = strDisplaySearchCriteria
            Else        'Has existing search
                If Session(SessionVariable_Prefix & "searchParam1Value") <> "" Then
                    Dim strSearchStr As String = "%" + searchStr + "%"

                    myComm.CommandText = myComm.CommandText & " AND " & Session(SessionVariable_Prefix & "searchParam1ColumnName") & " LIKE @SearchedParm1 "
                    myComm.Parameters.AddWithValue("@SearchedParm1", strSearchStr)

                End If

                myComm.CommandText = myComm.CommandText & strCommandOrderBy
            End If
            myConn.Open()
            myReader = myComm.ExecuteReader()
            grid1.DataSource = myReader
            grid1.DataBind()

        Catch ex As Exception
            errString = ex.Message
            errLocation = "CreateGrid()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myReader Is Nothing Then
                myReader.Close()
            End If
            myConn.Close()
            myConn.Dispose()
            myConn = Nothing
        End Try
    End Sub

    Sub RowDataBound(ByVal sender As Object, ByVal e As GridRowEventArgs)   'Set Row highlights as data is bound based on cell values

    End Sub

    Protected Sub btnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim htSearchParams As Hashtable = New Hashtable()
        Dim strWhereClause As StringBuilder = New StringBuilder()

        lbSearchApplied.Text = ""
        lbSearchApplied.ForeColor = Drawing.Color.Green
        strDisplaySearchCriteria = "Displaying: All " + strDisplayItemName
        If Not String.IsNullOrEmpty(tbSearch1.Text) Then
            If CommonUtilsv2.Validate(tbSearch1.Text, CommonUtilsv2.DataTypes.String, False, True, False, 100) = False Then
                SearchParam1Value = ""
            Else
                SearchParam1Value = tbSearch1.Text.Trim()
            End If
        End If


        'SEARCH
        'Dim startDate As String = ""
        'Dim endDate As String = ""

        If SearchParam1Value <> "" Then
            Session(SessionVariable_Prefix & "IsSearch") = "TRUE"
            strDisplaySearchCriteria = "Displaying: "
            If SearchParam1Value <> "" Then
                Session(SessionVariable_Prefix & "GridPageIndex") = 0
                grid1.CurrentPageIndex = 0
                'Determine Search by for ddSearch1
                Select Case ddSearch1.SelectedIndex
                    Case 0
                        strDisplaySearchCriteria += "User ID contains '" + SearchParam1Value + "'"
                        SearchParam1ColumnName = SearchParam1ColumnNameDefault
                    Case 1
                        strDisplaySearchCriteria += "Last Name contains '" + SearchParam1Value
                        SearchParam1ColumnName = SearchParam1ColumnNameAlt2
                    Case 2
                        strDisplaySearchCriteria += "Email contains '" + SearchParam1Value
                        SearchParam1ColumnName = SearchParam1ColumnNameAlt3
                    Case Else
                        ddSearch1.SelectedValue = "UserID"
                        strDisplaySearchCriteria = strDisplaySearchCriteria + "User ID contains '" + SearchParam1Value + "'"
                        ddSearch1.SelectedIndex = 0
                        SearchParam1ColumnName = SearchParam1ColumnNameDefault
                End Select

                strWhereClause.Append(" AND " & SearchParam1ColumnName & " LIKE @SearchParam1")
                htSearchParams.Add("@SearchParam1", "%" & SearchParam1Value & "%")

            Else
                Session(SessionVariable_Prefix & "ddSearch1Index") = 0
                SearchParam1ColumnName = SearchParam1ColumnNameDefault
            End If

        Else
            'User hit search, but with no criteria, use base search
            strDisplaySearchCriteria = "Displaying: All " + strDisplayItemName
            ddSearch1.SelectedIndex = 0
            Session(SessionVariable_Prefix & "IsSearch") = "FALSE"
        End If
        Session(SessionVariable_Prefix & "_SearchParams") = htSearchParams
        Session(SessionVariable_Prefix & "_WhereClause") = strWhereClause.ToString

        Session(SessionVariable_Prefix & "SearchLabel") = strDisplaySearchCriteria
        Session(SessionVariable_Prefix & "searchParam1ColumnName") = SearchParam1ColumnName
        SetSessionVariables()
        CreateGrid()
        SetFromSession()
    End Sub

    Private Sub RebindGrid(ByVal sender As Object, ByVal e As EventArgs)
        lbSearchApplied.Text = Session(SessionVariable_Prefix & "SearchLabel")
        CreateGrid()
        SetFromSession()
    End Sub

    Protected Sub btn_ViewAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_ViewAll.Click
        SetDefaultSession()
        CreateGrid()
        SetFromSession()
    End Sub

    Protected Sub SetSessionVariables()
        'Session(SessionVariable_Prefix & "SearchLabel") = lbSearchApplied.Text
        Session(SessionVariable_Prefix & "ddSearch1Index") = ddSearch1.SelectedIndex
        Session(SessionVariable_Prefix & "searchParam1Value") = tbSearch1.Text
        Session(SessionVariable_Prefix & "GridPageIndex") = grid1.CurrentPageIndex
        Session(SessionVariable_Prefix & "RecordsPerPage") = grid1.PageSize
    End Sub

    Protected Sub SetFromSession()
        Dim iTempInteger As Integer = Nothing
        Try
            lbSearchApplied.Text = Session(SessionVariable_Prefix & "SearchLabel")
            tbSearch1.Text = Session(SessionVariable_Prefix & "searchParam1Value")
            ddSearch1.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch1Index")
            If ddSearch1.Items.Count > 0 Then
                ddSearch1.Text = ddSearch1.Items(Session(SessionVariable_Prefix & "ddSearch1Index")).Text
            Else
                ddSearch1.Text = Session(SessionVariable_Prefix & "searchParam1Value")
            End If

            If Not String.IsNullOrEmpty(Session(SessionVariable_Prefix & "RecordsPerPage")) AndAlso Integer.TryParse(Session(SessionVariable_Prefix & "RecordsPerPage"), iTempInteger) = True Then
                If iTempInteger > 0 Then
                    grid1.PageSize = Session(SessionVariable_Prefix & "RecordsPerPage")
                End If
            End If
            If Not String.IsNullOrEmpty(Session(SessionVariable_Prefix & "GridPageIndex")) AndAlso Integer.TryParse(Session(SessionVariable_Prefix & "GridPageIndex"), iTempInteger) = True Then
                Try
                    grid1.CurrentPageIndex = Session(SessionVariable_Prefix & "GridPageIndex")
                Catch ex As Exception
                    grid1.CurrentPageIndex = 0
                    Session(SessionVariable_Prefix & "GridPageIndex") = 0
                End Try
            End If

        Catch ex As Exception
            errString = ex.Message
            errLocation = "SetFromSession()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try
    End Sub

    Protected Sub SetDefaultSession()
        Session(SessionVariable_Prefix & "GridPageIndex") = 0
        grid1.CurrentPageIndex = 0
        Session(SessionVariable_Prefix & "SearchLabel") = "Displaying: All " + strDisplayItemName
        Session(SessionVariable_Prefix & "ddSearch1Index") = 0
        Session(SessionVariable_Prefix & "IsSearch") = "FALSE"
        Session(SessionVariable_Prefix & "searchParam1Value") = ""
        Session(SessionVariable_Prefix & "_SearchParams") = Nothing
        Session(SessionVariable_Prefix & "_WhereClause") = ""
    End Sub

    Public Function GetCheckBox(ByVal strCheck As String) As String
        Dim ret As String = ""
        Dim strCheckValue As String = ""
        Dim intChecked As Integer = 0
        Try
            intChecked = Int32.Parse(strCheck)
        Catch ex As Exception
            intChecked = 0
        End Try
        If intChecked = 1 Then
            strCheckValue = "checked"
        End If

        ret = "<input type='checkbox' disabled='true' " & strCheckValue & " />"
        Return ret
    End Function

    Public Function GetCheckBoxInt(ByVal strCheck As String) As String

        Dim ret As String = "<img src=""Images/checkbox_unchecked.jpg"" alt=""0"" />"
        If Not String.IsNullOrEmpty(strCheck) Then
            If String.Compare("0", strCheck, True) = 0 Then
            Else
                ret = "<img src=""Images/checkbox_checked.jpg"" alt=""1"" />"
            End If
        Else

        End If
        Return ret
    End Function

    Public Function GetCheckBoxImage(ByVal strCheck As String) As String
        Dim strIsChecked As String = "<img src=""Images/checkbox_unchecked.jpg"" alt=""0"" />"
        If Not String.IsNullOrEmpty(strCheck) Then
            If String.Compare("FALSE", strCheck, True) = 0 Then
            Else
                strIsChecked = "<img src=""Images/checkbox_checked.jpg"" alt=""1"" />"

            End If
        Else

        End If
        Dim result As String = strIsChecked

        Return result

    End Function

    Private Sub gridSetColumnSort(ByVal sender As Object, ByVal e As EventArgs)
        Dim colCount As Integer = grid1.Columns.Count
        Dim i As Integer = 0
        Dim nWidth As Integer = 40 / (colCount - 4)
        Dim colnSort(1, colCount - 1) As String

        If Page.IsPostBack Or Page.IsCallback Then
            ' store sort settings
            For Each coloumn As Column In grid1.Columns
                colnSort(0, i) = coloumn.DataField
                colnSort(1, i) = coloumn.SortOrder
                i = i + 1
            Next
            Session(SessionVariable_Prefix & "ColumnSort") = colnSort
            If String.IsNullOrEmpty(grid1.GroupBy()) Then
                Session(SessionVariable_Prefix & "GroupBy") = ""
                Session(SessionVariable_Prefix & "GroupExpCollapse") = "False"
            Else
                Session(SessionVariable_Prefix & "GroupBy") = grid1.GroupBy()
                Session(SessionVariable_Prefix & "GroupExpCollapse") = grid1.ShowCollapsedGroups()
            End If
        Else
            If Not Session(SessionVariable_Prefix & "ColumnSort") Is Nothing Then
                colnSort = Session(SessionVariable_Prefix & "ColumnSort")
                ' store sort settings
                For Each coloumn As Column In grid1.Columns
                    If colnSort(0, i) = coloumn.DataField Then
                        coloumn.SortOrder = colnSort(1, i)
                    End If
                    i = i + 1
                Next
            End If
            If Not Session(SessionVariable_Prefix & "roupBy") Is Nothing Then
                grid1.GroupBy = Session(SessionVariable_Prefix & "GroupBy")
                If String.Compare("Tue", Session(SessionVariable_Prefix & "GroupExpCollapse"), True) = 0 Then
                    grid1.ShowCollapsedGroups = True
                Else
                    grid1.ShowCollapsedGroups = False
                End If
            End If
        End If
    End Sub


    Private Sub StoreSetColumnSortGrouping()
        Dim colCount As Integer = grid1.Columns.Count
        Dim i As Integer = 0
        Dim nWidth As Integer = 40 / (colCount - 4)
        Dim colnSort(1, colCount - 1) As String

        'If Page.IsPostBack Or Page.IsCallback Then
        If Page.IsCallback Then
            ' store sort settings
            For Each coloumn As Column In grid1.Columns
                colnSort(0, i) = coloumn.DataField
                colnSort(1, i) = coloumn.SortOrder
                i = i + 1
            Next
            Session(SessionVariable_Prefix & "ColumnSort") = colnSort
            If String.IsNullOrEmpty(grid1.GroupBy()) Then
                Session(SessionVariable_Prefix & "GroupBy") = ""
                Session(SessionVariable_Prefix & "GroupExpCollapse") = "False"
            Else
                Session(SessionVariable_Prefix & "GroupBy") = grid1.GroupBy()
                Session(SessionVariable_Prefix & "GroupExpCollapse") = grid1.ShowCollapsedGroups()
            End If
        ElseIf IsPostBack Then
        End If
    End Sub

    Private Sub RestoreSetColumnSortGrouping()
        Dim colCount As Integer = grid1.Columns.Count
        Dim i As Integer = 0
        Dim nWidth As Integer = 40 / (colCount - 4)
        Dim colnSort(1, colCount - 1) As String

        If Not Session(SessionVariable_Prefix & "ColumnSort") Is Nothing Then
            colnSort = Session(SessionVariable_Prefix & "ColumnSort")
            ' store sort settings
            For Each coloumn As Column In grid1.Columns
                If colnSort(0, i) = coloumn.DataField Then
                    coloumn.SortOrder = colnSort(1, i)
                End If
                i = i + 1
            Next
        End If
        If Not Session(SessionVariable_Prefix & "GroupBy") Is Nothing Then
            grid1.GroupBy = Session(SessionVariable_Prefix & "GroupBy")
            If String.Compare("True", Session(SessionVariable_Prefix & "GroupExpCollapse"), True) = 0 Then
                grid1.ShowCollapsedGroups = True
            Else
                grid1.ShowCollapsedGroups = False
            End If
        End If
    End Sub

</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
    <script type="text/javascript">
        function exportToExcel() {
            grid1.exportToExcel();
        }
    </script>
    <div align="center">
        <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0" class="reg_table_style">
            <tr>
                <td colspan="6" align="center"></td>
            </tr>
            <%--<tr>
                <td colspan="6" align="left" class="body_title">User List</td>
            </tr>--%>
            <tr><td colspan="6">&nbsp;</td></tr>
            <tr class="body_search">
                <td align="right" width="12%">Search By:</td>
                <td width="15%">
                    <asp:DropDownList ID="ddSearch1" runat="server" Width="100%">
                        <asp:ListItem Value="UserID">User ID</asp:ListItem>
                        <asp:ListItem Value="LastName">Last Name</asp:ListItem>
                        <asp:ListItem Value="Email">Email</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td align="left" width="35%" colspan="1">&nbsp;includes:&nbsp;&nbsp;<asp:TextBox ID="tbSearch1" runat="server" MaxLength="30" Wrap="False"></asp:TextBox></td>
                <td align="right" valign="top" width="3%" class="body_search" colspan="2"></td>
                <td width="21%" valign="top">
                   
                </td>
            </tr>

            <tr>
                <td colspan="5"></td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td colspan="2" align="left"></td>
                <td  colspan="2" align="right" valign="top" class="body_search">&nbsp;</td>
                <td colspan="2" valign="top" align="right">
                    <asp:Button ID="btnSearch" runat="server" Text="Search" class="Submit_button"/>
                    <asp:Button ID="btn_ViewAll" runat="server" Text="View All" class="Submit_button"/>
                </td>
            </tr>
            <tr>
                <td colspan="5" align="left">
                    <asp:Label ID="lbSearchApplied" runat="server" ForeColor="#15723b" Font-Bold="true" Visible="True" CssClass="body_search"></asp:Label><br />
                </td>
                <td>
                    &nbsp;</td>
            </tr>
            <tr>
                <td width="100%" align="left" valign="top" colspan="6">
                    <table width="100%" class="body_search">
                        <tr>
                            <td width="10%"></td>
                            <td align="right" width="65%" colspan="3">
                                <asp:Label ID="ExportLabel" runat="server" Visible="false" ForeColor="DarkRed" Text="Please narrow down the search as the number of filtered records exceeds the export limit." />
                            </td>
                            <td align="right" width="25%" colspan="3">
                                <asp:LinkButton ID="btnExportCSV" runat="server" Visible="false" Text="Export to CSV" />&nbsp;&nbsp;
                                                <asp:LinkButton ID="btnExportExcel" runat="server" Visible="false" OnClientClick="exportToExcel(); return false;" Text="Export to Excel" />&nbsp;&nbsp;
                                                <div id="dvAdd" runat="server">
                                                    <a href="ManageUsers-Add.aspx">Add New User</a>
                                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td width="100%" align="left" valign="top" colspan="6">
                    <obout:Grid id="grid1" runat="server" CallbackMode="true" Serialize="false" AutoGenerateColumns="false"
		                        PageSizeOptions="1,5,10,15,20,25,30,35,40,45,50,100" PageSize="10"
		                        Width="100%" FolderStyle="styles/grand_graydark" EnableRecordHover="true"
		                        AllowRecordSelection="true" AllowMultiRecordSelection="true" KeepSelectedRecords="true"
		                        AllowAddingRecords="false" AllowFiltering="false" ShowLoadingMessage="true" AllowSorting="true"
		                        AllowGrouping="false" ShowGroupsInfo="false" ShowCollapsedGroups="false" ShowMultiPageGroupsInfo="false"
		                        ShowColumnsFooter="false" ShowGroupFooter="false" OnRowDataBound="RowDataBound" OnRebind="RebindGrid"
                                >
		                        <Columns>
			                        <obout:Column ID="Column01" DataField="ID" HeaderText="ID" Align="right" Width="5%" Visible="true" Wrap="false" AllowGroupBy="false" runat="server" HeaderAlign="center" ReadOnly="true" />
                                    <obout:Column ID="Column02" DataField="UserID" HeaderText="UserID" TemplateId="tmplUserDetail" Align="left" Width="10%" Visible="true" Wrap="false" AllowGroupBy="false" runat="server" HeaderAlign="center" ExportAsText="true" ReadOnly="true" />
			                        <obout:Column ID="Column03" DataField="FirstName" HeaderText="FirstName" Align="left" ExportAsText="false" Width="10%" Visible="true" Wrap="true" AllowGroupBy="false" runat="server" HeaderAlign="center" readonly="true"/>
                                    <obout:Column ID="Column12" DataField="LastName" HeaderText="LastName" Align="left" ExportAsText="false" Width="10%" Visible="true" Wrap="true" AllowGroupBy="false" runat="server" HeaderAlign="center" readonly="true"/>
                                    <obout:Column ID="Column05" DataField="Company"  HeaderText="Company" Align="left" ExportAsText="false" Width="10%" Visible="true" Wrap="true" runat="server" HeaderAlign="center" readonly="true"/>
                                    <obout:Column ID="Column06" DataField="Email" HeaderText="Email" Align="left"  Width="10%" Visible="true" Wrap="true"  DataFormatString="{0:C2}" AllowGroupBy="false" runat="server" HeaderAlign="center" ExportAsText="true" readonly="true"/>	
                                    <obout:Column ID="Column07" DataField="LastLogin" HeaderText="LastLogin" Align="left"  Width="10%" Visible="true" Wrap="true" AllowGroupBy="false" runat="server" DataFormatString="{0:MM/dd/yyyy}" HeaderAlign="center" ExportAsText="true" readonly="true"/>
                                    <obout:Column ID="Column16" DataField="Disabled" HeaderText="Disaled" Align="left" ExportAsText="false" Width="5%" Visible="true" Wrap="true" AllowGroupBy="false" runat="server" HeaderAlign="center" readonly="true"/>
                                    <obout:Column ID="c05" HeaderText="Last 7" DataField="Last7" Align="center" Width="5%" Visible="true" Wrap="true" AllowGroupBy="true"  runat="server" HeaderAlign="center" />   
                                    <obout:Column ID="c06" HeaderText="Last 30" DataField="Last30" Align="center" Width="5%" Visible="true" Wrap="true" AllowGroupBy="true"  runat="server" HeaderAlign="center" />         
                                    <obout:Column ID="c07" HeaderText="Last 60" DataField="Last60" Align="center" Width="5%" Visible="true" Wrap="true" AllowGroupBy="true"  runat="server" HeaderAlign="center" />
                                    <obout:Column ID="c08" HeaderText="Last 90" DataField="Last90" Align="center" Width="5%" Visible="true" Wrap="true" AllowGroupBy="true"  runat="server" HeaderAlign="center" />
                                    <obout:Column ID="c09" HeaderText="Total" DataField="TotalLogons" Align="center" Width="5%" Visible="true" Wrap="true" AllowGroupBy="true"  runat="server" HeaderAlign="center" /> 
                                    <obout:Column ID="Column4" HeaderText="RoleCount" DataField="RoleCount" Align="center" Width="5%" Visible="true" Wrap="true" AllowGroupBy="true"  runat="server" HeaderAlign="center" /> 
                                </Columns>
		                  <Templates>
                            <obout:GridTemplate runat="server" ID="tmplUserDetail">
                                <Template>
                                    <a class="a.ob_gAL" href="ManageUsers-Edit.aspx?UID=<%#Container.DataItem("UserID")%>"><%# Container.Value()%></a>
                                </Template>
                            </obout:GridTemplate>

                        </Templates>
                    </obout:Grid>
                </td>
            </tr>
        </table>        
    </div>
</asp:Content>
