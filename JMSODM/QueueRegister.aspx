<%@ Page Language="VB" MasterPageFile="PageMaster.master" AutoEventWireup="false" EnableEventValidation="false" ViewStateEncryptionMode="Auto" AspCompat="TRUE" %>

<%@ MasterType VirtualPath="PageMaster.master" %>
<%@ Register TagPrefix="obout" Namespace="Obout.Grid" Assembly="obout_Grid_NET" %>
<%--<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>--%>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%--<%@ Import Namespace="System.Web.UI.Page" %> --%>
<%@ Import Namespace="Webapps.Utils" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="Ionic.Zip" %>
<%@ Import Namespace="System.Web.Services" %>

<script language="VB" runat="server">

    Dim SessionVariable_Prefix As String = "QueueRegister_"
    '************* Error logging Section ***********************
    Dim pageName As String
    Dim errLocation As String
    Dim errString As String
    Dim StrTextCount As String
    '********************** Bread & Crumb *****************
    Protected Sub SiteMapPath1_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SiteMapNodeItemEventArgs)
        If e.Item.ItemType = SiteMapNodeItemType.Root OrElse (e.Item.ItemType = SiteMapNodeItemType.PathSeparator AndAlso e.Item.ItemIndex = 1) Then
            e.Item.Visible = False
        End If
    End Sub

    '************* SQL Section ***********************      
    Dim strCommandBase As String = " SELECT *,(select count(*) from tbl_Web_Attachments atta where atta.DCN=a.DCN and RecordStatus=1) attachmentcount FROM [v_w_QueueRegister] as a WHERE 1 = 1 "

    Dim strCommandOrderBy As String = " ORDER BY QueueID ASC,DCN DESC  "
    'Dim strCommandOrderBy As String = " ORDER BY CR.ClaimNumber ASC "

    Dim strExportSQL As String = " SELECT * FROM [v_w_QueueRegisterExport] WHERE 1 = 1 "

    '************* Seach Section ***********************
    'All common variables and session vaiables should use prefix unique to the page (example CR_xxxx)

    'Display search criteria
    Dim strDisplayItemName As String = "Documents"
    Dim strDisplaySearchCriteria As String = "Displaying: All " + strDisplayItemName
    '---- Search Parameter #1 
    Dim CR_searchParam1ColumnName As String = "Index1"  ' name the textbox as tbSearch1, and the dropdown box as ddSearch1
    Dim CR_searchParam1ColumnNameAlt2 As String = "Index2"  ' name the textbox as tbSearch1, and the dropdown box as ddSearch1
    Dim CR_searchParam1ColumnNameAlt3 As String = "Index3" ' name the textbox as tbSearch1, and the dropdown box as ddSearch1
    Dim CR_searchParam1ColumnNameAlt4 As String = "Index4" ' name the textbox as tbSearch1, and the dropdown box as ddSearch1
    Dim CR_searchParam1ColumnNameAlt5 As String = "Index5" ' name the textbox as tbSearch1, and the dropdown box as ddSearch1
    Dim CR_searchParam1ColumnNameAlt6 As String = "Index6" ' name the textbox as tbSearch1, and the dropdown box as ddSearch1
    Dim CR_searchParam1ColumnNameAlt7 As String = "Index7" ' name the textbox as tbSearch1, and the dropdown box as ddSearch1
    Dim CR_searchParam1ColumnNameDefault As String = "Index1"
    Dim CR_searchParam1Value As String = "" ' use session to store Session(SessionVariable_Prefix & "searchParam1Value")

    '---- Search Parameter #2 
    Dim CR_searchParam2ColumnName As String = "DocTypeID"
    Dim CR_searchParam2PrettyName As String = "Doc Type"
    Dim CR_searchParam2Value As String = ""
    Dim CR_searchParam2Default As String = "0"
    Dim CR_searchParam2ValueDesc As String = "DocTypeID"
    '---- Search Parameter #3 
    Dim CR_searchParam3ColumnName As String = "DocStatusID"
    Dim CR_searchParam3PrettyName As String = "Doc Status"
    Dim CR_searchParam3Value As String = ""
    Dim CR_searchParam3Default As String = "0"
    Dim CR_searchParam3ValueDesc As String = "DocStatusID"
    '---- Search Parameter #4 

    Dim CR_searchParam4ColumnName As String = "PriorityID"
    Dim CR_searchParam4PrettyName As String = "Priority"
    Dim CR_searchParam4Value As String = ""
    Dim CR_searchParam4Default As String = "0"
    Dim CR_searchParam4ValueDesc As String = "PriorityID"

    '---- Search Parameter #5

    Dim CR_searchParam5ColumnName As String = "OwnerID"
    Dim CR_searchParam5PrettyName As String = "Owner"
    Dim CR_searchParam5Value As String = ""
    Dim CR_searchParam5Default As String = "0"
    Dim CR_searchParam5ValueDesc As String = "OwnerID"



    Dim IndexNum As String = ""
    Public IsDeleteGroup As Boolean = False
    Public Isprovdier As Integer = 0 '0 no provider  1 is provider 2 is provider but it is complete

    Protected Sub Page_init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        If Not CustomRoles.RolesForPageLoad() Then

            CustomRoles.TransferIfNotInRole(True)
            Response.End()
            Exit Sub
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        FileUpload1.Attributes.Add("onchange", "setFileName('" + FileUpload1.ClientID + "','" + tbNewFile.ClientID + "')")
        If String.IsNullOrEmpty(FileUpload1.FileName.ToString) Then
            tbNewFile.Text = ""
        End If

        pageName = Request.RawUrl.ToString
        Try
            Master.SetCurrentMenuItem = System.IO.Path.GetFileName(Request.RawUrl.ToString)
        Catch ex As Exception
            Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
            Exit Sub
        End Try
        Dim anID As String = Request.QueryString.Get("qid")
        ' IF ACCESS IS GOOD DO NOTHING, OTHERWISE TRANSFER
        If CustomRoles.RolesForPageLoad() Then
            If CommonUtilsv2.Validate(anID, CommonUtilsv2.DataTypes.Int, True, True, True) Then
            Else
                anID = -9999
                Exit Sub
            End If
            Session(SessionVariable_Prefix & "queueid") = anID

            Dim strPreviousSearch As String = Session(SessionVariable_Prefix & "IsSearch")
            Dim strAccountRegisterGridPageIndex As String = Session(SessionVariable_Prefix & "GridPageIndex")
            SetPrivoderType()
            If Not IsPostBack Then  'Perform on FIRST page load only
                If strPreviousSearch <> "TRUE" Then 'need to read from session
                End If
                If strAccountRegisterGridPageIndex = Nothing Then   'Set grid Page Index (for sticky pages)
                    Session(SessionVariable_Prefix & "GridPageIndex") = gridQueue.CurrentPageIndex
                Else
                    gridQueue.CurrentPageIndex = Int32.Parse(strAccountRegisterGridPageIndex)
                End If
                'CheckForComments()
                LoadSearchDropDownLists()
                GetDynamicIndexs()
                CreateGrid()
                SetFromSession()

                lbSubmitComment.Text = ""
                lbSubmitComment.ForeColor = Drawing.Color.Red
                lbSubmitComment.Visible = False
            Else
                GetDynamicIndexs()
                CreateGrid()
                Session(SessionVariable_Prefix & "GridPageIndex") = gridQueue.CurrentPageIndex    'Set the Session variable to the current grid page index (sticky grid pages)
                ' SetFromSession()
            End If

            If CustomRoles.IsInRole("R_Upload_Document") Then
                dvUpload.Visible = True
            Else
                dvUpload.Visible = False
            End If

            If CustomRoles.IsInRole("R_Bulk_Move") Then
                dvBulkMove.Visible = True
            Else
                dvBulkMove.Visible = False
            End If
            If CustomRoles.IsInRole("R_Bulk_Assign") Then
                dvBulkAssign.Visible = True
            Else
                dvBulkAssign.Visible = False
            End If

            If Not IsPostBack Then
                ddQueue.SelectedValue = Session(SessionVariable_Prefix & "queueid")
                'If String.IsNullOrEmpty(Session("AllowedUploadFileSize")) Or String.IsNullOrEmpty(Session("AllowedUploadFileSizeInMB")) Then
                '    Dim lFSizeLimit As Long = CommonUtilsv2.GetAllowedUploadFileSize()
                '    Dim sFSizeLimit As Single = lFSizeLimit / (1024 * 1024.0)
                '    Dim strDocSizeLimit As String = FormatNumber(sFSizeLimit, 2, TriState.True, TriState.False, TriState.True) & "MB"
                '    Session("AllowedUploadFileSize") = lFSizeLimit
                '    Session("AllowedUploadFileSizeInMB") = strDocSizeLimit
                'End If
            End If
        Else
            CustomRoles.TransferIfNotInRole()
            Exit Sub
        End If
    End Sub

    Private Sub SetPrivoderType()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim strSQL As String = " select ID from [dbo].[tbl_Web_WorkGroups] where ID=(select GroupID from [dbo].[tbl_Web_Queues] where ID=@Qid) "
        Dim params4 As SqlParameter() = {New SqlParameter("@Qid", Session(SessionVariable_Prefix & "queueid"))}
        Dim GroupId = StrHelp.GetInt(CommonUtilsv2.RunScalarQuery(dbKey, strSQL, CommandType.Text, params4))
        params4 = Nothing
        If (GroupId = 1) Then
            If (StrHelp.GetInt(Session(SessionVariable_Prefix & "queueid"))) = 6 Then ' is complete
                Isprovdier = 2
            Else
                Isprovdier = 1
            End If
        End If

    End Sub

    Private Sub LoadSearchDropDownLists()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim params As SqlParameter() = Nothing
        Dim strSQL As String = " "

        strSQL = "select QueueName from tbl_Web_Queues where Id=@Qid"
        params = {New SqlParameter("@Qid", Session(SessionVariable_Prefix & "queueid"))}
        Dim StrQueueName As String = CommonUtilsv2.RunScalarQuery(dbKey, strSQL, CommandType.Text, params)
        If (StrQueueName.ToLower().Contains("delete")) Then
            IsDeleteGroup = True
        End If
        params = Nothing


        strSQL = " select ID as DocTypeID,DocTypeName AS DocType from tbl_Web_DocTypes  WHERE RecordStatus=1 AND GroupID=(select GroupID from tbl_Web_Queues where ID=@Qid) order by DocTypeName "
        params = {New SqlParameter("@Qid", Session(SessionVariable_Prefix & "queueid"))}

        CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddSearch2, "DocTypeID", "DocType", params)
        ddSearch2.Items.Insert(0, New ListItem("Select A Value ", 0))
        ddSearch2.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch2Index")


        Dim StrUrl As String = HttpContext.Current.Request.RawUrl.ToString()
        Dim StrQueue As String = StrUrl.Substring(StrUrl.IndexOf("?") + 1).ToString()
        Dim QueueList() As String = StrQueue.Split("&")
        Dim QueueId As New Integer
        For Each str As String In QueueList
            Dim StrList() As String = str.Split("=")
            If StrList(0) = "qid" Then
                QueueId = StrList(1)
            End If

        Next
        params = Nothing

        strSQL = " select ID AS QueueID, QueueName FROM tbl_Web_Queues  where GroupID=(select GroupID from tbl_Web_Queues where  ID=@Qid)  and RecordStatus=1 order by QueueName "

        params = {New SqlParameter("@Qid", Session(SessionVariable_Prefix & "queueid"))}

        CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddQueue, "QueueID", "QueueName", params)
        ddQueue.Items.Insert(0, New ListItem("Select A Value ", 0))

        'strSQL = " select IndexName,ColumnName from tbl_Web_SearchIndexNames where GroupID=(select GroupID  from tbl_Web_Queues where ID=" & QueueId & ") order by SortOrder "
        'CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddSearch1, "IndexName", "ColumnName", Nothing)
        'ddSearch1.Items.Insert(0, New ListItem("Select A Value ", 0))
        'ddSearch1.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch1Index")


        ' Move document between Queues  
        params = Nothing
        ' strSQL = " select ID AS QueueID, QueueName FROM tbl_Web_Queues WHERE  RecordStatus=1 order by QueueName "
        strSQL = " select ID AS QueueID, QueueName FROM v_w_DestinationQueueList  where GroupID=(select GroupID from tbl_Web_Queues where  ID=@Qid) and RecordStatus=1 order by QueueName "
        params = {New SqlParameter("@Qid", Session(SessionVariable_Prefix & "queueid"))}
        CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddDestinationQueue, "QueueID", "QueueName", params)
        ddDestinationQueue.Items.Insert(0, New ListItem("Select A Value ", 0))

        params = Nothing

        '  strSQL = " SELECT ID AS DocTypeID, DocTypeName AS DocType FROM tbl_Web_DocTypes WHERE RecordStatus=1 order by DocTypeName "
        strSQL = " select ID as DocTypeID,DocTypeName AS DocType from tbl_Web_DocTypes  WHERE RecordStatus=1 AND GroupID=(select GroupID from tbl_Web_Queues where ID=@Qid) order by DocTypeName "

        'Dim params2 As SqlParameter() = {New SqlParameter("@Qid", Session(SessionVariable_Prefix & "queueid"))}
        'CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddDestinationDocType, "DocTypeID", "DocType", params2)
        'ddDestinationDocType.Items.Insert(0, New ListItem("Select A Value ", 0))
        'params2 = Nothing

        'GetOwner from tbl_web_lookup
        strSQL = " SELECT LookupID As OwnerID, LookupDesc As OwnerName FROM tbl_Web_Lookup  WHERE Lookuptype = @LookupType Order by SortOrder, LookupDesc "
        params = {New SqlParameter("@LookupType", "DocOwner")}
        CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddDocUser, "OwnerID", "OwnerName", params)
        ddDocUser.Items.Insert(0, New ListItem("Select A Value ", 0))

        strSQL = " select ID as DocTypeID,DocTypeName AS DocType from tbl_Web_DocTypes  WHERE RecordStatus=1 AND GroupID=(select GroupID from tbl_Web_Queues where ID=@Qid) order by DocTypeName "
        Dim params3 As SqlParameter() = {New SqlParameter("@Qid", Session(SessionVariable_Prefix & "queueid"))}

        CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddDocType, "DocTypeID", "DocType", params3)
        ddDocType.Items.Insert(0, New ListItem("Select A Value ", 0))

        params3 = Nothing

        If (Isprovdier > 0) Then
            'If (Isprovdier = 2) Then ' is complete
            Dim params5 As SqlParameter() = Nothing
            strSQL = " SELECT LookupID As OwnerID, LookupDesc As OwnerName FROM tbl_Web_Lookup  WHERE Lookuptype = @LookupType Order by SortOrder, LookupDesc  "
            params = {New SqlParameter("@LookupType", "DocOwner")}
            CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddOwner, "OwnerID", "OwnerName", params)
            ddOwner.Items.Insert(0, New ListItem("Select A Value ", 0))
            'End If
        Else
            strSQL = " SELECT LookupID, LookupDesc as LookupDescription From tbl_Web_Lookup WHERE Lookuptype = @LookupType Order by SortOrder ASC "
            params = {New SqlParameter("@LookupType", "DocStatus")}
            CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddSearch3, "LookupID", "LookupDescription", params)
            ddSearch3.Items.Insert(0, New ListItem("Select A Value ", 0))
            ddSearch3.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch3Index")

            strSQL = " SELECT LookupID,  LookupDesc as LookupDescription From tbl_Web_Lookup WHERE Lookuptype = @LookupType Order by SortOrder, LookupDesc ASC "
            params = {New SqlParameter("@LookupType", "DocPriority")}
            CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddSearch4, "LookupID", "LookupDescription", params)
            ddSearch4.Items.Insert(0, New ListItem("Select A Value ", 0))
            ddSearch4.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch4Index")
        End If


    End Sub

    Private Sub GetDynamicIndexs()
        Dim ds As DataSet
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim params As SqlParameter() = Nothing

        AddHandler CType(gridQueue, Grid).Rebind, AddressOf RebindGrid
        Dim strSQL As String = " select IndexName,ColumnName,DateType from tbl_Web_SearchIndexNames where GroupID=(select GroupID  from tbl_Web_Queues where ID=@qid ) order by SortOrder "

        params = {New SqlParameter("@qid", Session(SessionVariable_Prefix & "queueid"))}

        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)
            If Not ds Is Nothing Then
                Dim SearchStr As String = String.Empty
                Dim Searchi As Integer = 0
                Dim rw1 As TableRow = New TableRow()
                rw1.Attributes("class") = "body_search"
                Dim rw2 As TableRow = New TableRow()
                rw2.Attributes("class") = "body_search"
                Dim rw3 As TableRow = New TableRow()
                rw3.Attributes("class") = "body_search"
                Dim rw4 As TableRow = New TableRow()
                rw4.Attributes("class") = "body_search"
                For Each row As System.Data.DataRow In ds.Tables(0).Rows

                    Dim tb1 As TableCell = New TableCell()
                    tb1.Attributes("align") = "right"
                    tb1.Attributes("width") = "10%"
                    Dim lable As Label = New Label()
                    lable.Text = row("ColumnName").ToString().Trim() & ":"
                    tb1.Controls.Add(lable)
                    'rw.Cells.Add(tb1)

                    Dim tb2 As TableCell = New TableCell()
                    tb2.Attributes("align") = "left"
                    tb2.Attributes("class") = "tdPadding"
                    If (Searchi Mod 3) = 1 Then
                        tb2.Attributes("width") = "24%"
                    Else
                        tb2.Attributes("width") = "23%"
                    End If



                    If (row("DateType").ToString().Trim().ToUpper() = "SELECT") Then

                        Dim txtSelect As DropDownList = New DropDownList()
                        txtSelect.ID = "tbSelect" & row("IndexName").ToString().Trim()
                        'txtSelect.Attributes("colunname") = row("ColumnName").ToString().Trim()
                        txtSelect.Attributes("name") = row("IndexName").ToString().Trim()
                        txtSelect.Attributes("style") = "height:  25px;width:78%;"

                        ''state
                        If (row("ColumnName").ToString().ToUpper().Contains("STATE")) Then
                            txtSelect.Attributes("class") = "selectstate"
                            strSQL = "select * from Tbl_Web_States  "
                            CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, txtSelect, "state_name", "state_name")
                            txtSelect.Items.Insert(0, New ListItem("Select A Value ", ""))
                        End If
                        'County
                        If (row("ColumnName").ToString().ToUpper().Contains("COUNTY")) Then
                            txtSelect.Attributes("class") = "selectcounty"
                            txtSelect.Items.Insert(0, New ListItem("Select A Value ", ""))
                        End If
                        tb2.Controls.Add(txtSelect)
                    Else
                        If row("DateType").ToString().Trim().ToUpper() = "DATE" Then
                            If (row("ColumnName").ToString().ToUpper().Contains("LOAD")) Or (row("ColumnName").ToString().ToUpper().Contains("CREATE")) Then
                                Dim txt As TextBox = New TextBox()
                                txt.ID = "tb" & "DateStarte" & row("IndexName").ToString().Trim()
                                txt.Attributes("maxlength") = "30"
                                txt.Attributes("name") = "DateStarte" & row("IndexName").ToString().Trim()
                                txt.Attributes("style") = "height:  25px;width:35%;"
                                'txt.ReadOnly = True
                                txt.Attributes("class") = "Wdate"
                                txt.Attributes("onclick") = "WdatePicker({lang:'en',dateFmt:'MM/dd/yyyy'})"
                                tb2.Controls.Add(txt)

                                Dim Lable1 As Label = New Label()
                                Lable1.Text = "     To     "
                                tb2.Controls.Add(Lable1)

                                Dim txt2 As TextBox = New TextBox()
                                txt2.ID = "tb" & "DateEnd" & row("IndexName").ToString().Trim()
                                txt2.Attributes("maxlength") = "30"
                                txt2.Attributes("name") = "DateEnd" & row("IndexName").ToString().Trim()
                                txt2.Attributes("style") = "height:  25px;width:35%;"
                                'txt2.ReadOnly = True
                                txt2.Attributes("class") = "Wdate"
                                txt2.Attributes("onclick") = "WdatePicker({lang:'en',dateFmt:'MM/dd/yyyy'})"
                                tb2.Controls.Add(txt2)

                            Else
                                Dim txt As TextBox = New TextBox()
                                txt.ID = "tb" & row("IndexName").ToString().Trim()
                                txt.Attributes("maxlength") = "30"
                                txt.Attributes("name") = row("IndexName").ToString().Trim()
                                txt.Attributes("style") = "height:  25px;width:78%;"
                                tb2.Controls.Add(txt)
                                txt.ReadOnly = True
                                txt.Attributes("class") = "Wdate"
                                txt.Attributes("onclick") = "WdatePicker({lang:'en',dateFmt:'MM/dd/yyyy'})"
                            End If
                        Else
                            Dim txt As TextBox = New TextBox()
                            txt.ID = "tb" & row("IndexName").ToString().Trim()
                            txt.Attributes("maxlength") = "30"
                            txt.Attributes("name") = row("IndexName").ToString().Trim()
                            txt.Attributes("style") = "height:  25px;width:78%;"
                            tb2.Controls.Add(txt)
                        End If
                    End If


                    Select Case Searchi
                        Case 0
                            rw1.Cells.Add(tb1)
                            rw1.Cells.Add(tb2)
                        Case 1
                            rw1.Cells.Add(tb1)
                            rw1.Cells.Add(tb2)
                        Case 2
                            rw1.Cells.Add(tb1)
                            rw1.Cells.Add(tb2)
                        Case 3
                            rw2.Cells.Add(tb1)
                            rw2.Cells.Add(tb2)
                        Case 4
                            rw2.Cells.Add(tb1)
                            rw2.Cells.Add(tb2)
                        Case 5
                            rw2.Cells.Add(tb1)
                            rw2.Cells.Add(tb2)
                        Case 6
                            rw3.Cells.Add(tb1)
                            rw3.Cells.Add(tb2)
                        Case 7
                            rw3.Cells.Add(tb1)
                            rw3.Cells.Add(tb2)
                        Case 8
                            rw3.Cells.Add(tb1)
                            rw3.Cells.Add(tb2)
                        Case 9
                            rw4.Cells.Add(tb1)
                            rw4.Cells.Add(tb2)
                    End Select

                    Searchi = Searchi + 1
                Next

                dynamicIndexs.Rows.Add(rw1)
                dynamicIndexs.Rows.Add(rw2)
                dynamicIndexs.Rows.Add(rw3)
                dynamicIndexs.Rows.Add(rw4)
                txtTextCount.Value = Searchi
            End If
        Catch ex As Exception
            errString = ex.Message
            errLocation = "GetDynamicIndexs()"
            CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try

    End Sub

    'Add column names in the background  create:2021-03-22
    Private Sub CreateGridColumns()
        Dim ds As DataSet
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim params As SqlParameter() = Nothing

        AddHandler CType(gridQueue, Grid).Rebind, AddressOf RebindGrid
        Dim strSQL As String = " select IndexName,ColumnName from tbl_Web_SearchIndexNames where GroupID=(select GroupID  from tbl_Web_Queues where ID=@qid ) order by SortOrder "

        params = {New SqlParameter("@qid", Session(SessionVariable_Prefix & "queueid"))}

        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)
            If Not ds Is Nothing Then
                Dim i As Integer = 111
                Dim SearchStr As String = String.Empty
                For Each row As System.Data.DataRow In ds.Tables(0).Rows
                    ''add about suite table
                    Dim oCol1 As Column = New Column()
                    oCol1.ID = "Column" & i
                    oCol1.HeaderText = row("ColumnName").ToString().Trim()
                    oCol1.DataField = row("IndexName").ToString().Trim()
                    oCol1.Align = "center"
                    oCol1.Wrap = True
                    oCol1.HeaderAlign = "center"
                    oCol1.ReadOnly = True
                    oCol1.Width = "120"
                    gridQueue.Columns.Add(oCol1)
                    i = i + 1
                Next
            End If
        Catch ex As Exception
            errString = ex.Message
            errLocation = "CreateGridColumns()"
            CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try

    End Sub

    Private Sub CreateGrid()
        If String.IsNullOrEmpty(Session("User")) Then
            Exit Sub
        End If

        CreateGridColumns() 'Add column names in the background  create:2021-03-22
        StrTextCount = txtTextCount.Value
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim myComm As New SqlCommand()
        Dim searchStr As String = Session(SessionVariable_Prefix & "searchParam1Value")
        Dim myReader As SqlDataReader = Nothing
        Dim htSearchParams As Hashtable = New Hashtable()
        Dim strWhereClause As StringBuilder = New StringBuilder()
        Try
            myComm.Connection = myConn
            myComm.CommandText = strCommandBase

            Dim strFilter As String = CommonUtilsv2.GetDCNRecordFilterByUser(Session("User"))
            If String.IsNullOrEmpty(strFilter) Then

            Else
                myComm.CommandText = myComm.CommandText & " AND ( " & strFilter & " ) "
                strWhereClause.Append(" AND ( " & strFilter & " ) ")
            End If


            If Session(SessionVariable_Prefix & "IsSearch") <> "TRUE" Then
                If mydocument.Checked = True Then
                    myComm.CommandText = myComm.CommandText & " AND " & " Owner=@Owner "
                    myComm.Parameters.AddWithValue("@Owner", Session("User"))
                    htSearchParams.Add("@Owner", Session("User"))
                    strWhereClause.Append(" AND " & " Owner= @Owner ")
                End If
                strDisplaySearchCriteria = "Displaying: All " + strDisplayItemName
                myComm.CommandText = myComm.CommandText & " AND " & " QueueID= @qid "
                myComm.Parameters.AddWithValue("@qid", Session(SessionVariable_Prefix & "queueid"))
                htSearchParams.Add("@qid", Session(SessionVariable_Prefix & "queueid"))
                strWhereClause.Append(" AND " & " QueueID= @qid ")

                myComm.CommandText = myComm.CommandText & strCommandOrderBy
                Session(SessionVariable_Prefix & "IsSearch") = "FALSE"
                Session(SessionVariable_Prefix & "SearchBy") = "ViewAll"
                Session(SessionVariable_Prefix & "SearchLabel") = strDisplaySearchCriteria
            Else
                Select Case Session(SessionVariable_Prefix & "SearchBy")

                    Case "Search"

                        If Session(SessionVariable_Prefix & "searchParam2Value") <> CR_searchParam2Default Then
                            If Session(SessionVariable_Prefix & "searchParam2Value") <> "" Then
                                myComm.CommandText = myComm.CommandText & " AND " & CR_searchParam2ColumnName & " = @SearchedParm2 "
                                myComm.Parameters.AddWithValue("@SearchedParm2", Session(SessionVariable_Prefix & "searchParam2Value"))
                                htSearchParams.Add("@SearchedParm2", ddSearch2.SelectedValue)
                                strWhereClause.Append(" AND " & CR_searchParam2ColumnName & " = @SearchedParm2 ")
                            Else
                            End If
                        End If

                        If StrHelp.GetInt(Session(SessionVariable_Prefix & "searchParam3Value")) <> StrHelp.GetInt(CR_searchParam3Default) Then
                            If Session(SessionVariable_Prefix & "searchParam3Value") <> "" Then
                                myComm.CommandText = myComm.CommandText & " AND " & CR_searchParam3ColumnName & " = @SearchedParm3 "
                                myComm.Parameters.AddWithValue("@SearchedParm3", Session(SessionVariable_Prefix & "searchParam3Value"))
                                htSearchParams.Add("@SearchedParm3", ddSearch3.SelectedValue)
                                strWhereClause.Append(" AND " & CR_searchParam3ColumnName & " = @SearchedParm3 ")
                            Else
                            End If
                        End If
                        If StrHelp.GetInt(Session(SessionVariable_Prefix & "searchParam4Value")) <> StrHelp.GetInt(CR_searchParam4Default) Then
                            If Session(SessionVariable_Prefix & "searchParam4Value") <> "" Then
                                myComm.CommandText = myComm.CommandText & " AND " & CR_searchParam4ColumnName & " = @SearchedParm4 "
                                myComm.Parameters.AddWithValue("@SearchedParm4", Session(SessionVariable_Prefix & "searchParam4Value"))
                                htSearchParams.Add("@SearchedParm4", ddSearch4.SelectedValue)
                                strWhereClause.Append(" AND " & CR_searchParam4ColumnName & " = @SearchedParm4 ")
                            Else
                            End If
                        End If
                        If StrHelp.GetInt(Session(SessionVariable_Prefix & "searchParam5Value")) <> StrHelp.GetInt(CR_searchParam5Default) Then
                            myComm.CommandText = myComm.CommandText & " AND " & CR_searchParam5ColumnName & " = @SearchedParm5 "
                            myComm.Parameters.AddWithValue("@SearchedParm5", Session(SessionVariable_Prefix & "searchParam5Value"))
                            htSearchParams.Add("@SearchedParm5", ddOwner.SelectedValue)
                            strWhereClause.Append(" AND " & CR_searchParam5ColumnName & " = @SearchedParm5 ")
                        End If

                        If Session(SessionVariable_Prefix & "SQueueStartDateParam1Value") <> "" Then
                            myComm.CommandText = myComm.CommandText & " AND  DATEDIFF(day,@SQueueStartDateParam,QueueStartDate)>=0 "
                            myComm.Parameters.AddWithValue("@SQueueStartDateParam", Session(SessionVariable_Prefix & "SQueueStartDateParam1Value"))
                            htSearchParams.Add("@SQueueStartDateParam", Session(SessionVariable_Prefix & "SQueueStartDateParam1Value"))
                            strWhereClause.Append(" AND  DATEDIFF(day,@SQueueStartDateParam,QueueStartDate)>0 ")
                        End If
                        If Session(SessionVariable_Prefix & "EQueueStartDateParam1Value") <> "" Then
                            myComm.CommandText = myComm.CommandText & " AND  DATEDIFF(day,QueueStartDate,@EQueueStartDateParam)>=0 "
                            myComm.Parameters.AddWithValue("@EQueueStartDateParam", Session(SessionVariable_Prefix & "EQueueStartDateParam1Value"))
                            htSearchParams.Add("@EQueueStartDateParam", Session(SessionVariable_Prefix & "EQueueStartDateParam1Value"))
                            strWhereClause.Append(" AND DATEDIFF(day,QueueStartDate,@EQueueStartDateParam)>0 ")
                        End If

                        Dim txtCount As String = txtTextCount.Value
                        For i As Integer = 1 To txtCount
                            If Request.Form.Count > 0 Then
                                Try
                                    For Each key As String In Request.Form.Keys
                                        If key Is Nothing Then
                                            Continue For
                                        End If
                                        If (key.Trim().Contains("Index" & i.ToString())) Then

                                            If (key.Trim() = ("ctl00$Body$tbDateStarteIndex" & i.ToString())) Then
                                                Dim strtxt As String = Request.Form("ctl00$Body$tbDateStarteIndex" & i.ToString()).ToString()
                                                If strtxt IsNot Nothing Then
                                                    If Not String.IsNullOrEmpty(strtxt.TrimEnd(",")) Then
                                                        Dim strSearchStr As String = strtxt.TrimEnd(",")
                                                        myComm.CommandText = myComm.CommandText & " AND DATEDIFF(day,@StarteIndex" & i.ToString() & ",UploadedDate)>=0"
                                                        myComm.Parameters.AddWithValue("@StarteIndex" & i.ToString(), strSearchStr)
                                                        strWhereClause.Append(" AND DATEDIFF(day,@StarteIndex" & i.ToString() & ",UploadedDate)>=0")
                                                        htSearchParams.Add("@StarteIndex" & i.ToString(), strSearchStr)
                                                    End If
                                                End If
                                            End If

                                            If (key.Trim() = ("ctl00$Body$tbDateEndIndex" & i.ToString())) Then
                                                Dim strtxt As String = Request.Form("ctl00$Body$tbDateEndIndex" & i.ToString()).ToString()
                                                If strtxt IsNot Nothing Then
                                                    If Not String.IsNullOrEmpty(strtxt.TrimEnd(",")) Then
                                                        Dim strSearchStr As String = strtxt.TrimEnd(",")
                                                        myComm.CommandText = myComm.CommandText & " AND DATEDIFF(day,UploadedDate,@EndIndex" & i.ToString() & ")>=0"
                                                        myComm.Parameters.AddWithValue("@EndIndex" & i.ToString(), strSearchStr)
                                                        strWhereClause.Append(" AND DATEDIFF(day,UploadedDate,@EndIndex" & i.ToString() & ")>=0")
                                                        htSearchParams.Add("@EndIndex" & i.ToString(), strSearchStr)
                                                    End If
                                                End If
                                            End If

                                            If (key.Trim() = ("ctl00$Body$tbIndex" & i.ToString())) Then
                                                Dim strtxt As String = Request.Form("ctl00$Body$tbIndex" & i.ToString()).ToString()
                                                If strtxt IsNot Nothing Then
                                                    If Not String.IsNullOrEmpty(strtxt.TrimEnd(",")) Then
                                                        Dim strSearchStr As String = "%" + strtxt.TrimEnd(",") + "%"
                                                        myComm.CommandText = myComm.CommandText & " AND Index" & i.ToString() & " LIKE @Index" & i.ToString()
                                                        myComm.Parameters.AddWithValue("@Index" & i.ToString(), strSearchStr)
                                                        strWhereClause.Append(" AND Index" & i.ToString() & " LIKE  @Index" & i.ToString())
                                                        htSearchParams.Add("@Index" & i.ToString(), strSearchStr)
                                                    End If
                                                End If
                                            End If

                                        End If
                                        If (key.Trim().Contains("tbSelectIndex")) Then
                                            If (key.Trim() = ("ctl00$Body$tbSelectIndex" & i.ToString())) Then
                                                Dim strtxt As String = Request.Form("ctl00$Body$tbSelectIndex" & i.ToString()).ToString()
                                                If strtxt IsNot Nothing Then
                                                    If Not String.IsNullOrEmpty(strtxt.TrimEnd(",")) Then
                                                        Dim strSearchStr As String = strtxt.TrimEnd(",")
                                                        myComm.CommandText = myComm.CommandText & " AND Index" & i.ToString() & "=@Index" & i.ToString()
                                                        myComm.Parameters.AddWithValue("@Index" & i.ToString(), strSearchStr)
                                                        strWhereClause.Append(" AND Index" & i.ToString() & "=@Index" & i.ToString())
                                                        htSearchParams.Add("@Index" & i.ToString(), strSearchStr)
                                                    End If
                                                End If
                                            End If
                                        End If

                                    Next
                                Catch exIdx As Exception
                                    errString = exIdx.Message
                                    errLocation = "Index search in CreateGrid()"
                                    CommonUtilsv2.CreateErrorLog(errLocation, exIdx, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
                                End Try
                            End If
                        Next

                    Case Else
                        'reset all
                        strDisplaySearchCriteria = "Displaying: All documents "

                End Select

                If mydocument.Checked = True Then
                    myComm.CommandText = myComm.CommandText & " AND " & " Owner=@Owner "
                    myComm.Parameters.AddWithValue("@Owner", Session("User"))
                    htSearchParams.Add("@Owner", Session("User"))
                    strWhereClause.Append(" AND " & " Owner= @Owner ")
                End If


                myComm.CommandText = myComm.CommandText & " AND " & " QueueID= @qid "
                myComm.Parameters.AddWithValue("@qid", Session(SessionVariable_Prefix & "queueid"))
                htSearchParams.Add("@qid", Session(SessionVariable_Prefix & "queueid"))
                strWhereClause.Append(" AND " & " QueueID= @qid ")

                myComm.CommandText = myComm.CommandText & strCommandOrderBy
            End If


            myConn.Open()
            myReader = myComm.ExecuteReader()
            gridQueue.DataSource = myReader
            gridQueue.DataBind()

            If CustomRoles.IsInRole("R_Queue_Export") Then
                Dim lExportLimit As Long
                Long.TryParse(Application("AllowedMaxExportCount"), lExportLimit)
                If gridQueue.TotalRowCount() < lExportLimit Then
                    ExportLabel.Visible = False
                    dvExportExcel.Visible = True
                Else
                    ExportLabel.Visible = True
                    dvExportExcel.Visible = False
                End If
            Else
                dvExportExcel.Visible = False
                ExportLabel.Visible = False
            End If


            'export
            Session(SessionVariable_Prefix & "SearchParams") = htSearchParams
            Session(SessionVariable_Prefix & "WhereClause") = strWhereClause.ToString
            Session(SessionVariable_Prefix & "SelectColumns") = strExportSQL
        Catch ex As Exception
            errString = ex.Message
            errLocation = "CreateGrid()"
            CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
            myComm.Dispose()
            myConn.Close()
            myConn.Dispose()
            myConn = Nothing
        End Try
    End Sub

    Sub RowDataBound(ByVal sender As Object, ByVal e As GridRowEventArgs)   'Set Row highlights as data is bound based on cell values

    End Sub

    Protected Sub SetSearchSettings()
        Dim bHasFirstSearch As Boolean = False

        lbSearchApplied.Text = ""
        lbSearchApplied.ForeColor = Drawing.Color.Green
        strDisplaySearchCriteria = "" '"Displaying: All " + strDisplayItemName
        'CR_searchParam1Value = tbSearch1.Text.Trim()

        CR_searchParam2Value = ddSearch2.SelectedValue.ToString
        CR_searchParam3Value = ddSearch3.SelectedValue.ToString
        CR_searchParam4Value = ddSearch4.SelectedValue.ToString
        CR_searchParam5Value = ddOwner.SelectedValue.ToString
        Session(SessionVariable_Prefix & "SQueueStartDateParam1Value") = SQueueStartDate.Text
        Session(SessionVariable_Prefix & "EQueueStartDateParam1Value") = EQueueStartDate.Text


        If CR_searchParam2Value <> CR_searchParam2Default Or CR_searchParam3Value <> CR_searchParam3Default Or CR_searchParam4Value <> CR_searchParam4Default Then
            If CR_searchParam2Value <> CR_searchParam2Default Then
                If bHasFirstSearch Then
                    strDisplaySearchCriteria += " AND " & CR_searchParam2PrettyName & "= '" + ddSearch2.SelectedItem.Text + "'"
                Else
                    strDisplaySearchCriteria += CR_searchParam2PrettyName & "= '" + ddSearch2.SelectedItem.Text + "'"
                    bHasFirstSearch = True
                End If
            End If

            If StrHelp.GetInt(CR_searchParam3Value) <> StrHelp.GetInt(CR_searchParam3Default) Then
                If bHasFirstSearch Then
                    strDisplaySearchCriteria += " AND " & CR_searchParam3PrettyName & "= '" + ddSearch3.SelectedItem.Text + "'"
                Else
                    strDisplaySearchCriteria += CR_searchParam3PrettyName & "= '" + ddSearch3.SelectedItem.Text + "'"
                    bHasFirstSearch = True
                End If
            End If

            If StrHelp.GetInt(CR_searchParam4Value) <> StrHelp.GetInt(CR_searchParam4Default) Then
                If bHasFirstSearch Then
                    strDisplaySearchCriteria += " AND " & CR_searchParam4PrettyName & "= '" + ddSearch4.SelectedItem.Text + "'"
                Else
                    strDisplaySearchCriteria += CR_searchParam4PrettyName & "= '" + ddSearch4.SelectedItem.Text + "'"
                    bHasFirstSearch = True
                End If
            End If

            If StrHelp.GetInt(CR_searchParam4Value) <> StrHelp.GetInt(CR_searchParam4Default) Then
                If bHasFirstSearch Then
                    strDisplaySearchCriteria += " AND " & CR_searchParam4PrettyName & "= '" + ddSearch4.SelectedItem.Text + "'"
                Else
                    strDisplaySearchCriteria += CR_searchParam4PrettyName & "= '" + ddSearch4.SelectedItem.Text + "'"
                    bHasFirstSearch = True
                End If
            End If

            If StrHelp.GetInt(CR_searchParam5Value) <> StrHelp.GetInt(CR_searchParam5Default) Then
                If bHasFirstSearch Then
                    strDisplaySearchCriteria += " AND " & CR_searchParam5PrettyName & "= '" + ddOwner.SelectedItem.Text + "'"
                Else
                    strDisplaySearchCriteria += CR_searchParam5PrettyName & "= '" + ddOwner.SelectedItem.Text + "'"
                    bHasFirstSearch = True
                End If
            End If
            If Session(SessionVariable_Prefix & "SQueueStartDateParam1Value") <> "" Then
                If bHasFirstSearch Then
                    strDisplaySearchCriteria += " AND Complete Date  >= '" + Session(SessionVariable_Prefix & "SQueueStartDateParam1Value") + "'"
                Else
                    strDisplaySearchCriteria += " Complete Date >= '" + Session(SessionVariable_Prefix & "SQueueStartDateParam1Value") + "'"
                    bHasFirstSearch = True
                End If
            End If
            If Session(SessionVariable_Prefix & "EQueueStartDateParam1Value") <> "" Then
                If bHasFirstSearch Then
                    strDisplaySearchCriteria += " AND Complete Date<= '" + Session(SessionVariable_Prefix & "EQueueStartDateParam1Value") + "'"
                Else
                    strDisplaySearchCriteria += " Complete Date <= '" + Session(SessionVariable_Prefix & "EQueueStartDateParam1Value") + "'"
                    bHasFirstSearch = True
                End If
            End If

        End If

        Dim txtCount As String = txtTextCount.Value
        Dim ds As DataSet
        Dim params As SqlParameter() = Nothing
        Dim anID As String = Request.QueryString.Get("qid")
        Dim StrSql = "Select IndexName,ColumnName,DateType from tbl_Web_SearchIndexNames  where GroupID=(Select GroupID from tbl_Web_Queues where ID=@ID)"
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        params = {New SqlParameter("@ID", anID)}
        ds = CommonUtilsv2.GetDataSet(dbKey, StrSql, CommandType.Text, params)
        For i As Integer = 1 To txtCount
            Dim dr() As DataRow = ds.Tables(0).Select("IndexName='Index" + i.ToString() + "'")
            Dim IndexsName As String = dr(0)("ColumnName").ToString()
            Dim IndexType As String = dr(0)("DateType").ToString()
            If (IndexsName.ToUpper().Contains("LOAD") Or IndexsName.ToUpper().Contains("CREATE")) Then
                Dim strtxtstarte As String = Request.Form("ctl00$Body$tbDateStarteIndex" & i.ToString()).ToString()
                If strtxtstarte IsNot Nothing Then
                    If Not String.IsNullOrEmpty(strtxtstarte) Then
                        If bHasFirstSearch Then
                            strDisplaySearchCriteria += " AND " & IndexsName & " >= '" + strtxtstarte + "'"
                        Else
                            strDisplaySearchCriteria += IndexsName & " >= '" + strtxtstarte + "'"
                            bHasFirstSearch = True
                        End If
                    End If
                End If

                Dim strtxtend As String = Request.Form("ctl00$Body$tbDateEndIndex" & i.ToString()).ToString()
                If strtxtend IsNot Nothing Then
                    If Not String.IsNullOrEmpty(strtxtend) Then
                        If bHasFirstSearch Then
                            strDisplaySearchCriteria += " AND " & IndexsName & " <= '" + strtxtend + "'"
                        Else
                            strDisplaySearchCriteria += IndexsName & "  <= '" + strtxtend + "'"
                            bHasFirstSearch = True
                        End If
                    End If
                End If

            ElseIf IndexType.ToUpper().Contains("SELECT") Then
                Dim strtxt As String = Request.Form("ctl00$Body$tbSelectIndex" & i.ToString()).ToString()
                If strtxt IsNot Nothing Then
                    If Not String.IsNullOrEmpty(strtxt) Then
                        If bHasFirstSearch Then
                            strDisplaySearchCriteria += " AND " & IndexsName & "  = '" + strtxt + "'"
                        Else
                            strDisplaySearchCriteria += IndexsName & "  = '" + strtxt + "'"
                            bHasFirstSearch = True
                        End If
                    End If
                End If
            Else
                Dim strtxt As String = Request.Form("ctl00$Body$tbIndex" & i.ToString()).ToString()
                If strtxt IsNot Nothing Then
                    If Not String.IsNullOrEmpty(strtxt) Then
                        If bHasFirstSearch Then
                            strDisplaySearchCriteria += " AND " & IndexsName & "  contains '" + strtxt + "'"
                        Else
                            strDisplaySearchCriteria += IndexsName & "  contains '" + strtxt + "'"
                            bHasFirstSearch = True
                        End If
                    End If
                End If

            End If
        Next


        Select Case Session(SessionVariable_Prefix & "SearchBy")

            Case "Search"
                If bHasFirstSearch Then
                    Session(SessionVariable_Prefix & "IsSearch") = "TRUE"
                    strDisplaySearchCriteria = "Displaying: Documents with " + strDisplaySearchCriteria
                Else
                    strDisplaySearchCriteria = "Displaying: All " + strDisplayItemName
                    Session(SessionVariable_Prefix & "IsSearch") = "FALSE"
                End If

            Case "MyClaims"
                Session(SessionVariable_Prefix & "IsSearch") = "TRUE"
                If bHasFirstSearch Then
                    strDisplaySearchCriteria = "Displaying: My Documents with " + strDisplaySearchCriteria
                Else
                    strDisplaySearchCriteria = "Displaying: My Documents"
                End If

            Case Else
                strDisplaySearchCriteria = "Displaying: All " + strDisplayItemName
                Session(SessionVariable_Prefix & "IsSearch") = "FALSE"
        End Select

        Session(SessionVariable_Prefix & "GridPageIndex") = 0
        gridQueue.CurrentPageIndex = 0

        Session(SessionVariable_Prefix & "ddSearch2Index") = ddSearch2.SelectedIndex
        Session(SessionVariable_Prefix & "ddSearch3Index") = ddSearch3.SelectedIndex
        Session(SessionVariable_Prefix & "ddSearch4Index") = ddSearch4.SelectedIndex
        Session(SessionVariable_Prefix & "ddSearch5Index") = ddOwner.SelectedIndex


        Session(SessionVariable_Prefix & "SearchLabel") = strDisplaySearchCriteria
        Session(SessionVariable_Prefix & "searchParam1ColumnName") = CR_searchParam1ColumnName
        Session(SessionVariable_Prefix & "searchParam1Value") = CR_searchParam1Value
        Session(SessionVariable_Prefix & "searchParam2Value") = CR_searchParam2Value
        Session(SessionVariable_Prefix & "searchParam3Value") = CR_searchParam3Value
        Session(SessionVariable_Prefix & "searchParam4Value") = CR_searchParam4Value
        Session(SessionVariable_Prefix & "searchParam5Value") = CR_searchParam5Value


    End Sub

    Private Sub RebindGrid(ByVal sender As Object, ByVal e As EventArgs)
        lbSearchApplied.Text = Session(SessionVariable_Prefix & "SearchLabel")
        CreateGrid()
        SetFromSession()
    End Sub

    Protected Sub btn_ViewAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_ViewAll.Click
        mydocument.Checked = False
        SetDefaultSession()
        CreateGrid()
        SetFromSession()
    End Sub

    ''' <summary>
    ''' btn_Search 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btn_Search_Click_All(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_Search.Click
        Session(SessionVariable_Prefix & "SearchBy") = "Search"
        SetSearchSettings()
        CreateGrid()
        SetFromSession()
    End Sub

    Protected Sub SetSessionVariables()
        If Session(SessionVariable_Prefix & "IsSearch") = "TRUE" Then
            Select Case Session(SessionVariable_Prefix & "SearchBy")

                Case "Search"
                    Session(SessionVariable_Prefix & "ddSearch2Index") = ddSearch2.SelectedIndex
                    Session(SessionVariable_Prefix & "ddSearch3Index") = ddSearch3.SelectedIndex
                    Session(SessionVariable_Prefix & "ddSearch4Index") = ddSearch4.SelectedIndex
                    Session(SessionVariable_Prefix & "ddSearch5Index") = ddOwner.SelectedIndex

                    'Session(SessionVariable_Prefix & "ddSearch1Index") = ddSearch1.SelectedIndex
                    'Session(SessionVariable_Prefix & "searchParam1Value") = tbSearch1.Text

                Case "MyClaims"
                    Session(SessionVariable_Prefix & "GridPageIndex") = 0
                    gridQueue.CurrentPageIndex = 0
                    Session(SessionVariable_Prefix & "SearchLabel") = "Displaying: " + strDisplayItemName

                    Session(SessionVariable_Prefix & "ddSearch2Index") = 0 ' ddSearch2.SelectedIndex
                    Session(SessionVariable_Prefix & "ddSearch3Index") = ddSearch3.SelectedIndex
                    Session(SessionVariable_Prefix & "ddSearch4Index") = ddSearch4.SelectedIndex
                    Session(SessionVariable_Prefix & "ddSearch5Index") = ddOwner.SelectedIndex

                    'Session(SessionVariable_Prefix & "ddSearch1Index") = ddSearch1.SelectedIndex
                    'Session(SessionVariable_Prefix & "searchParam1Value") = tbSearch1.Text

                Case Else
                    'reset all
                    SetDefaultSession()
            End Select
        Else
            'reset all
            SetDefaultSession()
        End If
    End Sub

    Protected Sub SetFromSession()
        Try
            If Session(SessionVariable_Prefix & "IsSearch") = "TRUE" Then
                Select Case Session(SessionVariable_Prefix & "SearchBy")
                    Case "Search"
                        Dim replaceIndex As String
                        If Not String.IsNullOrEmpty(IndexNum) Then
                            Dim ds As DataSet = GetIndexDataSet(Request.QueryString.Get("qid"))
                            Dim changIndex As String = ds.Tables(0).Select("IndexName='" + IndexNum + "'")(0)(1).ToString()
                            replaceIndex = Session(SessionVariable_Prefix & "SearchLabel").ToString().Replace(IndexNum, changIndex)

                        Else
                            replaceIndex = Session(SessionVariable_Prefix & "SearchLabel").ToString()
                        End If

                        lbSearchApplied.Text = replaceIndex
                        ddSearch2.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch2Index")
                        ddSearch3.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch3Index")
                        ddSearch4.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch4Index")
                        ddOwner.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch5Index")

                        SQueueStartDate.Text = Session(SessionVariable_Prefix & "SQueueStartDateParam1Value")
                        EQueueStartDate.Text = Session(SessionVariable_Prefix & "EQueueStartDateParam1Value")
                        'ddSearch1.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch1Index")
                        'tbSearch1.Text = Session(SessionVariable_Prefix & "searchParam1Value")

                    Case "MyClaims"
                        lbSearchApplied.Text = Session(SessionVariable_Prefix & "SearchLabel")
                        ' ddSearch2.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch2Index")
                        ddSearch3.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch3Index")
                        ddSearch4.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch4Index")
                        'ddSearch1.SelectedIndex = Session(SessionVariable_Prefix & "ddSearch1Index")
                        'tbSearch1.Text = Session(SessionVariable_Prefix & "searchParam1Value")

                        gridQueue.CurrentPageIndex = Session(SessionVariable_Prefix & "GridPageIndex")
                        ddSearch2.SelectedIndex = 0 ' Cannot use this value
                        ddOwner.SelectedIndex = 0
                        SQueueStartDate.Text = Session(SessionVariable_Prefix & "SQueueStartDateParam1Value")
                        EQueueStartDate.Text = Session(SessionVariable_Prefix & "EQueueStartDateParam1Value")
                    Case Else
                        'reset all
                        gridQueue.CurrentPageIndex = 0
                        gridQueue.CurrentPageIndex = Session(SessionVariable_Prefix & "GridPageIndex")
                        lbSearchApplied.Text = "Displaying: All " + strDisplayItemName

                        'ddSearch1.SelectedIndex = 0
                        'tbSearch1.Text = ""

                        ddSearch2.SelectedIndex = 0
                        ddSearch3.SelectedIndex = 0
                        ddSearch4.SelectedIndex = 0
                        ddOwner.SelectedIndex = 0
                        SQueueStartDate.Text = ""
                        EQueueStartDate.Text = ""
                End Select
            Else
                'reset all
                gridQueue.CurrentPageIndex = 0
                gridQueue.CurrentPageIndex = Session(SessionVariable_Prefix & "GridPageIndex")
                lbSearchApplied.Text = "Displaying: All " + strDisplayItemName
                'ddSearch1.SelectedIndex = 0
                'tbSearch1.Text = ""
                ddSearch2.SelectedIndex = 0
                ddSearch3.SelectedIndex = 0
                ddSearch4.SelectedIndex = 0
                ddOwner.SelectedIndex = 0
                SQueueStartDate.Text = ""
                EQueueStartDate.Text = ""
            End If
        Catch ex As Exception
            errString = ex.Message
            errLocation = "SetFromSession()"
            '  CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try

    End Sub

    Protected Sub SetDefaultSession()
        Session(SessionVariable_Prefix & "GridPageIndex") = 0
        gridQueue.CurrentPageIndex = 0
        Session(SessionVariable_Prefix & "SearchLabel") = "Displaying: All " + strDisplayItemName
        Session(SessionVariable_Prefix & "IsSearch") = "FALSE"
        Session(SessionVariable_Prefix & "SearchBy") = "ViewAll"

        Session(SessionVariable_Prefix & "cbFilterSchedules") = False
        Session(SessionVariable_Prefix & "cbExcludeClaims") = False

        Session(SessionVariable_Prefix & "ddSearch1Index") = 0
        Session(SessionVariable_Prefix & "searchParam1Value") = ""
        Session(SessionVariable_Prefix & "ddSearch2Index") = 0
        Session(SessionVariable_Prefix & "ddSearch3Index") = 0
        Session(SessionVariable_Prefix & "ddSearch4Index") = 0
        Session(SessionVariable_Prefix & "ddSearch5Index") = 0

        Session(SessionVariable_Prefix & "searchParam2Value") = 0
        Session(SessionVariable_Prefix & "searchParam3Value") = 0
        Session(SessionVariable_Prefix & "searchParam4Value") = 0
        Session(SessionVariable_Prefix & "searchParam5Value") = 0

        Session(SessionVariable_Prefix & "SQueueStartDateParam1Value") = ""
        Session(SessionVariable_Prefix & "EQueueStartDateParam1Value") = ""

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

    Public Function GetCheckBoxImage(ByVal strCheck As String) As String
        Dim strIsChecked As String = "<img src=""Images/checkbox_unchecked.jpg"" alt=""0"" />"
        If Not String.IsNullOrEmpty(strCheck) Then
            If String.Compare("0", strCheck, True) = 0 Or String.Compare("False", strCheck, True) = 0 Then
            Else
                strIsChecked = "<img src=""Images/checkbox_checked.jpg"" alt=""1"" />"
            End If
        Else

        End If
        Dim result As String = strIsChecked

        Return result

    End Function

    Private Sub gridQueue_SetColumnSort(ByVal sender As Object, ByVal e As EventArgs)

        'Dim colCount As Integer = gridQueue.Columns.Count
        'Dim i As Integer = 0
        'Dim nWidth As Integer = 40 / (colCount - 4)
        'Dim colnSort(1, colCount - 1) As String

        'If Page.IsPostBack Or Page.IsCallback Then
        '    ' store sort settings
        '    For Each coloumn As Column In gridQueue.Columns
        '        colnSort(0, i) = coloumn.DataField
        '        colnSort(1, i) = coloumn.SortOrder
        '        i = i + 1
        '    Next
        '    Session(SessionVariable_Prefix & "ColumnSort") = colnSort
        '    If String.IsNullOrEmpty(gridQueue.GroupBy()) Then
        '        Session(SessionVariable_Prefix & "GroupBy") = ""
        '        Session(SessionVariable_Prefix & "GroupExpCollapse") = "False"
        '    Else
        '        Session(SessionVariable_Prefix & "GroupBy") = gridQueue.GroupBy()
        '        Session(SessionVariable_Prefix & "GroupExpCollapse") = gridQueue.ShowCollapsedGroups()
        '    End If
        'Else
        '    If Not Session(SessionVariable_Prefix & "ColumnSort") Is Nothing Then
        '        colnSort = Session(SessionVariable_Prefix & "ColumnSort")
        '        ' store sort settings
        '        For Each coloumn As Column In gridQueue.Columns
        '            If colnSort(0, i) = coloumn.DataField Then
        '                coloumn.SortOrder = colnSort(1, i)
        '            End If
        '            i = i + 1
        '        Next
        '    End If
        '    If Not Session(SessionVariable_Prefix & "GroupBy") Is Nothing Then
        '        gridQueue.GroupBy = Session(SessionVariable_Prefix & "GroupBy")
        '        If String.Compare("Tue", Session(SessionVariable_Prefix & "GroupExpCollapse"), True) = 0 Then
        '            gridQueue.ShowCollapsedGroups = True
        '        Else
        '            gridQueue.ShowCollapsedGroups = False
        '        End If
        '    End If
        'End If

    End Sub

    Public Function LinkToDetail(ByVal strClaimNo As String, ByVal strClaimType As String) As String

        Dim strRet As String = strClaimNo
        If strClaimNo = "" Then
            Return strRet
        End If

        If String.Compare(strClaimType, "C") = 0 Or String.Compare(strClaimType, "D") = 0 Then
            strRet = "<div title='" & strClaimNo & "' id='hover_" & strClaimNo & "'><a class='a.ob_gAL' href='ClaimRoom-ManageClaim.aspx?ClaimNumber=" & strClaimNo & "'> " & strClaimNo & "</a></div>"
        ElseIf String.Compare(strClaimType, "S") = 0 Then
            strRet = "<div title='" & strClaimNo & "' id='hover_" & strClaimNo & "'><a class='a.ob_gAL' href='ClaimRoom-ManageSchedule.aspx?ID=" & strClaimNo & "'> " & strClaimNo & "</a></div>"
        Else
            strRet = strClaimNo
        End If
        Return strRet
    End Function

    '---- Export custom delimited text file ----
    Private Sub ExportGrid()
        Response.Clear()
        Dim j As Integer = 0
        Dim strTemp As String = ""
        Dim ExportTXTFileName As String = "FiledClaimRegister_FullExport.txt"
        For Each col As Column In gridQueue.Columns
            If j > 0 Then
                Response.Write("|")
            End If
            Response.Write(col.HeaderText)
            j += 1
        Next
        For i As Integer = 0 To gridQueue.Rows.Count - 1
            Dim dataItem As Hashtable = gridQueue.Rows(i).ToHashtable()
            j = 0
            Response.Write(vbLf)
            For Each col As Column In gridQueue.Columns
                If j > 0 Then
                    Response.Write("|")
                End If
                If dataItem(col.DataField) Is Nothing Then
                    strTemp = ""
                Else
                    strTemp = dataItem(col.DataField).ToString()
                End If
                Response.Write(strTemp)
                j += 1
            Next
        Next
        Response.AddHeader("content-disposition", "attachment;filename=" & ExportTXTFileName)
        Response.ContentType = "text/plain"
        Response.[End]()
    End Sub

    Protected Sub btnExportCSV_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExportCSV.Click
        ExportGridCSVZip()
    End Sub

    Private Sub ExportGridCSVZip()
        Dim ExportCSVFileName As String = String.Format("Documents-{0}", DateTime.Now.ToString("yyyyMMddHHmmss")) ' "Claims.csv"
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim myConn As New SqlClient.SqlConnection(dbKey)

        Dim STR_WhereClause As String = Session(SessionVariable_Prefix & "WhereClause")
        Dim HT_SearchParams As Hashtable = New Hashtable
        If Not Session(SessionVariable_Prefix & "SearchParams") Is Nothing Then
            HT_SearchParams = DirectCast(Session(SessionVariable_Prefix & "SearchParams"), Hashtable)
        End If

        Dim strFilter As String = CommonUtilsv2.GetDCNRecordFilterByUser(Session("User"))
        Dim strSQL As String = Session(SessionVariable_Prefix & "SelectColumns") ' strExportSQL ' strCommandBase
        If String.IsNullOrEmpty(strFilter) Then
        Else
            strSQL = strSQL & " AND (" & strFilter & ") "
        End If

        If Not String.IsNullOrEmpty(STR_WhereClause) Then
            strSQL = strSQL & STR_WhereClause
        End If
        strSQL = strSQL & strCommandOrderBy
        Dim myComm As New SqlCommand(strSQL, myConn)
        Dim sqlDa As New SqlClient.SqlDataAdapter
        Dim sqlDt As New DataTable
        myComm.CommandTimeout = 0
        If Not String.IsNullOrEmpty(STR_WhereClause) Then
            Dim param As DictionaryEntry
            For Each param In HT_SearchParams
                myComm.Parameters.AddWithValue(param.Key, param.Value)
            Next
        End If
        Try
            sqlDa.SelectCommand = myComm
            sqlDa.Fill(sqlDt)
        Catch ex As Exception
            errString = ex.Message
            errLocation = "ExportGridCSV()"
            CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            sqlDa.Dispose()
            myComm.Dispose()
            myConn.Close()
            myConn.Dispose()
        End Try

        Dim sb As StringBuilder = New StringBuilder
        'index header
        Dim StrQid As Integer = Request.QueryString.Get("qid")
        Dim paramsLog As SqlParameter() = Nothing
        Dim dsIndex As DataSet
        Dim StrInsertAttachment As StringBuilder = New StringBuilder()
        Dim strSQLAttachment As String = "select IndexName,ColumnName from tbl_Web_SearchIndexNames where  GroupID=(select top 1 GroupID from tbl_Web_Queues where ID=@DCN)"
        paramsLog = {New SqlParameter("@DCN", StrQid)}
        dsIndex = CommonUtilsv2.GetDataSet(dbKey, strSQLAttachment, CommandType.Text, paramsLog)

        For Each column As DataColumn In sqlDt.Columns
            'Add the Header row for CSV file.
            If (column.ColumnName.Contains("Index")) Then
                Try
                    Dim IndexName As String = dsIndex.Tables(0).Select(" IndexName='" & column.ColumnName.ToString().Trim() & "'")(0)("ColumnName").ToString()
                    sb.Append(IndexName + ","c)
                Catch

                End Try
            Else
                sb.Append(column.ColumnName + ","c)
            End If

        Next

        'Add new line.
        sb.Append(vbCr & vbLf)

        For Each row As DataRow In sqlDt.Rows
            For Each column As DataColumn In sqlDt.Columns
                'Add the Data rows.
                If String.IsNullOrEmpty(row(column.ColumnName).ToString()) Then
                    If (column.ColumnName.Contains("Index")) Then
                        Try
                            Dim IndexName As String = dsIndex.Tables(0).Select(" IndexName='" & column.ColumnName.ToString().Trim() & "'")(0)("ColumnName").ToString()
                            sb.Append(","c)
                        Catch

                        End Try
                    Else
                        sb.Append(","c)
                    End If
                Else
                    If (column.ColumnName.Contains("Index")) Then
                        Try
                            Dim IndexName As String = dsIndex.Tables(0).Select(" IndexName='" & column.ColumnName.ToString().Trim() & "'")(0)("ColumnName").ToString()
                            sb.Append(row(column.ColumnName).ToString().Replace(",", ";") + ","c)
                        Catch

                        End Try
                    Else
                        sb.Append(row(column.ColumnName).ToString().Replace(",", ";") + ","c)
                    End If
                End If
            Next
            'Add new line.
            sb.Append(vbCr & vbLf)
        Next

        Dim zip As ZipFile = New ZipFile()

        Response.Clear()
        Response.BufferOutput = False
        Response.AddHeader("content-disposition", "attachment;filename=" & ExportCSVFileName & ".zip")
        Response.Charset = ""
        Response.ContentType = "application/zip"
        zip.AddEntry(ExportCSVFileName & ".csv", sb.ToString)
        zip.Save(Response.OutputStream)
        Response.Close()
    End Sub

    Protected Sub UploadDocument(ByVal sender As Object, ByVal e As EventArgs)
        Dim intDocId As Integer = -1
        Dim Filename As String = ""
        Dim lFSizeLimit As Long = 0
        Dim bValidateFile As Boolean = True
        Dim sbValidateMessage As StringBuilder = New StringBuilder()

        If FileUpload1.HasFile = True Then
            Filename = FileUpload1.PostedFile.FileName.Substring(FileUpload1.PostedFile.FileName.LastIndexOf("\") + 1)
            Dim strFileExt As String = FileUpload1.PostedFile.FileName.Substring(FileUpload1.PostedFile.FileName.LastIndexOf("."))
            If InStr(Application("AllowedUploadFiletypes"), strFileExt, CompareMethod.Text) = 0 Then
                sbValidateMessage.AppendLine(" The type " & strFileExt & " of the file to be uploaded is not supported. ")
                bValidateFile = False
            End If
            '      lFSizeLimit = Long.Parse(Application("AllowedUploadFileSize"))
            Long.TryParse(Application("AllowedUploadFileSize"), lFSizeLimit)

            If FileUpload1.PostedFile.ContentLength > lFSizeLimit Then
                sbValidateMessage.AppendLine(" Document size exceeds the" & Application("AllowedUploadFileSizeInMB") & " limit.")
                bValidateFile = False
            End If
        End If
        If textddoctype.Text = 0 Then
            'If ddDocType.SelectedItem.Value = 0 Then
            sbValidateMessage.AppendLine(" Doc Type is required.")
            bValidateFile = False
        End If
        If ddQueue.SelectedIndex = 0 Then
            sbValidateMessage.AppendLine(" Queue is required.")
            bValidateFile = False
        End If
        If bValidateFile = False Then
            sbValidateMessage.AppendLine(" File was not uploaded. ")
            lbSubmitComment.ForeColor = Drawing.Color.Red
            lbSubmitComment.Font.Bold = True
            lbSubmitComment.Text = sbValidateMessage.ToString
            lbSubmitComment.Visible = True
            Exit Sub
        End If
        If Not Filename Is Nothing And Filename <> "" Then
            intDocId = StoreDocument()
            If intDocId = -1 Then
                lbSubmitComment.Text = "File has not been uploaded."
                lbSubmitComment.ForeColor = Drawing.Color.Red
                lbSubmitComment.Visible = True
            Else
                Response.Redirect("/Document-Detail.aspx?id=" + intDocId.ToString())
                'CreateGrid()
                'lbSubmitComment.Text = "File has been uploaded."
                'lbSubmitComment.ForeColor = Drawing.Color.Green
                'lbSubmitComment.Visible = True
            End If
        End If
    End Sub

    Private Function StoreDocument() As Integer
        Dim insSuccess As Boolean = False
        Dim intId As Integer = -1
        Dim strDocumentSize As String = ""
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        If Not FileUpload1.PostedFile Is Nothing AndAlso FileUpload1.PostedFile.ContentLength > 0 Then

            Dim sqlCmd2 As New SqlClient.SqlCommand
            Dim ImageData(FileUpload1.PostedFile.ContentLength - 1) As Byte
            Dim Md5 As New System.Security.Cryptography.MD5CryptoServiceProvider
            Dim SqlCnn2 As New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("dbKey"))
            Dim sqlDa2 As New SqlClient.SqlDataAdapter
            Dim sqlDs2 As New DataSet
            FileUpload1.PostedFile.InputStream.Read(ImageData, 0, FileUpload1.PostedFile.ContentLength)
            Dim rSize As Single = FileUpload1.PostedFile.ContentLength
            rSize = rSize / (1024 * 1024.0)
            strDocumentSize = FormatNumber(rSize, 2, TriState.True, TriState.False, TriState.True)
            sqlCmd2.CommandTimeout = 0
            sqlCmd2.Connection = SqlCnn2
            sqlCmd2.CommandText = "prc_uploadDocument"
            sqlCmd2.CommandType = CommandType.StoredProcedure
            Dim Filename As String = FileUpload1.PostedFile.FileName.Substring(FileUpload1.PostedFile.FileName.LastIndexOf("\") + 1)
            Filename = Replace(Filename, " ", "_")
            Filename = Replace(Filename, "'", "")
            ' Filename=".PDF"
            sqlCmd2.Parameters.AddWithValue("@FileName", Filename)
            sqlCmd2.Parameters.AddWithValue("@ImageBlob", ImageData)
            sqlCmd2.Parameters.AddWithValue("@QueueID", ddQueue.SelectedValue)
            sqlCmd2.Parameters.AddWithValue("@DocType", ddDocType.SelectedValue)
            sqlCmd2.Parameters.AddWithValue("@DocStatus", 1)
            sqlCmd2.Parameters.AddWithValue("@Priority", 2)

            sqlCmd2.Parameters.AddWithValue("@UploadedBy", Session("User"))

            Try
                SqlCnn2.Open()
                sqlDa2.SelectCommand = sqlCmd2
                sqlDa2.Fill(sqlDs2)
                If sqlDs2.Tables(0).Rows.Count > 0 Then
                    Dim dr1 As DataRow
                    Dim dt1 As DataTable
                    dt1 = sqlDs2.Tables(0)
                    dr1 = dt1.Rows(0)
                    intId = dr1(0)
                    If (intId > 0) Then
                        Dim strindex As String
                        Dim Strselect = "select IndexName  from tbl_Web_SearchIndexNames where ColumnName='Load Date' and GroupID=(select  GroupID from tbl_Web_Queues where ID=(select QueueID from tbl_Web_DocumentAttributes where DCN=@DCN))"
                        Dim params As System.Data.SqlClient.SqlParameter() = {
                            New System.Data.SqlClient.SqlParameter("@DCN", intId)
                        }
                        strindex = CommonUtilsv2.RunScalarQuery(dbKey, Strselect, CommandType.Text, params)
                        If Not IsNothing(strindex) Then
                            Strselect = "update  tbl_Web_DocumentAttributes set " & strindex & "= CONVERT(varchar(100), GETDATE(), 101) where DCN=@DCN"
                            Dim params2 As System.Data.SqlClient.SqlParameter() = {
                            New System.Data.SqlClient.SqlParameter("@DCN", intId)
                        }
                            CommonUtilsv2.RunNonQuery(dbKey, Strselect, CommandType.Text, params2)
                        End If
                    End If
                    insSuccess = True
                End If
            Catch ex As SqlClient.SqlException
                errString = ex.Message
                errLocation = "StoreDocument() "
                CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            Finally
                sqlCmd2.Dispose()
                sqlCmd2 = Nothing
                SqlCnn2.Close()
                SqlCnn2.Dispose()
                SqlCnn2 = Nothing
            End Try
        End If
        tbNewFile.Text = ""

        Return intId
    End Function

    Protected Sub MoveDocument(ByVal sender As Object, ByVal e As EventArgs)
        Dim bSuccess As Boolean = True
        Dim bUpdate As Boolean = True
        Dim iDestQueue As Integer = -1
        Dim iDestDocType As Integer = -1

        Dim nMovedDocCount As Integer = 0
        Dim nSelectedCount As Integer = 0
        If Not gridQueue.SelectedRecords Is Nothing Then
            nSelectedCount = gridQueue.SelectedRecords.Count
        End If
        If ddDestinationQueue.SelectedIndex = 0 Then
            lbSubmitComment.Text = "Please select a Destination Queue from the list."
            lbSubmitComment.Visible = True
            lbSubmitComment.ForeColor = Drawing.Color.Red
            Exit Sub
        Else
            lbSubmitComment.Visible = False
            iDestQueue = ddDestinationQueue.SelectedValue
        End If

        'If TextddDestinationDocType.Text = 0 Then
        '    lbSubmitComment.Text = "Please select a Destination Doc Type from the list."
        '    lbSubmitComment.Visible = True
        '    lbSubmitComment.ForeColor = Drawing.Color.Red
        '    Exit Sub
        'Else
        '    lbSubmitComment.Visible = False
        '    iDestDocType = ddDestinationDocType.SelectedValue
        'End If

        If nSelectedCount < 1 Then
            lbSubmitComment.Text = "Please select at least one row from the grid."
            lbSubmitComment.Visible = True
            lbSubmitComment.ForeColor = Drawing.Color.Red
            Exit Sub
        Else
            lbSubmitComment.Visible = False
            lbSubmitComment.Text = ""
        End If

        lbSubmitComment.Visible = False
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim sqlCmd As New SqlClient.SqlCommand

        sqlCmd.CommandTimeout = 0
        sqlCmd.Connection = myConn
        sqlCmd.CommandText = "prc_MoveDocument"
        sqlCmd.CommandType = CommandType.StoredProcedure

        Dim oRecord As Hashtable

        Dim iDCN As Integer = -1
        Try
            Dim strFullChangeLogTable As String = Nothing
            myConn.Open()
            nMovedDocCount = 0
            For Each oRecord In gridQueue.SelectedRecords
                iDCN = oRecord("DCN")

                sqlCmd.Parameters.Clear()
                Try
                    sqlCmd.Parameters.AddWithValue("@DCN", iDCN)
                    sqlCmd.Parameters.AddWithValue("@DestQId", iDestQueue)
                    '  sqlCmd.Parameters.AddWithValue("@DestDocType", iDestDocType)
                    sqlCmd.Parameters.AddWithValue("@UserID", Session("User"))

                    sqlCmd.ExecuteNonQuery()
                    nMovedDocCount = nMovedDocCount + 1
                Catch ex As Exception
                    bSuccess = False
                    Throw ex
                End Try
            Next
        Catch ex As Exception
            bSuccess = False
            errString = ex.Message
            errLocation = "MoveDocument()"
            CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            gridQueue.SelectedRecords.Clear()
            myConn.Close()
            sqlCmd.Parameters.Clear()
        End Try

        If nMovedDocCount > 0 Then
            CreateGrid()
            lbSubmitComment.Text = nMovedDocCount.ToString & " documents have been moved to " & ddDestinationQueue.SelectedItem.Text & " queue."
            lbSubmitComment.ForeColor = Drawing.Color.Blue
            lbSubmitComment.Visible = True
        Else
            lbSubmitComment.Text = "Unable to move any of the " + nSelectedCount.ToString + " selected document(s) to <b>" + ddDestinationQueue.SelectedItem.Text + " queue</b>."
            lbSubmitComment.Text = " <br/> Ensure you select at least one document from the grid above. "
            lbSubmitComment.ForeColor = Drawing.Color.Red
            lbSubmitComment.Visible = True
        End If

    End Sub

    Protected Sub AssignOwner(ByVal sender As Object, ByVal e As EventArgs)
        Dim bSuccess As Boolean = True
        Dim bUpdate As Boolean = True

        Dim nAssignCount As Integer = 0
        Dim nSelectedCount As Integer = 0
        If Not gridQueue.SelectedRecords Is Nothing Then
            nSelectedCount = gridQueue.SelectedRecords.Count
        End If
        If ddDocUser.SelectedIndex = 0 Then
            lbSubmitComment.Text = "Please select an User."
            lbSubmitComment.Visible = True
            lbSubmitComment.ForeColor = Drawing.Color.Red
            Exit Sub
        Else
            lbSubmitComment.Visible = False
        End If

        If nSelectedCount < 1 Then
            lbSubmitComment.Text = "Please select at least one row from the grid."
            lbSubmitComment.Visible = True
            lbSubmitComment.ForeColor = Drawing.Color.Red
            Exit Sub
        Else
            lbSubmitComment.Visible = False
            lbSubmitComment.Text = ""
        End If

        lbSubmitComment.Visible = False
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim sqlCmd As New SqlClient.SqlCommand

        sqlCmd.CommandTimeout = 0
        sqlCmd.Connection = myConn
        sqlCmd.CommandText = "prc_AssignOwner"
        sqlCmd.CommandType = CommandType.StoredProcedure

        Dim oRecord As Hashtable

        Dim iDCN As Integer = -1
        Try
            Dim strFullChangeLogTable As String = Nothing
            myConn.Open()
            nAssignCount = 0
            For Each oRecord In gridQueue.SelectedRecords
                iDCN = oRecord("DCN")

                sqlCmd.Parameters.Clear()
                Try
                    sqlCmd.Parameters.AddWithValue("@DCN", iDCN)
                    sqlCmd.Parameters.AddWithValue("@DocOwner", ddDocUser.SelectedItem.Text)
                    sqlCmd.Parameters.AddWithValue("@UserID", Session("User"))

                    sqlCmd.ExecuteNonQuery()
                    nAssignCount = nAssignCount + 1
                Catch ex As Exception
                    bSuccess = False
                    Throw ex
                End Try
            Next
        Catch ex As Exception
            bSuccess = False
            errString = ex.Message
            errLocation = "AssignOwner()"
            CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            gridQueue.SelectedRecords.Clear()
            myConn.Close()
            sqlCmd.Parameters.Clear()
        End Try

        If nAssignCount > 0 Then
            CreateGrid()
            lbSubmitComment.Text = nAssignCount.ToString & " documents have been assigned to user: " & ddDocUser.SelectedItem.Text & "."
            lbSubmitComment.ForeColor = Drawing.Color.Blue
            lbSubmitComment.Visible = True
        Else
            lbSubmitComment.Text = "Unable to assign user to any of the " + nSelectedCount.ToString + " selected document(s) to <b>"
            lbSubmitComment.Text = " <br/> Ensure you select at least one document from the grid above. "
            lbSubmitComment.ForeColor = Drawing.Color.Red
            lbSubmitComment.Visible = True
        End If

    End Sub

    Public Function GetIndexDataSet(ByVal queueID As Integer) As DataSet
        ' status 0:error,1:no data,2:success
        Dim ReturnStr As String = String.Empty
        Dim ds As DataSet

        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim sp As String = "prc_GetTableIndex"
        Dim params As System.Data.SqlClient.SqlParameter() = {
        New System.Data.SqlClient.SqlParameter("@queueId", queueID)
        }


        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            Return Webapps.Utils.CommonUtilsv2.GetDataSet(dbKey, sp, System.Data.CommandType.StoredProcedure, params)

        Catch ex As Exception

            Return Nothing

        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try
    End Function

    <WebMethod()>
    Public Shared Function GetDocType(ByVal queueID As Integer) As String
        Dim StrUser As String = HttpContext.Current.Session("User")
        If StrUser Is Nothing Then
            Return "{""msg"":""You don't have access to this Function "",""status"":-1}"
        End If
        ' status 0:error,1:no data,2:success
        Dim ReturnStr As String = String.Empty
        Dim ds As DataSet
        Dim params As SqlParameter() = Nothing

        Dim strSQL As String = " select ID as DocTypeID,DocTypeName from tbl_Web_DocTypes where GroupID=(select GroupID from tbl_Web_Queues where ID=@queueID) order by SortOrder "
        params = {New SqlParameter("@queueID", queueID)}
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")


        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)
            If Not ds Is Nothing Then
                Dim strjson As String = String.Empty
                For Each row As System.Data.DataRow In ds.Tables(0).Rows
                    strjson += "{""Index"":""" + row("DocTypeID").ToString() + """,""Name"":""" + row("DocTypeName") + """},"

                Next
                ReturnStr = "{""msg"":""this success!"",""status"":2,""data"":" + "[" & strjson.ToString().TrimEnd(",") & "]" + "}"

            Else

                ReturnStr = "{""msg"":""this no data!"",""status"":1}"

            End If

        Catch ex As Exception

            ReturnStr = "{""msg"":""" + ex.ToString() + """,""status"":0}"

        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try




        Return ReturnStr
    End Function
</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
<script language="javascript" type="text/javascript" src="includes/js/My97DatePicker/WdatePicker.js"></script>
    <script type="text/javascript">
        function expandBlock(blockName) {
            var blockID = blockName + '_block';
            var expandID = blockName + '_expand';
            var collapseID = blockName + '_collapse';

            if (document.getElementById(blockID)) {
                document.getElementById(blockID).style.display = "inline";
            }
            if (document.getElementById(expandID)) {
                document.getElementById(expandID).style.display = "none";
            }
            if (document.getElementById(collapseID)) {
                document.getElementById(collapseID).style.display = "inline";
            }
        }
        function collapseBlock(blockName) {
            var blockID = blockName + '_block';
            var expandID = blockName + '_expand';
            var collapseID = blockName + '_collapse';

            if (document.getElementById(blockID)) {
                document.getElementById(blockID).style.display = "none";
            }
            if (document.getElementById(expandID)) {
                document.getElementById(expandID).style.display = "inline";
            }
            if (document.getElementById(collapseID)) {
                document.getElementById(collapseID).style.display = "none";
            }
        }
        function ClaimsRoomExport(eType) {
            window.open("H_CreateDownload.ashx?filename=" + eType);
            return false;
        }
        function setFileName(fuID, tbID) {
            var indFileName = document.getElementById(fuID).value.split(',');
            var arrFileName = [];
            var concatFileName = "";

            if (indFileName) {
                for (var i = 0; i < indFileName.length; i++) {
                    arrFileName.push(indFileName[i].split('\\'));
                }

                if (arrFileName) {
                    for (var i = 0; i < arrFileName.length; i++) {
                        var tempStr = arrFileName[i];
                        concatFileName += tempStr[tempStr.length - 1] + ',';
                    }
                }
            }

            if (concatFileName.length > 1) {
                concatFileName = concatFileName.slice(0, -1);
            }
            document.getElementById(tbID).value = concatFileName;
        }
        function backfunction() {
            if (window.history.length > 1) {
                window.history.go(-1);
            } else {
                window.opener.location.reload();
            }
        }
        $(function () {
            SetGridScrollWidth();
            let queueId = getUrlParam('qid');
            //UpdateTablehead(queueId);
            GetDocTypList(queueId);
            $("#ctl00_Body_ddQueue").on("change", function () {
                var selectText = $(this).find('option:selected').val();
                GetDocTypList(selectText);
                $("#ctl00_Body_textddoctype").val("0");
            })
            
            $("#ctl00_Body_ddDocType").on("change", function () {
                $("#ctl00_Body_textddoctype").val($(this).find('option:selected').val());
            })
           
            $(".Wdate").on("change", function () {               
                $(this).val($(this).val());
            })
            $("#ctl00_Body_Table1").on("change", '.selectstate', function () {
                var strid = $(this).find("option:selected").val();
                if (strid) {
                    $.ajax({
                        type: "Post",
                        contentType: "application/json;charset=UTF-8",
                        url: "/Document-Detail.aspx/GetCountiesSelect",
                        data: "{ID:'" + strid + "'}",
                        success: function (result) {
                            var strjson = JSON.parse(result.d);
                            if (strjson.status == -1) {
                                window.location.href = "/login.aspx";
                            }
                            var strchildoption = "<option value=''>select a value</option>";
                            if (strjson.status == 2) {
                                $(strjson.data).each(function (i, dom) {
                                    strchildoption += "<option value='" + dom.Name + "'>" + dom.Name + "</option>";
                                });
                            }
                            $(".selectcounty").html(strchildoption);
                        }
                    });
                } else {
                    var strchildoption = "<option value=''>select a value</option>";
                    $(".selectcounty").html(strchildoption);
                }
            });

            $(".Wdate").attr("readonly", "readonly");
        });

        function getUrlParam(name) {
            var reg = new RegExp("(^|&)" + name + "=([^&]*)(&|$)"); 
            var r = window.location.search.substr(1).match(reg);  
            if (r != null) return unescape(r[2]); return null; 
        }
        
        function GetDocTypList(queueId) {
            $.ajax({
                type: "Post",
                contentType: "application/json;charset=UTF-8",
                url: "/QueueRegister.aspx/GetDocType",
                data: "{queueID:" + queueId + "}",
                success: function (result) {
                    var strjson = JSON.parse(result.d);
                    if (strjson.status == -1) {
                        window.location.href = "/login.aspx";
                    }
                    if (strjson.status == 2) {
                        var selectList="<option value=\"0\">Select A Value </option>";
                        $(strjson.data).each(function (i, dom) {
                            selectList += "<option value=\"" + dom.Index + "\">" + dom.Name + "</option>"
                        });
                        $("#ctl00_Body_ddDocType").html(selectList);

                    }
                }
            });
        }
       
        function getWindowWidth() {
            var windowWidth = 0;
            if (typeof (window.innerWidth) == 'number') {
                windowWidth = window.innerWidth;
            }
            else {
                if (document.documentElement && document.documentElement.clientWidth) {
                    windowWidth = document.documentElement.clientWidth;
                }
                else {
                    if (document.body && document.body.clientWidth) {
                        windowWidth = document.body.clientWidth;
                    }
                }
            }
            return windowWidth;
        }

        function SetGridScrollWidth() {
            var scrollWidth = ($("#mainContent").outerWidth()-15) + 'px';
            var boublescrollWidth = $(".ob_gBody").outerWidth() + 2;
            gridQueue.GridMainContainer.style.width = scrollWidth;
            $("#ctl00_Body_Table2").attr("style", "width:" + ($("#mainContent").outerWidth() - 7) + "px;margin-right: 7px;")
            $("#ctl00_Body_gridQueue_ob_gridQueueHS").attr("style", "height:17px");
            $("#ctl00_Body_gridQueue_ob_gridQueueHS").children("div").attr("style", "height:37px");
            $("#ctl00_Body_gridQueue_ob_gridQueueHS").children("div").children("div").attr("style", "width:" + boublescrollWidth + "px");
        }

        function doclearsearch()
        {
            $("#ctl00_Body_tblDisplay").find("select").attr('value', '');
            $("#ctl00_Body_dynamicIndexs").find("input").val("");
            $("#ctl00_Body_dynamicIndexs").find("select").val('');
        }

    </script>   
    <input type="hidden"  id="txtTextCount" runat="server" />
    <div align="center">
        <div class="body_title">
            <a href="javascript:void(0);" onclick="backfunction()"><span>back</span></a>
            Document Queue
        </div>
        <table id="Table1" width="100%"  border="0" cellpadding="0" cellspacing="0" runat="server" style="padding-right: 10px;">
            <tr class="header_spacer">
                <td colspan="6">&nbsp;</td>
            </tr>
            <tr>
                <td colspan="6" style="border-bottom: solid 1px #00437a;">
                    <div title="Click to Expand" id="search_expand" style="display: none;">
                        <img alt="" src="Images/expand.gif" width="12" height="12" onclick="expandBlock('search');" />
                    </div>
                    <div title="Click to Hide" id="search_collapse" style="display: inline;">
                        <img alt="" src="Images/collapse.gif" width="12" height="12" onclick="collapseBlock('search');" />
                    </div>
                    <div class="form_label_medium_nobackground" style="display: inline;"><b>Search & Filter</b></div>
                </td>
            </tr>

            <tr class="popupbody_simple">
                <td colspan="6" align="center">
                    <div id="search_block" style="width: 98%;margin-top: 5px;">
                        <table id="tblDisplay" width="100%"  border="0" cellpadding="0" cellspace="0" runat="server" >
                            <tr class="body_search">   
                                 <td align="right" width="10%">
                                    <% If (Isprovdier > 0) Then %>
                                    <label>Complete Date:</label>&nbsp;
                                    <% Else%>
                                    <label>Doc Status:</label>&nbsp;
                                    <%   End If %>
                                    
                                </td>
                                <td align="left" valign="bottom" width="23%" class="tdPadding" style="vertical-align: middle;">
                                    <% If (Isprovdier = 2) Then %>
                                    <asp:TextBox id="SQueueStartDate" MaxLength="30" style="height:25px;width:35%;" CssClass="Wdate" onclick="WdatePicker({lang:'en',dateFmt:'MM/dd/yyyy'})" runat="server"></asp:TextBox><span> To </span>
                                    <asp:TextBox id="EQueueStartDate" MaxLength="30" style="height:25px;width:35%;" CssClass="Wdate" onclick="WdatePicker({lang:'en',dateFmt:'MM/dd/yyyy'})" runat="server"></asp:TextBox>
                                    <% ElseIf (Isprovdier = 1) Then %>
                                        <asp:TextBox id="TextBox1" MaxLength="30" style="height:25px;width:35%;background: #cdcaca  url(/includes/js/My97DatePicker/skin/datePicker.gif) no-repeat right;" CssClass="Wdate"  runat="server"  disabled="disabled" ></asp:TextBox><span> To </span>
                                        <asp:TextBox id="TextBox2" MaxLength="30" style="height:25px;width:35%;background: #cdcaca  url(/includes/js/My97DatePicker/skin/datePicker.gif) no-repeat right;" CssClass="Wdate"  runat="server" disabled="disabled" ></asp:TextBox>
                                    <% Else%>
                                    <asp:DropDownList ID="ddSearch3" Width="78%" Height="25px" runat="server" class="select-css" ></asp:DropDownList>
                                    <%   End If %>
                                    
                                </td>
                                <td align="right" width="10%">
                                    <label>Doc Type:</label>&nbsp;
                                </td>
                                <td align="left" valign="bottom" width="24%" class="tdPadding" style="vertical-align: middle;">
                                     <asp:DropDownList ID="ddSearch2" Width="78%" Height="25px" runat="server" class="select-css" ></asp:DropDownList>
                                </td>
                               
                                <td align="right" width="10%">
                                    <% If (Isprovdier > 0) Then %>
                                    <label>Owner:</label>&nbsp;
                                    <% Else%>
                                    <label>Priority:</label>&nbsp;
                                    <%   End If %>
                                </td>
                                <td align="left" valign="bottom" width="23%" class="tdPadding" style="vertical-align: middle;">
                                    <% If (Isprovdier > 0) Then %>
                                    <asp:DropDownList ID="ddOwner"  Width="78%" Height="25px" runat="server" class="select-css"></asp:DropDownList>
                                    <%--<% ElseIf (Isprovdier = 1) Then %>
                                        <asp:DropDownList ID="DropDownList2" Width="80%" Height="25px" runat="server" class="select-css" ReadOnly ="ReadOnly" style="background: #d5d2d2;" >
                                            <asp:ListItem Value="">Select a Value</asp:ListItem>
                                        </asp:DropDownList>--%>
                                    <% Else%>
                                     <asp:DropDownList ID="ddSearch4" Width="78%" Height="25px" runat="server" class="select-css" ></asp:DropDownList>
                                    <%   End If %>
                                   
                                </td>
                            </tr>    
                        </table>                       
                        <asp:Table ID="dynamicIndexs" runat="server" width="100%"  border="0" cellpadding="0" cellspace="0"></asp:Table>
                        <table id="Table3" width="100%"  border="0" cellpadding="0" cellspace="0" runat="server">
                         <tr class="body_search">
                                <td width="15%" style="font-size: 16px;">
                                   <asp:CheckBox ID="mydocument" name="mydocument"  runat="server" Text="My Document"  style="margin-left:50px;"  />
                                </td>
                                <td align="right"  style="padding-right: 10px;">
                                    <asp:Button ID="btn_Search" class="Submit_button" Style="width: 90px;" runat="server" Text="Search" />&nbsp;
                                    <asp:Button ID="btn_ViewAll" class="Submit_button" Style="width: 90px;" runat="server" Text="View All" Onmousedown="doclearsearch()" />
                                </td>
                            </tr>
                            <tr>
                                <td align="left" colspan="6" class="body_search">
                                    <asp:Label ID="lbSearchApplied" runat="server" ForeColor="#5cb335" Font-Bold="true" Visible="True" />
                                </td>                              
                            </tr>
                            <tr>
                            <td align="right" width="80%" colspan="4">
                                <asp:Label ID="ExportLabel" runat="server" Visible="false" ForeColor="DarkRed" Text="Please narrow down the search as the number of filtered records exceeds the export limit." />
                            </td>
                                <td align="right" width="20%" class="body_search" colspan="2">
                                    <div id="dvExportExcel" runat="server" style="margin-right: 20px;">
                                        <asp:LinkButton ID="btnExportCSV" runat="server" Visible="true" Text="Export Current View" />
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
        </table>

        <table id="Table2"  runat="server" border="0" cellpadding="0" cellspace="0"  style="width: 1600px;">
            <tr>
                <td width="100%" align="left" valign="top">
                    <obout:Grid ID="gridQueue" runat="server" CallbackMode="true" Serialize="false" AutoGenerateColumns="false"
                        PageSizeOptions="1,5,10,15,20,25,30,35,40,45,50,100" PageSize="15" Width="100%"
                         FolderStyle="styles/grand_graydark" EnableRecordHover="true"
                        AllowRecordSelection="true" AllowMultiRecordSelection="true" KeepSelectedRecords="false"
                        AllowAddingRecords="false" AllowFiltering="false" ShowLoadingMessage="true" AllowSorting="true"
                        AllowGrouping="false" ShowGroupsInfo="false" ShowCollapsedGroups="false" ShowMultiPageGroupsInfo="false"
                        ShowColumnsFooter="false" ShowGroupFooter="false" OnRowDataBound="RowDataBound" OnRebind="RebindGrid"
                        OnColumnsCreated="gridQueue_SetColumnSort">
                        <Columns>
                            <obout:CheckBoxSelectColumn ShowHeaderCheckBox="true" ControlType="Standard" Width="40"></obout:CheckBoxSelectColumn>
                            <obout:Column ID="Column101" HeaderText="DCN" DataField="DCN" Align="center" Width="0%" Visible="false" Wrap="true" AllowGroupBy="false" runat="server" HeaderAlign="center" ReadOnly="true" />
                            <obout:Column ID="Column2" HeaderText="DCN #" DataField="" TemplateId="tmplDocDetail" Align="center" ReadOnly="true" Width="120" Visible="true" Wrap="true" AllowGroupBy="false" runat="server" HeaderAlign="center" AllowSorting="true" />
                            <obout:Column ID="Column3" HeaderText="Attachment Count" DataField="attachmentcount" Align="center" Width="100" Visible="true" Wrap="true"  runat="server" HeaderAlign="center"  ReadOnly="true" />
                            <obout:Column ID="Column102" HeaderText="PriorityID" DataField="PriorityID" Align="left" Width="0%" Visible="false" Wrap="false" AllowGroupBy="false" runat="server" HeaderAlign="center" />
                            <obout:Column ID="Column1" HeaderText="Owner" DataField="Owner" Align="left" Width="120" Visible="true" Wrap="false" AllowGroupBy="false" runat="server" HeaderAlign="center" />
                            <obout:Column ID="Column104" HeaderText="DocStatusID" DataField="DocStatusID" Align="left" Width="120" Visible="false" Wrap="false" AllowGroupBy="false" runat="server" HeaderAlign="center" />
                            <obout:Column ID="Column105" HeaderText="Doc Status" DataField="DocStatusName" Align="left" Width="120" Visible="true" Wrap="true" runat="server" HeaderAlign="center" ReadOnly="true" />
                            <obout:Column ID="Column106" HeaderText="DocTypeID" DataField="DocTypeID" Align="left" Width="0%" Visible="false" Wrap="false" AllowGroupBy="false" runat="server" HeaderAlign="center" />
                            <obout:Column ID="Column107" HeaderText="Doc Type" DataField="DocTypeName" Align="left" Width="120" Visible="true" Wrap="true" runat="server" HeaderAlign="center" ReadOnly="true" />
                            <obout:Column ID="Column108" DataField="UploadedDate" HeaderText="Uploaded Date" Align="left" Width="120" Visible="true" Wrap="true" AllowGroupBy="false" runat="server" HeaderAlign="center" ReadOnly="true" DataFormatString="{0:MM/dd/yyyy HH:mm}" />
                            <obout:Column ID="Column109" DataField="QueueAgeInSec" HeaderText="Queue Age"  TemplateId="tmplQueueAge" Align="left" Width="120" Visible="true" Wrap="true" AllowGroupBy="false" AllowSorting="true" runat="server" HeaderAlign="center" ReadOnly="true" />
                            <obout:Column ID="Column110" DataField="DocAgeInSec" HeaderText="Doc Age" TemplateId="tmplDocAge" Align="left" Width="120" Visible="true" Wrap="true" AllowGroupBy="false" AllowSorting="true"  runat="server" HeaderAlign="center" ReadOnly="true" />
                            <obout:Column ID="Column4" DataField="QueueAge" HeaderText="QueueAge" Align="left" Width="120" Visible="false" Wrap="true" AllowGroupBy="false" runat="server" HeaderAlign="center" ReadOnly="true" />
                            <obout:Column ID="Column5" DataField="DocAge" HeaderText="DocAge" Align="left" Width="120" Visible="false" Wrap="true" AllowGroupBy="false" runat="server" HeaderAlign="center" ReadOnly="true" />
                            <obout:Column ID="Column6" DataField="ImageArchived" HeaderText="ImageArchived" Align="left" Width="120" Visible="false" Wrap="true" AllowGroupBy="false" runat="server" HeaderAlign="center" ReadOnly="true" />
                        </Columns>
                         <ScrollingSettings ScrollWidth="900" />
                        <Templates>
                            <obout:GridTemplate runat="server" ID="tplEditReviewer" ControlID="ddReviewer" ControlPropertyName="value">
                                <Template>
                                    <asp:DropDownList runat="server" ID="ddReviewer" DataValueField="ReviewerID" DataTextField="ReviewerName" Height="20px" CssClass="ob_gEC" />
                                </Template>
                            </obout:GridTemplate>
                            <obout:GridTemplate runat="server" ID="updateBtnTemplate">
                                <Template>
                                    <a class="ob_gAL" href="javascript: //" onclick="gridQueue.update_record(this);this.blur();return false;">Update</a>
                                    <a class="ob_gAL" href="javascript: //" onclick="gridQueue.cancel_edit(this);this.blur();return false;">Cancel</a>
                                </Template>
                            </obout:GridTemplate>

                            <obout:GridTemplate runat="server" ID="tmplDownloadDoc">
                                <Template>
                                    <div title="C" id='hover_<%# Container.DataItem("DCN")%>'><a class="a.ob_gAL" href="Download-Document.aspx?id=<%# Container.DataItem("DCN")%>&IArchived=<%# Container.DataItem("ImageArchived")%>"><%# Container.DataItem("DCN")%></a></div>
                                </Template>
                            </obout:GridTemplate>
                            <obout:GridTemplate runat="server" ID="tmplDocDetail">
                                <Template>
                                    <div title="<%# Container.DataItem("DCN")%>" id='hover_<%# Container.DataItem("DCN")%>'>
                                        <a class="a.ob_gAL" href="Document-Detail.aspx?id=<%# Container.DataItem("DCN")%>"><%# Container.DataItem("DCN")%></a> &nbsp;<a href='download-Document.aspx?id=<%# Container.DataItem("DCN")%>&IArchived=<%# Container.DataItem("ImageArchived")%>'>
                                            <img src='Images/download16.png' style='width:12px;'></a>
                                    </div>
                                </Template>
                            </obout:GridTemplate>
                            <obout:GridTemplate runat="server" ID="tmplTransferredChkBox">
                                <Template>
                                    <%#GetCheckBoxImage(Container.Value())%>
                                </Template>
                            </obout:GridTemplate>
                            <obout:GridTemplate runat="server" ID="tmplChkBox">
                                <Template>
                                    <%#GetCheckBoxImage(Container.Value())%>
                                </Template>
                            </obout:GridTemplate>
                              <obout:GridTemplate runat="server" ID="tmplQueueAge">
                                <Template>
                                   <%# Container.DataItem("QueueAge")%>
                                </Template>
                            </obout:GridTemplate>
                            <obout:GridTemplate runat="server" ID="tmplDocAge">
                                <Template>
                                   <%# Container.DataItem("DocAge")%>
                                </Template>
                            </obout:GridTemplate>
                        </Templates>
                    </obout:Grid>
                </td>
            </tr>
        </table>
        
       <% If Not IsDeleteGroup Then %>
        <div runat="server" id="dvUpload" align="left" class="content-modfi">
            <span style="display: block; width: 150px; height: 20px; position: relative; top: -12px; text-align: center; background: white; color: #0075c0; font-size: 16px;">Upload Document</span>
            <table width="100%">
                <tr class="popupbody_simple">
                    <td align="right" valign="top" colspan="1" width="170px">Upload Document:</td>
                    <td align="left" valign="top" style="width: 350px;">
                        <asp:TextBox ID="tbNewFile" runat="server" Text="" Width="330px" Height="25px" ReadOnly="true"></asp:TextBox>
                    </td>
                    <td>
                        <label class="file-upload">
                            <span style="vertical-align: top;"><strong>Browse</strong></span>
                            <asp:FileUpload ID="FileUpload1" runat="server" onchange="setFileName()"></asp:FileUpload>
                        </label>
                    </td>
                </tr>
                <tr class="popupbody_simple">
                    <td align="right" class="popupbody_simple" style="height: 23px">Queue:</td>
                    <td align="left" valign="top" class="popupbody_simple" style="height: 23px">
                        <asp:DropDownList ID="ddQueue" runat="server" Width="330px" AppendDataBoundItems="true" TabIndex="101" Height="25px"></asp:DropDownList>
                    </td>
                    <td style="height: 23px"></td>
                </tr>
                <tr class="popupbody_simple">
                    <asp:TextBox ID="textddoctype" runat="server" type="hidden" value="0"></asp:TextBox>
                    <td align="right" class="popupbody_simple">Doc Type:</td>
                    <td align="left" valign="top" class="popupbody_simple">
                        <asp:DropDownList ID="ddDocType" runat="server" Width="330px" AppendDataBoundItems="true" TabIndex="102" Height="25px"></asp:DropDownList>
                        &nbsp;&nbsp; 
                    </td>
                    <td>
                        <asp:Button ID="btnUploadDoc" runat="server" Text="Upload Document" CssClass="Submit_button" OnClick="UploadDocument" TabIndex="105" CausesValidation="false" ValidationGroup="uploadAction" />
                    </td>

                </tr>
            </table>
        </div>
        <% End If %>
        <div runat="server" id="dvBulkMove" align="left" class="content-modfi">
            <span style="display: block; width: 150px; height: 20px; position: relative; top: -12px; text-align: center; background: white; color: #0075c0; font-size: 16px;">Move Document</span>
            <table width="100%">
                <tr>
                    <td align="right" valign="top" colspan="1" width="170px">Destination Queue:</td>
                    <td align="left" valign="top" style="width: 350px;">
                        <asp:DropDownList ID="ddDestinationQueue" runat="server" Width="330px" AppendDataBoundItems="true" TabIndex="120" Height="25px"></asp:DropDownList>
                    </td>
                    <td>
                        <asp:Button ID="btnMove" runat="server" Text="Move Document" CssClass="Submit_button" OnClick="MoveDocument" TabIndex="125" CausesValidation="false" ValidationGroup="MoveAction" />
                    </td>
                </tr>
            </table>
        </div>
        <div runat="server" id="dvBulkAssign" align="left" class="content-modfi">
            <span style="display: block; width: 150px; height: 20px; position: relative; top: -12px; text-align: center; background: white; color: #0075c0; font-size: 16px;">Assign Owner</span>
            <table width="100%">
                <tr class="popupbody_simple">
                    <td align="right" valign="top"  width="170px">User:</td>
                    <td align="left" valign="top" style="width: 350px;">
                        <asp:DropDownList ID="ddDocUser" runat="server" Width="330px" AppendDataBoundItems="true" TabIndex="211" Height="25px"></asp:DropDownList>
                       
                    </td>
                </tr>
               <tr class="popupbody_simple">
                   <td></td>
                    <td align="left" valign="top"> <asp:CheckBox ID="cbSendEMail" Text="Send eMail Notification" Checked="false" runat="server" TabIndex="221" /></td> 
                    <td>
                        <asp:Button ID="btnAssign" runat="server" Text="Assign Owner" CssClass="Submit_button" OnClick="AssignOwner" TabIndex="223" CausesValidation="false" ValidationGroup="AssignAction" />
                    </td>
                </tr>
            </table>
        </div>
        <div runat="server" id="Div1" align="left">
            <table>
                <tr> 
                    <td align="left" valign="middle" colspan="3">
                        <asp:Label ID="lbSubmitComment" runat="server" CssClass="body_text_errors" Visible="false" /></td>
                </tr>
            </table>
        </div>

    </div>
    <script type="text/javascript">
        window.onresize = SetGridScrollWidth;
    </script>
</asp:Content>
