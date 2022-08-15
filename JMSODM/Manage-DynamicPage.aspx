<%@ Page Language="VB" MasterPageFile="PageMaster.master" AutoEventWireup="false" EnableEventValidation="false" viewStateEncryptionMode="Auto" ASPCOMPAT="TRUE" %>
<%@ MasterType VirtualPath="PageMaster.master" %>
<%@ Register Assembly="Obout.Ajax.UI" Namespace="Obout.Ajax.UI.HTMLEditor" TagPrefix="obout" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.Page" %>
<%@ Import Namespace="Webapps.Utils" %>

<script language="VB" runat="server">
    '************* Error logging Section ***********************
    Dim pageName As String
    Dim errLocation As String
    Dim errString As String
    Dim strLanguageCode As String = "en-us"

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
        If String.IsNullOrEmpty(CustomRoles.GetLanguageSubTag()) Then
        Else
            strLanguageCode = CustomRoles.GetLanguageSubTag()
        End If

        Session("PreviousPage") = "SiteManager.aspx"
        Dim strLogout As String = "False"
        strLogout = Request.QueryString.Get("logout")
        If strLogout = "True" Then
            Session("Level") = "0"
        End If

        If CustomRoles.RolesForPageLoad() Then
            If Not IsPostBack Then
                ddLanguage.Items.Insert(0, New ListItem("Select a Language", 0))
                ddLanguage.Items.Insert(1, New ListItem("English", "en-us"))
                'ddLanguage.Items.Insert(2, New ListItem("Japanese", "ja"))
                Me.ddEditDocument.SelectedIndex = 0
                Me.ddPageName.SelectedIndex = 0
                GetCurrentHomePageContent()
                LoadDocumentList()
                LoadPageNameList()
            Else
            End If
        Else
            CustomRoles.TransferIfNotInRole()
            Exit Sub
        End If
    End Sub

    Private Sub LoadDropDownBox(ByVal dbCNStr As String, ByVal strSQL As String, ByVal dd As DropDownList, ByVal dataValueField As String, ByVal dataTextField As String)
        Try
            CommonUtilsv2.PopulateDropDownBox(dbCNStr, strSQL, dd, dataValueField, dataTextField)
        Catch ex As Exception
            errString = "SQLString: " & strSQL & ". " & ex.Message()
            errLocation = "PopulateDropDownBox: " & dd.ID
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try
    End Sub

    Dim strImageTableName As String = "[tbl_Web_Page_Content]"
    Dim strCommandBase As String = String.Format("SELECT [ID], [Reference_ID],FileName, Description, InfoType, Comments, isActive, DateUploaded, Document_Size_In_MB as [ReportSize] FROM {0}", strImageTableName)
    Dim strWhereClause As String = " WHERE 1=1  " ' AND [isActive] =1 AND now() between [EffectiveDate] and [ExpirationDate] 
    Dim strCommandOrderBy As String = " ORDER BY [ID] ASC"

    Private Sub LoadDocumentList()
        Dim strPage As String = ddPageName.SelectedValue
        Dim strLangCode As String = ddLanguage.SelectedValue

        If CommonUtilsv2.Validate(strPage, CommonUtilsv2.DataTypes.String, True, True, False, 50) = False Then
            Exit Sub
        End If
        If CommonUtilsv2.Validate(strLangCode, CommonUtilsv2.DataTypes.String, True, True, False, 50) = False Then
            Exit Sub
        End If

        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strSQLSubStatus As String = Nothing
        strSQLSubStatus = "SELECT ID, '# ' + CAST(ID As Varchar(3)) + ' Updated on '+CAST([DateLastUpdated] AS Varchar(20))+' by '+[RecordUpdatedBy] as [Report] FROM tbl_Web_Page_Content WHERE RecordStatus = 1 AND PageName = @PageName AND LanguageCode= @LanguageCode ORDER BY [DateLastUpdated] DESC"

        Dim params As SqlParameter() = { _
                New SqlParameter("@PageName", strPage), _
                New SqlParameter("@LanguageCode", strLangCode)}
        CommonUtilsv2.LoadDropDownBox(dbKey, strSQLSubStatus, Me.ddEditDocument, "ID", "Report", params)
        ddEditDocument.Items.Insert(0, New ListItem("Select a Content Set", 0))
        Me.ddEditDocument.SelectedIndex = 0
    End Sub

    Private Sub LoadPageNameList()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim params As SqlParameter() = Nothing
        Dim strSQLPageName As String = "SELECT distinct LookupCode FROM tbl_WEB_Lookup WHERE LookupType =@LookupType And RecordStatus=1 ORDER BY LookupCode"
        params = {New SqlParameter("@LookupType", "WebPageName")}
        CommonUtilsv2.LoadDropDownBox(dbKey, strSQLPageName, Me.ddPageName, "LookupCode", "LookupCode", params)
        ddPageName.Items.Insert(0, New ListItem("Select a Page Name", 0))
        If Not IsPostBack Then
            ddPageName.SelectedIndex = 0
        End If
    End Sub

    Private Sub LoadLanguageList()
        'Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        'Dim strSQLLanguages As String = String.Format("SELECT distinct LanguageCode FROM tbl_Web_Page_Content WHERE RecordStatus = 1 AND PageName = '{0}' ORDER BY [LanguageCode] ", ddPageName.SelectedValue)
        'LoadDropDownBox(dbKey, strSQLLanguages, Me.ddLanguage, "LanguageCode", "LanguageCode")
        ddLanguage.Items.Insert(0, New ListItem("Select a Language", 0))
        'Me.ddLanguage.SelectedIndex = 0
    End Sub

    Protected Sub GetCurrentHomePageContent()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
        Dim sql As String = "SELECT TOP 1 ID, htmlPageContent, RecordUpdatedBy, DateLastUpdated FROM tbl_Web_Page_Content WHERE RecordStatus = '1' AND PageName = @PageName AND LanguageCode= @LanguageCode ORDER BY DateLastUpdated DESC"
        Dim sqlCmd As New SqlCommand(sql, sqlCnn)
        Dim strPage As String = ddPageName.SelectedValue
        Dim strLangCode As String = ddLanguage.SelectedValue

        If CommonUtilsv2.Validate(strPage, CommonUtilsv2.DataTypes.String, True, True, False, 50) = False Then
            Exit Sub
        End If
        If CommonUtilsv2.Validate(strLangCode, CommonUtilsv2.DataTypes.String, True, True, False, 50) = False Then
            Exit Sub
        End If

        sqlCmd.Parameters.AddWithValue("@PageName", strPage)
        sqlCmd.Parameters.AddWithValue("@LanguageCode", strLangCode)
        Dim htmlContent As String = ""
        Dim strContentID As String = ""
        Dim myReader3 As SqlDataReader = Nothing
        Try
            sqlCnn.Open()
            myReader3 = sqlCmd.ExecuteReader()
            While myReader3.Read()
                '---------- Current Content ID ----------
                If myReader3.GetValue(0) Is System.DBNull.Value Then
                    strContentID = ""
                Else
                    strContentID = myReader3.GetInt32(0).ToString
                    lblContentID.Text = strContentID
                    Session("HomePageManager_SelectedContentID") = strContentID
                End If
                '---------- Current Page Content ----------
                If myReader3.GetValue(1) Is System.DBNull.Value Then
                    htmlContent = ""
                Else
                    htmlContent = myReader3.GetString(1)
                End If
                '---------- Last Updated By ----------
                If myReader3.GetValue(2) Is System.DBNull.Value Then
                    lblLastUpdatedBy.Text = ""
                Else
                    lblLastUpdatedBy.Text = myReader3.GetString(2)
                End If
                '---------- Last Updated Date ----------
                If myReader3.GetValue(3) Is System.DBNull.Value Then
                    lblLastUpdatedDate.Text = ""
                Else
                    lblLastUpdatedDate.Text = myReader3.GetDateTime(3).ToString
                End If
            End While
            editor.Content = htmlContent
        Catch ex As Exception
            errString = ex.Message
            errLocation = "Get Current Home Page Content"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myReader3 Is Nothing And Not myReader3.IsClosed Then
                myReader3.Close()
            End If
            sqlCmd.Dispose()
            sqlCmd = Nothing
            sqlCnn.Close()
        End Try
    End Sub

    Protected Sub GetHomePageContent(ByVal strContentID As String)
        If Not CommonUtilsv2.Validate(strContentID, CommonUtilsv2.DataTypes.Int, True, True, True) Then
            Exit Sub
        End If
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
        Dim sql As String = "SELECT ID, htmlPageContent, RecordUpdatedBy, DateLastUpdated, Pagetitle FROM tbl_Web_Page_Content WHERE RecordStatus = 1 AND ID = @ID ORDER BY DateLastUpdated DESC"
        Dim sqlCmd As New SqlCommand(sql, sqlCnn)
        sqlCmd.Parameters.AddWithValue("@ID", strContentID)
        Dim htmlContent As String = ""
        Dim myReader3 As SqlDataReader = Nothing
        Try
            sqlCnn.Open()
            myReader3 = sqlCmd.ExecuteReader()
            While myReader3.Read()
                '---------- Current Content ID ----------
                If myReader3.GetValue(0) Is System.DBNull.Value Then
                    strContentID = ""
                Else
                    strContentID = myReader3.GetInt32(0).ToString
                    lblContentID.Text = strContentID
                    Session("HomePageManager_SelectedContentID") = strContentID
                End If
                '---------- Current Page Content ----------
                If myReader3.GetValue(1) Is System.DBNull.Value Then
                    htmlContent = ""
                Else
                    htmlContent = myReader3.GetString(1)
                End If
                '---------- Last Updated By ----------
                If myReader3.GetValue(2) Is System.DBNull.Value Then
                    lblLastUpdatedBy.Text = ""
                Else
                    lblLastUpdatedBy.Text = myReader3.GetString(2)
                End If
                '---------- Last Updated Date ----------
                If myReader3.GetValue(3) Is System.DBNull.Value Then
                    lblLastUpdatedDate.Text = ""
                Else
                    lblLastUpdatedDate.Text = myReader3.GetDateTime(3).ToString
                End If
                '----------PageTitle ----------
                If myReader3.GetValue(4) Is System.DBNull.Value Then
                    tbPagetitle.Text = ""
                Else
                    tbPagetitle.Text = myReader3.GetString(4)
                End If
            End While
            editor.Content = htmlContent
        Catch ex As Exception
            errString = ex.Message
            errLocation = "Get Current Home Page Content"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            'LogError(errString, errLocation)
        Finally
            If Not myReader3 Is Nothing And Not myReader3.IsClosed Then
                myReader3.Close()
            End If
            sqlCmd.Dispose()
            sqlCmd = Nothing
            sqlCnn.Close()
        End Try
    End Sub

    Private Sub Editor_SubmitClicked(ByVal sender As Object, ByVal e As EventArgs)
        Dim strPage As String = ddPageName.SelectedValue
        Dim strLangCode As String = ddLanguage.SelectedValue
        Dim strPageTitle As String = tbPagetitle.Text

        If CommonUtilsv2.Validate(strPage, CommonUtilsv2.DataTypes.String, True, True, False, 50) = False Then
            Exit Sub
        End If
        If CommonUtilsv2.Validate(strLangCode, CommonUtilsv2.DataTypes.String, True, True, False, 50) = False Then
            Exit Sub
        End If
        If CommonUtilsv2.Validate(strPageTitle, CommonUtilsv2.DataTypes.String, True, True, False, 100) = False Then
            Exit Sub
        End If

        Dim htmlUpdatedContent As String = editor.Content
        Dim rightNow As DateTime = DateTime.Now
        Dim strDateTime As String
        strDateTime = rightNow.ToString("yyyyMMdd_hhmmss")
        Dim strFullDateTime As String
        strFullDateTime = rightNow.ToString("MMMM dd, yyyy @ hh:mm:ss")
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim strNow As Date = Now().ToString
        Dim strParQuery As String
        strParQuery = "INSERT INTO tbl_Web_Page_Content (htmlPageContent, RecordUpdatedBy, PageName, PageTitle, LanguageCode)"
        strParQuery = strParQuery + " VALUES (@htmlUpdatedContent, @RecordUpdatedBy, @PageName, @Pagetitle, @LanguageCode) "
        Dim myComm As New SqlCommand(strParQuery, myConn)
        Try
            myConn.Open()
            myComm.Parameters.AddWithValue("@htmlUpdatedContent", htmlUpdatedContent)
            myComm.Parameters.AddWithValue("@RecordUpdatedBy", Session("User"))
            myComm.Parameters.AddWithValue("@PageName", strPage)
            myComm.Parameters.AddWithValue("@PageTitle", strPageTitle)
            myComm.Parameters.AddWithValue("@LanguageCode", strLangCode)
            myComm.ExecuteNonQuery()
        Catch ex As Exception
            errString = ex.Message
            errLocation = "Insert Updated Home Page Content to DB"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            'LogError(errString, errLocation)
        Finally
            lbUploadDocument.Text = "Home Page Content updated successfully."
            lbUploadDocument.ForeColor = Drawing.Color.Green
            lbUploadDocument.Visible = True
            GetCurrentHomePageContent()
            LoadDocumentList()
            myComm.Dispose()
            myConn.Close()
        End Try
    End Sub

    Private Sub ddEditDocumet_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddEditDocument.SelectedIndexChanged
        Dim strContentID As String = ddEditDocument.SelectedValue.ToString
        GetHomePageContent(strContentID)
    End Sub

    Private Sub ddLanguage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddLanguage.SelectedIndexChanged
        LoadDocumentList()
        editor.Content = String.Empty
    End Sub

    Protected Sub ddPageName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'LoadLanguageList()
        editor.Content = String.Empty
    End Sub
</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
    <div align="center">
       <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0" class="wrapper">
            <%--<tr>
                 <td colspan="3" align="left" class="body_title" >
                    Manage Dynamic Page Content
                 </td>           
            </tr>--%>
            <%--<tr><td>&nbsp;</td></tr>--%>
            <tr>
                <td align="center" width="20%" colspan="1">
                    <asp:DropDownList ID="ddPageName" CssClass="form_input_required" Width="95%" runat="server" AutoPostBack="true" onselectedindexchanged="ddPageName_SelectedIndexChanged" TabIndex="3"></asp:DropDownList>
                </td>
                <td align="center" width="15%" colspan="1">
                    <asp:DropDownList ID="ddLanguage" CssClass="form_input_required" Width="95%" runat="server" AutoPostBack="true"  TabIndex="4"></asp:DropDownList>
                </td>
                <td align="center" width="65%" colspan="1">
                    <asp:DropDownList ID="ddEditDocument" CssClass="form_input_required" Width="95%" runat="server" AutoPostBack="true"  TabIndex="6"></asp:DropDownList>
                </td>
            </tr>
            <tr class="body_small">
                <td align="center" colspan="3">
                    <table id="Table2" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td align="right" width="5%"># :</td>
                            <td align="left" width="5%"><b><asp:Label ID="lblContentID" runat="server" Visible="True"></asp:Label></b></td>
                            <td align="right" width="15%">Last Updated By :</td>
                            <td align="left" width="25%"><b><asp:Label ID="lblLastUpdatedBy" runat="server" Visible="True"></asp:Label></b></td>
                            <td align="right" width="15%">Last Updated On :</td>
                            <td align="left" width="25%"><b><asp:Label ID="lblLastUpdatedDate" runat="server" Visible="True"></asp:Label></b></td>
                        </tr>
                    </table>    
                </td>
            </tr>
            <tr><td colspan="3"><hr /></td></tr>
            <tr>
                <td align="right" colspan="1">
                    Page Title :
                </td>
                <td align="center" colspan="2">
                    <asp:TextBox ID="tbPagetitle" enabled="false" CssClass="form_input_required" runat="server" MaxLength="50"  Width="95%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td colspan="3"><obout:Editor id="editor" runat="server" Appearance="full" 
                        ModeSwitch="true" Cancel="true" PreviewMode="true" FullHTML="true" 
                        ShowQuickFormat="true" height="400" QuickFormatFolder="~/includes" QuickFormatFile="homepage.css"  
                        StyleFile="" StyleFolder="">
                        </obout:Editor> </td>
            </tr>
            <tr>
                <td colspan="3" align="left">
                    <asp:Button ID="btnSubmit" runat="server" Text="Submit" OnClick="Editor_SubmitClicked" CssClass="Submit_button" />
                </td>
            </tr>
            <tr>
                <td align="left" colspan="3"><asp:Label ID="lbUploadDocument" runat="server" Visible="False"></asp:Label></td>
            </tr>
       </table>
    </div>
</asp:Content>