<%@ Page Language="VB" MasterPageFile="PageMaster.master" AutoEventWireup="false" EnableEventValidation="false" viewStateEncryptionMode="Auto" ASPCOMPAT="TRUE" %>
<%@ MasterType VirtualPath="PageMaster.master" %>
<%@ Register TagPrefix="obout" Namespace="Obout.Grid" Assembly="obout_Grid_NET" %>
<%@ Register TagPrefix="obout" Namespace="OboutInc.Calendar2" Assembly="obout_Calendar2_NET" %> 
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
        Session("PreviousPage") = "SiteManager.aspx"
        Dim strLogout As String = "False"
        strLogout = Request.QueryString.Get("logout")
        If strLogout = "True" Then
            Session("Level") = "0"
        End If

        If CustomRoles.RolesForPageLoad() Then
            If Not IsPostBack Then
                If rdAddDocument.Checked = True Then
                    fsAdd.Visible = True
                    fsEdit.Visible = False
                Else
                    rdAddDocument.Checked = False
                    rdEditDocument.Checked = True
                    fsAdd.Visible = False
                    fsEdit.Visible = True
                End If
                'LoadDomainLookup()
                LoadDocumentList()
            Else
                If rdAddDocument.Checked = False Then
                    Session("PublicDocument_DocumentManager_SelectedFileID") = ddEditDocument.SelectedValue
                End If
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

    Private Sub LoadDocumentList()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strSQLDropdown As String = "SELECT [ID], [Title] as [Report] FROM " & strImageTableName & " WHERE 1 =1 ORDER BY [ID] ASC" '
        LoadDropDownBox(dbKey, strSQLDropdown, Me.ddEditDocument, "ID", "Report")
        'insert the first item. Use the following to insert a listItem with Item.value = strFirstItemValue and Item.value = strFirstItemText
        ' dd.Items.Insert(0, New ListItem(strFirstItemText, strFirstItemValue))
        ddEditDocument.Items.Insert(0, New ListItem("Select a Document", 0))
        Try
            If Session("PublicDocument_DocumentManager_SelectedFileID") Is Nothing Or Session("PublicDocument_DocumentManager_SelectedFileID") = "" Then
            Else
                ddEditDocument.SelectedValue = Session("PublicDocument_DocumentManager_SelectedFileID")
            End If
        Catch ex As Exception
            Session("PublicDocument_DocumentManager_SelectedFileID") = 0
            ddEditDocument.SelectedIndex = Session("PublicDocument_DocumentManager_SelectedFileID")
        End Try
        LoadDocument()
    End Sub

    'Dim strDomainLookupType As String = "FAQCategory"
    Dim strImageTableName As String = "[tbl_Web_Public_Documents]"
    Dim strCommandBase As String = "SELECT [ID],FileName, Title, Contents, isActive, DateUploaded, Document_Size_In_MB as [ReportSize], EffectiveDate, ExpirationDate FROM " & strImageTableName
    Dim strWhereClause As String = " WHERE 1=1  " ' AND [isActive] =1 AND now() between [EffectiveDate] and [ExpirationDate] 
    Dim strCommandOrderBy As String = " ORDER BY [ID] ASC"

    Private Sub LoadDocument()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strSQLLoadDocument As String = strCommandBase & strImageTableName & strWhereClause
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim myComm As New SqlCommand()
        Dim searchStr As String = Session("PublicDocument_DocumentManager_SelectedFileID")
        Dim strDocumentDownloadID As String = ""
        Dim searchId As Integer
        If searchStr Is Nothing Or searchStr = "" Then
            ResetEditField()
            Exit Sub
        Else
            searchId = Convert.ToInt32(searchStr)
        End If
        Dim myReader As SqlDataReader = Nothing
        Try
            myComm.Connection = myConn
            myComm.CommandText = strSQLLoadDocument
            myComm.CommandText = myComm.CommandText & " AND [ID] =@Id "
            myComm.Parameters.AddWithValue("@Id", searchId)
            myComm.CommandText = myComm.CommandText & strCommandOrderBy
            myConn.Open()
            myReader = myComm.ExecuteReader()
            If myReader.HasRows Then
                While myReader.Read()
                    strDocumentDownloadID = myReader.GetInt32(0).ToString
                    tbEditDescription.Text = myReader.GetString(2)
                    tbEditComments.Text = myReader.GetString(3)
                    cbIsAvailiable.Checked = myReader.GetBoolean(4)
                    'lblDocumentSize.Text = myReader.GetString(7)
                    lblFileName.Text = myReader.GetString(1) + "&nbsp;&nbsp;(" + myReader.GetString(6) + ")"
                    If IsDBNull(myReader.GetValue(7)) Then
                        tbEditEffectiveDate.Text = ""
                    Else
                        tbEditEffectiveDate.Text = myReader.GetDateTime(7).ToString("MM/dd/yyyy")
                    End If
                    If IsDBNull(myReader.GetValue(8)) Then
                        tbEditExpirationDate.Text = ""
                    Else
                        tbEditExpirationDate.Text = myReader.GetDateTime(8).ToString("MM/dd/yyyy")
                    End If
                End While
            Else
                ResetEditField()
            End If
        Catch ex As Exception
            errString = ex.Message
            errLocation = "LoadDocument()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            'LogError(errString, errLocation)
        Finally
            If Not myReader Is Nothing And Not myReader.IsClosed Then
                myReader.Close()
            End If
            myComm.Dispose()
            myConn.Close()
            myConn.Dispose()
            myConn = Nothing
        End Try
        'tbDownloadLink.Text = "Download-PublicDocument.aspx?Document=" & strDocumentDownloadID
        '<a href="documents/1a-sample.pdf" target="_self" title="Tool Tip goes here">here</a>
        'tbDownloadLink.Text = "<a href='documents/1a-sample.pdf' target='_blank' title='Tool Tip goes here'>here</a>"
        tbDownloadLink.Text = "Download-PublicDocument.aspx?Document=" & strDocumentDownloadID
        'tbDownloadLink.Text = "<div title='" + tbEditComments.Text + "'><a href='Download-PublicDocument.aspx?Document=" + strDocumentDownloadID + "'>" + tbEditDescription.Text + "</a>"
    End Sub

    Protected Sub btnUploadDocument_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUploadDocument.Click
        lbUploadDocument.Text = ""
        lbUploadDocument.Visible = False
        Dim intId As Integer = 0
        Dim strReferenceID As String = ""
        Dim strDescription As String = StripCRLF(tbDescription.Text)
        ''Dim strInfoType As String = ddType.SelectedValue.ToString
        Dim strComments As String = tbComments.Text
        Dim strDocumentSize As String = ""
        Dim blIsActive As Boolean = cbUploadIsAvailiable.Checked
        Dim strEffectiveDate As String = tbEffectiveDate.Text
        Dim strExpirationDate As String = tbExpirationDate.Text
        Dim strUser As String = Session("User")
        Dim insSuccess As Boolean = False
        'Upload Image to RESOURCES
        If Not FileUpload1.PostedFile Is Nothing AndAlso FileUpload1.PostedFile.ContentLength > 0 Then
            Dim sqlCmd2 As New SqlClient.SqlCommand
            Dim FileData(FileUpload1.PostedFile.ContentLength - 1) As Byte
            Dim Md5 As New System.Security.Cryptography.MD5CryptoServiceProvider
            Dim SqlCnn2 As New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced"))
            Dim sqlDa2 As New SqlClient.SqlDataAdapter
            Dim sqlDs2 As New DataSet
            FileUpload1.PostedFile.InputStream.Read(FileData, 0, FileUpload1.PostedFile.ContentLength)
            Dim rSize As Single = FileUpload1.PostedFile.ContentLength
            rSize = rSize / (1024 * 1024.0)
            strDocumentSize = FormatNumber(rSize, 2, TriState.True, TriState.False, TriState.True) & " MB"
            Dim ByteHash() As Byte = Md5.ComputeHash(FileData)
            sqlCmd2.CommandTimeout = 0
            sqlCmd2.Connection = SqlCnn2
            sqlCmd2.CommandText = "prc_uploadPublicDocument"
            sqlCmd2.CommandType = CommandType.StoredProcedure
            'Catch the info
            Dim Filename As String = FileUpload1.PostedFile.FileName.Substring(FileUpload1.PostedFile.FileName.LastIndexOf("\") + 1)
            Filename = Replace(Filename, " ", "_")
            Filename = Replace(Filename, "'", "")
            'strReferenceID = ddReference.SelectedValue
            'Input Data
            sqlCmd2.Parameters.AddWithValue("@ID", intId)
            sqlCmd2.Parameters.AddWithValue("@FileName", Filename)
            sqlCmd2.Parameters.AddWithValue("@ImageBLOB", FileData)
            sqlCmd2.Parameters.AddWithValue("@Title", strDescription)
            sqlCmd2.Parameters.AddWithValue("@Contents", strComments)
            sqlCmd2.Parameters.AddWithValue("@RecordUpdateBy", Session("User"))
            sqlCmd2.Parameters.AddWithValue("@IsActive", blIsActive)
            sqlCmd2.Parameters.AddWithValue("@Document_Size_In_MB", strDocumentSize)
            If strEffectiveDate <> "" Then
                sqlCmd2.Parameters.AddWithValue("@EffectiveDate", strEffectiveDate)
            End If
            If strExpirationDate <> "" Then
                sqlCmd2.Parameters.AddWithValue("@ExpirationDate", strExpirationDate)
            End If
            'Run the query
            If lbUploadDocument.Text = "" Then
                Try
                    SqlCnn2.Open()
                    sqlDa2.SelectCommand = sqlCmd2
                    sqlDa2.Fill(sqlDs2)
                    If sqlDs2.Tables(0).Rows.Count > 0 Then
                        insSuccess = True
                    End If
                Catch ex As SqlClient.SqlException
                    errString = ex.Message
                    errLocation = "btnUpdateDocument_Click"
                    CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
                    'LogError(errString, errLocation)
                    lbUploadDocument.Text += ex.Message.ToString
                Finally
                    sqlCmd2.Dispose()
                    sqlCmd2 = Nothing
                    SqlCnn2.Close()
                    SqlCnn2.Dispose()
                    SqlCnn2 = Nothing
                End Try
            End If
        End If
        If insSuccess = True Then
            LoadDocumentList()
            lbUploadDocument.Text += "Document was uploaded successfully. "
            lbUploadDocument.Visible = True
            lbUploadDocument.ForeColor = Drawing.Color.Green
        Else
            'lbUploadDocument.Text += "Cannot find document. "
            lbUploadDocument.Visible = True
            lbUploadDocument.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Protected Sub btnEditDocument_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEditDocument.Click
        lbEditDocument.Text = ""
        Dim intId As Integer

        'If CommonUtilsv2.Validate(tbEditDescription.Text, CommonUtilsv2.DataTypes.String, False, True, False, 4000) = False OrElse CommonUtilsv2.Validate(tbEditComments.Text, CommonUtilsv2.DataTypes.String, False, True, False, 4000) = False OrElse CommonUtilsv2.Validate(tbEditEffectiveDate.Text, CommonUtilsv2.DataTypes.String, False, True, False, 20) = False OrElse CommonUtilsv2.Validate(tbEditExpirationDate.Text, CommonUtilsv2.DataTypes.String, False, True, False, 20) = False Then
        '    Exit Sub
        'End If

        Dim strComments As String = StripCRLF(tbEditComments.Text)
        Dim strDescription As String = StripCRLF(tbEditDescription.Text)
        ''Dim strType As String = ddEditType.SelectedValue.ToString
        Dim bIsActive As Boolean = cbIsAvailiable.Checked
        Dim strEffectiveDate As String = tbEditEffectiveDate.Text
        Dim strExpirationDate As String = tbEditExpirationDate.Text
        Dim strUser As String = Session("User")
        Dim strDocumentSize As String = ""
        Dim editSuccess As Boolean = False
        Dim FileData() As Byte = New Byte() {}
        Dim Filename As String = ""
        'Upload Image to RESOURCES
        Dim SqlCnn2 As New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced"))
        Dim sqlCmd2 As New SqlClient.SqlCommand
        Dim sqlDa2 As New SqlClient.SqlDataAdapter
        Dim sqlDs2 As New DataSet
        Try
            If Not FileUpload2.PostedFile Is Nothing AndAlso FileUpload2.PostedFile.ContentLength > 0 Then
                Filename = FileUpload2.PostedFile.FileName.Substring(FileUpload2.PostedFile.FileName.LastIndexOf("\") + 1)
                ReDim FileData(FileUpload2.PostedFile.ContentLength - 1)
                FileUpload2.PostedFile.InputStream.Read(FileData, 0, FileUpload2.PostedFile.ContentLength)
                Dim rSize As Single = FileUpload2.PostedFile.ContentLength
                rSize = rSize / (1024 * 1024.0)
                strDocumentSize = FormatNumber(rSize, 2, TriState.True, TriState.False, TriState.True) & " MB"
                Dim Md5 As New System.Security.Cryptography.MD5CryptoServiceProvider
                Dim ByteHash() As Byte = Md5.ComputeHash(FileData)
                'Else
                '    '   strDocumentSize = lblDocumentSize.Text
            End If
            Filename = Replace(Filename, " ", "_")
            Filename = Replace(Filename, "'", "")
            intId = Convert.ToInt32(ddEditDocument.SelectedValue.ToString)
            sqlCmd2.CommandTimeout = 0
            sqlCmd2.Connection = SqlCnn2
            sqlCmd2.CommandText = "prc_uploadPublicDocument"
            sqlCmd2.CommandType = CommandType.StoredProcedure
            'strReferenceID = ddReference.SelectedValue
            'Input Data
            sqlCmd2.Parameters.AddWithValue("@ID", intId)
            sqlCmd2.Parameters.AddWithValue("@FileName", Filename)
            sqlCmd2.Parameters.AddWithValue("@ImageBLOB", FileData)
            sqlCmd2.Parameters.AddWithValue("@Title", strDescription)
            sqlCmd2.Parameters.AddWithValue("@Contents", strComments)
            If strDocumentSize <> "" Then
                sqlCmd2.Parameters.AddWithValue("@Document_Size_In_MB", strDocumentSize)
            End If
            sqlCmd2.Parameters.AddWithValue("@IsActive", bIsActive)
            sqlCmd2.Parameters.AddWithValue("@RecordUpdateBy", strUser)
            If strEffectiveDate <> "" Then
                sqlCmd2.Parameters.AddWithValue("@EffectiveDate", strEffectiveDate)
            End If
            If strExpirationDate <> "" Then
                sqlCmd2.Parameters.AddWithValue("@ExpirationDate", strExpirationDate)
            End If
            'Run the query
            SqlCnn2.Open()
            sqlDa2.SelectCommand = sqlCmd2
            sqlDa2.Fill(sqlDs2)
            If sqlDs2.Tables(0).Rows.Count > 0 Then
                editSuccess = True
            End If
        Catch ex As Exception
            errString = ex.Message
            errLocation = "btnEditDocument_Click"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            'LogError(errString, errLocation)
            lbEditDocument.Text += ex.Message.ToString
        Finally
            SqlCnn2.Close()
            SqlCnn2.Dispose()
            SqlCnn2 = Nothing
            sqlCmd2.Dispose()
            sqlCmd2 = Nothing
        End Try
        If editSuccess = True Then
            LoadDocumentList()
            lbEditDocument.Text += "Document was updated successfully.  "
            lbEditDocument.Visible = True
            lbEditDocument.ForeColor = Drawing.Color.Green
        Else
            lbEditDocument.Visible = True
            lbEditDocument.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Function StripCRLF(ByVal HTMLToStrip As String) As String
        Dim stripped As String
        If HTMLToStrip <> "" Then
            stripped = HTMLToStrip.Replace(vbCr, " ").Replace(vbLf, " ")
            Return stripped
        Else
            Return ""
        End If
    End Function

    Private Sub ddEditDocumet_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddEditDocument.SelectedIndexChanged
        Session("PublicDocument_DocumentManager_SelectedFileID") = ddEditDocument.SelectedValue.ToString
        LoadDocument()
    End Sub

    Private Sub ActionType_Checked_Button(ByVal Src As Object, ByVal Args As EventArgs)
        If (Src.Text = "Add") Then
            If Src.checked = True Then
                fsAdd.Visible = True
                fsEdit.Visible = False
                ResetEditField()
            Else
                fsAdd.Visible = False
                fsEdit.Visible = True
                LoadDocumentList()
            End If
        Else
            If Src.checked = True Then
                fsAdd.Visible = False
                fsEdit.Visible = True
                LoadDocumentList()
            Else
                fsAdd.Visible = True
                fsEdit.Visible = False
                ResetEditField()
            End If
        End If
    End Sub


    Private Sub ResetEditField()
        ddEditDocument.SelectedIndex = 0
        tbEditDescription.Text = ""
        tbEditComments.Text = ""
        cbIsAvailiable.Checked = True
        lbEditDocument.Text = ""
        lbUploadDocument.Text = ""
        tbDescription.Text = ""
        tbComments.Text = ""
        cbUploadIsAvailiable.Checked = True
        lblFileName.Text = ""
        tbEffectiveDate.Text = ""
        tbExpirationDate.Text = ""
        tbEditEffectiveDate.Text = ""
        tbEditExpirationDate.Text = ""
    End Sub
</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
    <div align="center">
        <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0" class="wrapper">
         <%-- <tr>
             <td colspan="5" align="left" class="body_title">
                Manage Public Documents
             </td>           
          </tr>--%>
          <tr><td>&nbsp;</td></tr>
          <tr>
                <td align="right" width="10%" class="form_input"><b>Select Action:</b>&nbsp;</td>
                <td align="left" width="90%" rowspan="1" class="body_doclink">
                    <asp:RadioButton ID="rdAddDocument" runat="Server" Text="Add" GroupName="ActionRadioButtons" AutoPostBack="True" Checked="True" OnCheckedChanged="ActionType_Checked_Button" />
                    <asp:RadioButton ID="rdEditDocument" runat="Server" Text="Edit" GroupName="ActionRadioButtons" AutoPostBack="True" Checked="False" OnCheckedChanged="ActionType_Checked_Button" />
                </td>
            </tr>
            <tr><td>&nbsp;</td></tr>
            <tr>
                <td align="center" valign="top" colspan="2" rowspan="2" class="body_doclink" width="100%">
                    <%--Add/Upload an Existing Document Section--%>
                    <fieldset runat="server" id="fsAdd" style="width: 98%">
                        <legend class="body_text" style="border-style: none">&nbsp;&nbsp;<b>Upload New Document</b>&nbsp;&nbsp;</legend>
                        <table class="body_doclink" width="100%">
                            <tr>
                                <td align="right" width="20%">Link Display Title :</td>
                                <td align="left" width="60%">
                                    <asp:TextBox ID="tbDescription" CssClass="form_input_required" runat="server" MaxLength="80" Width="98%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="Requiredfieldvalidator1" runat="server" ErrorMessage="Title is a required value."
                                    ControlToValidate="tbDescription" Display="Dynamic" ValidationGroup="uploadDocument"></asp:RequiredFieldValidator>                                
                                </td>
                                <td align="left" width="20%" class="body_small">&nbsp;&nbsp;(Max 80 characters)</td>
                            </tr>
                            <tr>
                                <td align="right">Link Hover Content :</td>  <%--.{0,x}--%>
                                <td align="left">
                                    <asp:TextBox ID="tbComments" CssClass="form_input_required" runat="server" Height="80px" MaxLength="4000" TextMode="MultiLine" Width="98%" ></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="Requiredfieldvalidator4" runat="server" ErrorMessage="FAQ Contents is a required value."
                                    ControlToValidate="tbComments" Display="Dynamic" ValidationGroup="uploadDocument"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator 
                                        ID="regexValidatorFAQContents" runat="server" 
                                        ControlToValidate="tbComments"
                                        ValidationExpression="^.{1,4000}$"
                                        ErrorMessage="Maximum FAQ Content is 4,000 characters.">
                                    </asp:RegularExpressionValidator>
                                </td>
                                <td align="left" class="body_small">&nbsp;&nbsp;(Max 4,000 characters)</td>
                            </tr>
                            <tr>
                                <td align="right">Document File Name :</td>
                                <td align="left">
                                    <asp:FileUpload ID="FileUpload1" CssClass="form_input_required" runat="server"  />
                                    <asp:RequiredFieldValidator ID="Requiredfieldvalidator3" runat="server" ErrorMessage="Document File name is a required value."
                                        ControlToValidate="FileUpload1" Display="Dynamic" ValidationGroup="uploadDocument"></asp:RequiredFieldValidator>                                
                                </td>
                            </tr>
                            <tr><td colspan="3"><hr /></td></tr>
                            <tr>
                                <td align="right">Display on web site :</td>
                                <td align="left">
                                    <asp:CheckBox ID="cbUploadIsAvailiable" TextAlign="right" runat="server" Checked="true" AutoPostBack="false" />
                                </td>
                            </tr>
                             <tr>
                                <td align="right">Do Not Display Before Date :</td>
                                <td align="left">
                                    <asp:TextBox ID="tbEffectiveDate" CssClass="body_text_boxes" runat="server" MaxLength="10"></asp:TextBox>&nbsp;
                                    <obout:Calendar ID="EffectiveDate" runat="server" DatePickerMode="true" TextBoxId="tbEffectiveDate" EnableViewState="false" StyleFolder="styles/calendar/blue" DatePickerImagePath="styles/calendar/icon2.gif"></obout:Calendar>
                                    <asp:CompareValidator id="CompareValidator10" runat="server" ErrorMessage="Invalid date." ControlToValidate="tbEditEffectiveDate" Type="Date" Operator="DataTypeCheck" ValidationGroup="uploadDocument" />
                                    <asp:CompareValidator ID="CompareValidator11" runat="server" ErrorMessage="Do Not Publish Date must be earlier than Do Not Publish After Date."
                                        ControlToValidate="tbEffectiveDate" Type="Date" ControlToCompare="tbExpirationDate" Operator="LessThan" Display="Dynamic"
                                        ValidationGroup="uploadDocument"></asp:CompareValidator>                               
                                </td>
                                <td align="left" class="body_small">&nbsp;&nbsp;(Leave blank if this date is not used)</td>
                            </tr>
                            <tr>
                                <td align="right">Do Not Display After Date :</td>
                                <td align="left">
                                    <asp:TextBox ID="tbExpirationDate" CssClass="body_text_boxes" runat="server" MaxLength="10"></asp:TextBox>&nbsp;
                                    <obout:Calendar ID="ExpirationDate" runat="server" DatePickerMode="true" TextBoxId="tbExpirationDate" EnableViewState="false" StyleFolder="styles/calendar/blue" DatePickerImagePath="styles/calendar/icon2.gif"></obout:Calendar>
                                    <asp:CompareValidator id="CompareValidator8" runat="server" ErrorMessage="Invalid date." ControlToValidate="tbEditExpirationDate" Type="Date" Operator="DataTypeCheck" ValidationGroup="uploadDocument"/>
                                    <asp:CompareValidator ID="CompareValidator9" runat="server" ErrorMessage="Do Not Publish After Date must be later than Do Not Publish Date."
                                        ControlToValidate="tbExpirationDate" Type="Date" ControlToCompare="tbEffectiveDate" Operator="GreaterThan" Display="Dynamic"
                                        ValidationGroup="uploadDocument"></asp:CompareValidator>                               
                                </td>
                                <td align="left" class="body_small">&nbsp;&nbsp;(Leave blank if this date is not used)</td>
                            </tr>                         
                            
                            <tr>
                                <td align="right">&nbsp;</td>
                                <td align="left">
                                    <asp:Button ID="btnUploadDocument" runat="server" CssClass="body_text_boxes" Text="Upload Document" CausesValidation="True" ValidationGroup="uploadDocument" />
                                    <asp:Label ID="lbUploadDocument" runat="server" Visible="False"></asp:Label></td>
                            </tr>
                        </table>
                    </fieldset>
                    <%--Edit an Existing Document Section--%>
                    <fieldset runat="server" id="fsEdit" style="width: 98%">
                        <legend class="body_text" style="border-style: none">&nbsp;&nbsp;<b>Edit Existing Document</b>&nbsp;&nbsp;</legend>
                        <table class="body_doclink" width="100%">
                            <tr>
                                <td align="right">Select a Document:</td>
                                <td align="left" colspan="1">
                                    <asp:DropDownList ID="ddEditDocument" CssClass="form_input_required" Width="100%" runat="server" AutoPostBack="true"></asp:DropDownList>
                                    <asp:RequiredFieldValidator ID="Requiredfieldvalidator7" runat="server" ErrorMessage="A document must be selected."
                                        ControlToValidate="ddEditDocument" Display="Dynamic" ValidationGroup="editDocument"></asp:RequiredFieldValidator>
                                    <asp:CompareValidator ID="CompareValidator3" runat="server" ErrorMessage="A document must be selected."
                                        ControlToValidate="ddEditDocument" Type="Integer" ValueToCompare="0" Operator="NotEqual" Display="Dynamic"
                                        ValidationGroup="editDocument"></asp:CompareValidator>                                    
                               </td>
                            </tr>
                            <tr>
                                <td align="right" width="20%">Link Display Title :</td>
                                <td align="left" width="60%">
                                    <asp:TextBox ID="tbEditDescription" CssClass="form_input_required" runat="server" MaxLength="80" Width="98%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="Requiredfieldvalidator6" runat="server" ErrorMessage="Title is a required value."
                                    ControlToValidate="tbEditDescription" Display="Dynamic" ValidationGroup="editDocument"></asp:RequiredFieldValidator>                                
                                </td>
                                <td align="left" width="20%" class="body_small">&nbsp;&nbsp;(Max 80 characters)</td>
                            </tr>
                            <tr>
                                <td align="right">Link Hover Contents :</td>
                                <td align="left">
                                    <asp:TextBox ID="tbEditComments" CssClass="form_input_required" runat="server" Height="80px" MaxLength="4000" TextMode="MultiLine" Width="98%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="Requiredfieldvalidator40" runat="server" ErrorMessage="FAQ Contents is a required value."
                                    ControlToValidate="tbComments" Display="Dynamic" ValidationGroup="uploadDocument"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator 
                                        ID="RegularExpressionValidator1" runat="server" 
                                        ControlToValidate="tbComments"
                                        ValidationExpression="^.{1,4000}$"
                                        ErrorMessage="Maximum FAQ Content is 4,000 characters.">
                                    </asp:RegularExpressionValidator>
                                </td>
                                <td align="left" class="body_small">&nbsp;&nbsp;(Max 4,000 characters)</td>
                            </tr>
                            <tr>
                                <td align="right">Current File / Size :</td>
                                <td align="left" class="body_simple">
                                    <asp:Label ID="lblFileName" runat="server" CssClass="body_text_boxes" />
                                </td>
                            </tr>
                            <tr>
                                <td align="right">Replace with File :</td>
                                <td align="left">
                                    <asp:FileUpload ID="FileUpload2" runat="server" CssClass="body_text_boxes" />
                                </td>
                            </tr>
                            <tr><td colspan="3"><hr /></td></tr>
                            <tr>
                                <td align="right">Instructions :</td>
                                <td align="left" colspan="2" class="body_small">
                                    1)  In the Dynamic Page editor, select the page you would like to edit, and add the text that will contain the link.
                                </td>
                            </tr>
                            <tr>
                                <td align="right">&nbsp;</td>
                                <td align="left" colspan="2" class="body_small">
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;For example, type the word 'here' as in the this sentence: Click <u>here</u> for the document.                                  
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="20%">&nbsp;</td>
                                <td align="left" colspan="2" class="body_small">
                                    2) Select the word or phrase to hold the link, and click the HTML Link icon : 
                                    <img src="Images/EditorLink.png" alt="Editor HTML Link icon" />
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="20%">&nbsp;</td>
                                <td align="left" colspan="2" class="body_small">
                                    <img src="Images/EditorValues.png" alt="Editor HTML Link icon" />
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="20%">&nbsp;</td>
                                <td align="left" colspan="2" class="body_small">
                                    3) In the Create URL Link box, paste the green link below into the URL field, and set the Target to 'New Window' if desired.
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="20%">Download Link :</td>
                                <td align="left" width="60%">
                                    <asp:TextBox ID="tbDownloadLink" CssClass="form_input_special" runat="server" MaxLength="80" Width="98%"></asp:TextBox>
                                </td>
                                <td align="left" width="20%" class="body_small">&nbsp;&nbsp;(Copy/Paste this into Dynamic Page)</td>
                            </tr>
                            <tr><td colspan="3"><hr /></td></tr>
                            <tr>
                                <td align="right">Display on web site:</td>
                                <td align="left">
                                    <asp:CheckBox ID="cbIsAvailiable" TextAlign="right" runat="server" Checked="true" AutoPostBack="false" />
                                </td>
                            </tr>
                             <tr>
                                <td align="right">Do Not Display Before Date:</td>
                                <td align="left">
                                    <asp:TextBox ID="tbEditEffectiveDate" CssClass="body_text_boxes" runat="server" MaxLength="10"></asp:TextBox>&nbsp;
                                    <obout:Calendar ID="EditEffectiveDate" runat="server" DatePickerMode="true" TextBoxId="tbEditEffectiveDate" EnableViewState="false" StyleFolder="styles/calendar/blue" DatePickerImagePath="styles/calendar/icon2.gif"></obout:Calendar>
                                    <asp:CompareValidator id="CompareValidator5" runat="server" ErrorMessage="Invalid date." ControlToValidate="tbEditEffectiveDate" Type="Date" Operator="DataTypeCheck" ValidationGroup="editDocument"/>
                                    <asp:CompareValidator ID="CompareValidator4" runat="server" ErrorMessage="Do Not Publish Date must be earlier than Do Not Publish After Date."
                                        ControlToValidate="tbEditEffectiveDate" Type="Date" ControlToCompare="tbEditExpirationDate" Operator="LessThan" Display="Dynamic"
                                        ValidationGroup="editDocument"></asp:CompareValidator>                               
                                </td>
                                <td align="left" class="body_small">&nbsp;&nbsp;(Leave blank if this date is not used)</td>
                            </tr>
                            <tr>
                                <td align="right">Do Not Display After Date:</td>
                                <td align="left">
                                    <asp:TextBox ID="tbEditExpirationDate" CssClass="body_text_boxes" runat="server" MaxLength="10"></asp:TextBox>&nbsp;
                                    <obout:Calendar ID="EditExpirationDate" runat="server" DatePickerMode="true" TextBoxId="tbEditExpirationDate" EnableViewState="false" StyleFolder="styles/calendar/blue" DatePickerImagePath="styles/calendar/icon2.gif"></obout:Calendar>
                                    <asp:CompareValidator id="CompareValidator6" runat="server" ErrorMessage="Invalid date." ControlToValidate="tbEditExpirationDate" Type="Date" Operator="DataTypeCheck" ValidationGroup="editDocument"/>
                                    <asp:CompareValidator ID="CompareValidator7" runat="server" ErrorMessage="Do Not Publish After Date must be later than Do Not Publish Date."
                                        ControlToValidate="tbEditExpirationDate" Type="Date" ControlToCompare="tbEditEffectiveDate" Operator="GreaterThan" Display="Dynamic"
                                        ValidationGroup="editDocument"></asp:CompareValidator>                               
                                </td>
                                <td align="left" class="body_small">&nbsp;&nbsp;(Leave blank if this date is not used)</td>                           
                            </tr>                                
                                                         
                            <tr>
                                <td align="right">&nbsp;</td>
                                <td align="left">
                                    <asp:Button ID="btnEditDocument" runat="server" CssClass="body_text_boxes" Text="Update Document" CausesValidation="True" ValidationGroup="editDocument" />
                                    <asp:Label ID="lbEditDocument" runat="server" Visible="False"></asp:Label>
                                </td>
                            </tr>
                        </table>
                    </fieldset>
                </td>
            </tr> 
        </table>
    </div>
</asp:Content>