<%@ Page Language="VB" MasterPageFile="PageMaster.master" AutoEventWireup="false" EnableEventValidation="false" ViewStateEncryptionMode="Never" AspCompat="TRUE" %>

<%@ MasterType VirtualPath="PageMaster.master" %>
<%@ Register TagPrefix="obout" Namespace="Obout.Grid" Assembly="obout_Grid_NET" %>
<%@ Register Assembly="Obout.Ajax.UI" Namespace="Obout.Ajax.UI.HTMLEditor" TagPrefix="obout" %>
<%@ Register TagPrefix="obout" Namespace="Obout.Interface" Assembly="obout_Interface" %>
<%@ Register TagPrefix="obout" Namespace="OboutInc.Calendar2" Assembly="obout_Calendar2_NET" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.Page" %>
<%@ Import Namespace="Webapps.Utils" %>
<%@ Register Assembly="PdfViewer" Namespace="PdfViewer" TagPrefix="cc1" %>
<%@ Import Namespace="System.Web.Services" %>
<%@ Import Namespace="Newtonsoft.Json" %>
<%@ Import Namespace="Newtonsoft.Json.Linq" %>

<script language="VB" runat="server">
    Dim SessionVariable_Prefix As String = "DocumentDetail_"
    '************* Error logging Section ***********************
    Dim pageName As String
    Dim errLocation As String
    Dim errString As String
    Dim strImageFileID As String
    Dim strDocOwner As String
    Public StrListIndes As String
    Dim strfilename As String
    Public StrdocumentCopy As String
    Public IsAuthority As Boolean = False 'Editable: true,not Editable: false 
    Public IsNotCompleteAuthority As Boolean = True 'Editable: true,not Editable: false
    Public FileExt As String = ".doc,.docx,.xls,.ppt,.pptx,.mdb,.mdbx,.pdf,.tiff,.tif,.csv,.txt,.png"
    Public GroupID As Integer = 0
    Public IsReplaceDCNImage As Boolean = False
    Private CanBeDeleted As Boolean = False
    Public IsVisibleInDelete As Boolean = False

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
        FileExt = Application("AllowedUploadFiletypes").Replace(";", ",")


        pageName = Request.RawUrl.ToString
        tableuser.Value = Session("User")
        nodatadiv.Visible = False
        Try
            Master.SetCurrentMenuItem = System.IO.Path.GetFileName(Request.RawUrl.ToString)
        Catch ex As Exception
            Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
            Exit Sub
        End Try

        If Not IsPostBack Then
            Response.Cache.SetExpires(DateTime.Parse(DateTime.Now.ToString()))
            Response.Cache.SetCacheability(HttpCacheability.Private)
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            Response.Cache.SetNoStore()
            Response.AppendHeader("Pragma", "no-cache")
            Response.AppendHeader("cache-control", "private")
            Response.CacheControl = "no-cache"
        End If

        If CustomRoles.RolesForPageLoad() Then
            Dim strID As String = ""
            strID = Request.QueryString.Get("id")
            Session(SessionVariable_Prefix & "DCN") = strID

            If Not IsPostBack Then
                If Not CommonUtilsv2.Validate(strID, CommonUtilsv2.DataTypes.Int, True, True, True) Then
                    Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
                    Exit Sub
                End If

                'If IsExietSId(strID) = False Then
                '    documentdetail.Visible = False
                '    nodatadiv.Visible = True
                'End If

                If CommonUtilsv2.HasAccesstoDCN(Session("User"), strID) Then
                    LoadDropDownLists()
                    LoadRecordDetails(strID)
                    LoadDocument(strID, 1)
                    LoadAttachmentList()
                Else  ' No access or not exist
                    documentdetail.Visible = False
                    nodatadiv.Visible = True
                End If
                If GroupID = 1 Then

                    If CustomRoles.IsInRole("R_Delete_DCN") AndAlso CanBeDeleted = True Then
                        btnDeleteDCN.Visible = True
                        btnMoveDCNForDelete.Visible = False
                    Else
                        btnDeleteDCN.Visible = False
                        If IsVisibleInDelete = True Then
                            tbDeleteNotes.ReadOnly = 1
                        End If
                    End If

                    If CanBeDeleted = False AndAlso IsNotCompleteAuthority = True Then
                        btnDeleteDCN.Visible = False
                        btnMoveDCNForDelete.Visible = True
                    Else
                        btnMoveDCNForDelete.Visible = False
                    End If
                Else
                    btnDeleteDCN.Visible = False
                    btnMoveDCNForDelete.Visible = False
                    tbDeleteNotes.ReadOnly = 1

                End If
            End If

        Else
            CustomRoles.TransferIfNotInRole(True)
            Exit Sub
        End If
    End Sub

    Private Sub LoadDropDownLists()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim params As SqlParameter() = Nothing
        Dim strSQL As String = " SELECT LookupID As OwnerID, LookupDesc As OwnerName FROM tbl_Web_Lookup  WHERE Lookuptype = @LookupType Order by SortOrder, LookupDesc "
        params = {New SqlParameter("@LookupType", "DocOwner")}
        CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddDocUser, "OwnerID", "OwnerName", params)
        ddDocUser.Items.Insert(0, New ListItem("Select A Value ", 0))

        ''Doc Type
        strSQL = " select ID as DocTypeID,DocTypeName AS DocType from tbl_Web_DocTypes  WHERE RecordStatus=1 " +
        "AND GroupID=(select GroupID from tbl_Web_Queues where ID=(select QueueID from tbl_Web_DocumentAttributes where DCN=@DCN)) order by DocTypeName "
        Dim params3 As SqlParameter() = {New SqlParameter("@DCN", Request.QueryString.Get("id"))}

        CommonUtilsv2.LoadDropDownBox(dbKey, strSQL, ddDocType, "DocTypeID", "DocType", params3)
        ddDocType.Items.Insert(0, New ListItem("Select A Value ", 0))
        ddDocType.Enabled = False

    End Sub

    Private Sub LoadDocument(ByVal strID As String, ByVal strSourceID As String)
        Dim strEmbed As String = "<iframe id='myframe' width='99.5%' height='500px' class='iframeplain' frameborder='0' vspace='0' hspace='0' marginwidth='0' marginheight='0' scrolling='Yes' src='{0} #view=Fith'></iframe>"
        DocumentViewer.Text = String.Format(strEmbed, String.Format(ResolveUrl("~/PDFHandler.ashx?Id={0}&SId={1}&ImageArchived={2}"), strID, strSourceID, ImageArchived.Value))
    End Sub

    Protected Sub LoadRecordDetails(ByVal strID As String)

        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim sqlCnn As New SqlClient.SqlConnection(dbKey)

        Dim strSQL As String = " SELECT * FROM [v_w_DocumentDetail] WHERE DCN = @ID "

        Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
        sqlCmd.Parameters.AddWithValue("@ID", strID)
        Dim myReader As SqlDataReader = Nothing
        Try
            sqlCnn.Open()
            myReader = sqlCmd.ExecuteReader()
            While myReader.Read()
                'ImageArchived  zkf 2012107-9 update archive database  start
                Dim StrIArchived As String = "0"
                If Not IsDBNull(myReader(myReader.GetOrdinal("ImageArchived"))) Then
                    StrIArchived = StrHelp.GetInt(myReader.GetValue(myReader.GetOrdinal("ImageArchived"))).ToString()
                End If
                ImageArchived.Value = StrIArchived
                'end

                lblDCN.Text = strID
                lblDownloadDCN.Text = "<a href='download-Document.aspx?id=" & strID & "&IArchived=" & StrIArchived & "' title='download'><img src='/Images/download16.png' style='width:12px;margin-left:10px;border: solid 1px darkgrey;'></a>"
                If Not IsDBNull(myReader(myReader.GetOrdinal("FileName"))) Then
                    lblFileName.Text = myReader.GetValue(myReader.GetOrdinal("FileName"))
                    strfilename = myReader.GetValue(myReader.GetOrdinal("FileName"))
                End If
                If Not IsDBNull(myReader(myReader.GetOrdinal("QueueName"))) Then
                    lblQueueName.Text = myReader.GetValue(myReader.GetOrdinal("QueueName"))
                End If
                If Not IsDBNull(myReader(myReader.GetOrdinal("UploadedDate"))) Then
                    ' lblUploadedDate.Text = (myReader.GetValue(myReader.GetOrdinal("UploadedDate"))).ToString("MM/dd/yyyy")
                    lblUploadedDate.Text = myReader.GetDateTime(myReader.GetOrdinal("UploadedDate")).ToString("MM/dd/yyyy")
                End If
                If Not IsDBNull(myReader(myReader.GetOrdinal("DocStatusName"))) Then
                    lblDocStatus.Text = myReader.GetValue(myReader.GetOrdinal("DocStatusName"))
                End If
                'If Not IsDBNull(myReader(myReader.GetOrdinal("DocTypeName"))) Then
                '    lblDocType.Text = myReader.GetValue(myReader.GetOrdinal("DocTypeName"))
                'End If

                If Not IsDBNull(myReader(myReader.GetOrdinal("DocTypeID"))) Then
                    'strDocTypeId = myReader.GetValue(myReader.GetOrdinal("DocTypeID"))
                    Dim iDocTypeId As Integer = StrHelp.GetInt(myReader.GetValue(myReader.GetOrdinal("DocTypeID")))
                    ddDocType.SelectedValue = iDocTypeId
                    'Else
                    '    strDocOwner = ""
                End If


                If Not IsDBNull(myReader(myReader.GetOrdinal("Owner"))) Then
                    strDocOwner = myReader.GetValue(myReader.GetOrdinal("Owner"))
                    Dim iOwnerID As Integer = StrHelp.GetInt(myReader.GetValue(myReader.GetOrdinal("OwnerID")))
                    ddDocUser.SelectedValue = iOwnerID
                    '  ddDocUser.Items.FindByText(strDocOwner, ).Selected = True
                Else
                    strDocOwner = ""
                End If
                GroupID = StrHelp.GetInt(myReader.GetValue(myReader.GetOrdinal("GroupID")))
                Dim panQueuId = StrHelp.GetInt(myReader.GetValue(myReader.GetOrdinal("QueueID")))
                If GroupID = 1 Then
                    StrdocumentCopy = "<a href='/DocumnetCopy.aspx?id=" + strID + "' style='float: left;margin-left: 20px; height: 40px; background: #1c8450;line-height: 40px;border-radius: 10px;padding: 0px 5px;color: #fff;'><span>Document Copy</span></a>"
                End If
                If (myReader.GetValue(myReader.GetOrdinal("QueueName")).Trim().ToUpper().Contains("DELETE")) Then 'Whether the queueid contains DELETE, false means  
                    IsVisibleInDelete = True ' yes
                End If
                If (GroupID = 1 And panQueuId <> 6) Then
                    ddDocType.Enabled = True
                End If



                ''Authority handling 
                If (myReader.GetValue(myReader.GetOrdinal("QueueName")).Trim().ToUpper().Contains("COMPLETE")) Then 'Whether the queueid contains Complete, false means it cannot be edited 
                    IsNotCompleteAuthority = False 'Not editable 
                End If

                If String.IsNullOrEmpty(strDocOwner) Or String.Compare("Not Set", strDocOwner, True) = 0 Or String.Compare("API", strDocOwner, True) = 0 Then 'Is it empty 
                    Dim strGroupIdString As Integer = StrHelp.GetInt(myReader.GetValue(myReader.GetOrdinal("GroupID")))
                    If CustomRoles.IsInRole("R_WG_" & strGroupIdString.ToString()) Then 'Is in the groupword
                        IsAuthority = True
                    End If
                Else 'Is it empty
                    If String.Compare(Session("User"), strDocOwner, True) = 0 Then 'Is it oneself 
                        IsAuthority = True
                    End If
                End If

                If CustomRoles.IsInRole("R_Replace_DCN") Then
                    IsReplaceDCNImage = True
                End If
                If (myReader.GetValue(myReader.GetOrdinal("QueueName")).Trim().ToUpper().Contains("DELETE")) Then '
                    CanBeDeleted = True 'Not editable 
                Else
                    CanBeDeleted = False
                End If
                If Not IsDBNull(myReader(myReader.GetOrdinal("DeleteNotes"))) Then
                    tbDeleteNotes.Text = myReader.GetValue(myReader.GetOrdinal("DeleteNotes"))
                End If




            End While
            LoadDCNLog(strID)
        Catch ex As Exception
            errString = ex.Message
            errLocation = "Load record details"
            'CommonUtilsv2.CreateErrorLog(ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
            sqlCmd.Dispose()
            sqlCmd = Nothing
            sqlCnn.Close()
        End Try
    End Sub

    Protected Sub LoadDCNLog(ByVal strID As String)

        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim params As SqlParameter() = Nothing

        Dim strSQL As String = "select CONVERT(VARCHAR(10), CreatedDate, 101) + ' '+CONVERT(varchar(100),CreatedDate, 8) as createtime,[UserId],isNull(Comments,'') + char(13) as Comments from tbl_Web_DCNLog where DCN=@DCN   order by CreatedDate desc "
        params = {New SqlParameter("@DCN", strID)}
        Dim ds As System.Data.DataSet
        Dim ReturnStr As String = String.Empty
        Try
            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)

            If Not ds Is Nothing Then
                Dim fileName As String = String.Empty
                Dim fileType As String = String.Empty
                ReturnStr = "<table border='0' cellspacing='0' cellpadding='0'>"
                For Each row As System.Data.DataRow In ds.Tables(0).Rows
                    ReturnStr += "<tr>"
                    ReturnStr += "<td width='170px' align='center'>" & row("createtime") & "</td>"
                    ReturnStr += "<td width='100px' align='center'>" & row("UserId") & "</td>"
                    ReturnStr += "<td  width='380px'>" & row("Comments").Replace(vbCr, "").Replace(vbLf, "<br/>").Replace(vbCrLf, "") & "</td>"
                    ReturnStr += "</tr>"
                Next
                ReturnStr += "</table>"
            End If

            lbxAccumulatedComment.InnerHtml = ReturnStr
        Catch ex As Exception

        End Try

    End Sub

    Private Function checkRolePermission(ByVal strRoleID As String) As Boolean
        Dim ret As Boolean = False
        'the not-misc-roles version of misc roles
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim myReader As SqlDataReader = Nothing
        Dim mySQL As String = "Select Role_ID from tbl_ROLES_UserRoles Where User_ID = @User "
        Dim params As SqlParameter() = {
        New SqlParameter("@User", Session("User"))
            }
        Try
            myReader = CommonUtilsv2.GetDataReader(dbKey, mySQL, CommandType.Text, params)
            While myReader.Read
                If String.Compare(myReader(0), strRoleID, True) = 0 Then
                    ret = True
                    Exit While
                Else
                    ret = False
                End If
            End While
        Catch ex As Exception
            Throw ex
        Finally
            If Not myReader Is Nothing Then
                myReader.Close()
            End If
        End Try

        Return ret
    End Function

    Function StripHTMLTags(ByVal HTMLToStrip As String) As String
        Dim stripped As String
        If HTMLToStrip <> "" Then
            stripped = Regex.Replace(HTMLToStrip, "<(.|\n)+?>", String.Empty)
            Return stripped
        Else
            Return ""
        End If
    End Function

    Protected Sub SiteMapPath1_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SiteMapNodeItemEventArgs)
        If e.Item.ItemType = SiteMapNodeItemType.Root OrElse (e.Item.ItemType = SiteMapNodeItemType.PathSeparator AndAlso e.Item.ItemIndex = 1) Then
            e.Item.Visible = False
        End If
    End Sub

    Protected Sub UploadAttachment(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim intActionDocId As Integer = -1
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

            If bValidateFile = False Then
                sbValidateMessage.AppendLine(" File was not uploaded. ")
                lbSubmitComment.ForeColor = Drawing.Color.Red
                lbSubmitComment.Font.Bold = True
                lbSubmitComment.Text = sbValidateMessage.ToString
                lbSubmitComment.Visible = True
                Exit Sub
            Else
                lbSubmitComment.ForeColor = Drawing.Color.Green
                lbSubmitComment.Visible = False
                'Page.ClientScript.RegisterStartupScript(Page.GetType(), "myscript", "clearFileUploader()", False)
                'Page.ClientScript.RegisterStartupScript&#65288;Page.GetType(), "myscript", "alert&#65288;'hi'&#65289;;", True&#65289;
                'ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), "myscript", "clearFileUploader()", True)
                Page.ClientScript.RegisterStartupScript(Page.GetType(), "view", "clearFileUploader();", True)
            End If
            If Not Filename Is Nothing And Filename <> "" Then
                intActionDocId = StoreDocument()
                If intActionDocId = -1 Then
                Else
                    LoadAttachmentList()
                    LoadRecordDetails(Session(SessionVariable_Prefix & "DCN"))
                    LoadDropDownLists()
                    LoadDocument(Session(SessionVariable_Prefix & "DCN"), 1)
                End If
            End If
        Catch ex As Exception
            errLocation = "UploadAttachment"
            CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try
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
            sqlCmd2.CommandText = "prc_uploadAttachment"
            sqlCmd2.CommandType = CommandType.StoredProcedure
            Dim Filename As String = FileUpload1.PostedFile.FileName.Substring(FileUpload1.PostedFile.FileName.LastIndexOf("\") + 1)
            Filename = Replace(Filename, " ", "_")
            Filename = Replace(Filename, "'", "")
            sqlCmd2.Parameters.AddWithValue("@DCN", Session(SessionVariable_Prefix & "DCN"))
            sqlCmd2.Parameters.AddWithValue("@FileName", Filename)
            sqlCmd2.Parameters.AddWithValue("@ImageBlob", ImageData)
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
                    insSuccess = True
                End If
            Catch ex As SqlClient.SqlException
                errString = ex.Message
                errLocation = "StoreDocument() "
                ' CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
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

    Sub LoadAttachmentList()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim params As SqlParameter() = Nothing

        Dim strSQL As String = "SELECT [ID],[FileName],UploadedDate,isnull(ImageArchived,0) ImageArchived FROM [dbo].[tbl_Web_Attachments] Where Recordstatus=1 and DCN=@DCN"
        '  Dim strSQL As String = "SELECT [ID],[FileName],UploadedDate,ImageBlob,ImageArchived FROM [dbo].[tbl_Web_Attachments] Where Recordstatus=1 and DCN=@DCN"
        params = {New SqlParameter("@DCN", Session(SessionVariable_Prefix & "DCN"))}
        'Dim myReader As SqlDataReader = Nothing
        Dim ds As System.Data.DataSet
        Dim ReturnStr As String = String.Empty
        Try
            'myReader = CommonUtilsv2.GetDataReader(dbKey, strSQL, CommandType.Text, params)
            'lbxAttachments.Items.Clear()
            'lbxAttachments.DataSource = myReader
            'lbxAttachments.DataValueField = "ID"
            'lbxAttachments.DataTextField = "FileName"
            'lbxAttachments.DataBind()
            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)

            If Not ds Is Nothing And ds.Tables(0).Rows.Count > 0 Then
                Dim fileName As String = String.Empty
                Dim fileType As String = String.Empty
                ReturnStr = "<table>"
                For Each row As System.Data.DataRow In ds.Tables(0).Rows
                    fileType = String.Empty
                    fileName = row("FileName")
                    fileType = fileName.Substring(fileName.LastIndexOf(".") + 1)

                    ReturnStr += "<tr  StrId=""" & row("ID") & """>"

                    If ((String.Compare("pdf", fileType, True) = 0) OrElse (String.Compare("tif", fileType, True) = 0) OrElse (String.Compare("tiff", fileType, True) = 0) OrElse (String.Compare("jpg", fileType, True) = 0) OrElse (String.Compare("jpeg", fileType, True) = 0) OrElse (String.Compare("png", fileType, True) = 0)) Then
                        ReturnStr += "<td style='width:280px;'><div style='width: 280px;word-break: keep-all;white-space: nowrap;overflow: hidden;text-overflow: ellipsis;'><a href='javascript:void(0);' onclick='selectattac(this," & row("ImageArchived") & ")' style='color:#006176; text-decoration: underline;' title='" & row("FileName") & "'>" & row("FileName") & "</a></div></td>"
                    Else

                        ReturnStr += "<td style='width:280px;'><div style='width: 280px;word-break: keep-all;white-space: nowrap;overflow: hidden;text-overflow: ellipsis;'><span title='" & row("FileName") & "'>" & row("FileName") & "</span></div></td>"

                    End If

                    ReturnStr += "<td style='width: 100px;'>" & Convert.ToDateTime(row("UploadedDate")).ToString("MM/dd/yyyy") & "</td><td style='width:70px;padding: 5px;'><a href='Download-Attachment.aspx?id=" & row("ID") & "&IArchived=" & row("ImageArchived") & "' title='download'><img src='/Images/download16.png' style='width:12px;' style='margin-left:10px;'/></a>"

                    If IsNotCompleteAuthority And IsAuthority And GroupID <> 2 Then
                        'If IsAuthority Then

                        ReturnStr += "<a href='javascript:void(0);' onclick='colsetattac(this)' style='color:red;margin-left:10px;' title='delete'><img src='/Images/delete16.png' style='width:12px;'/></a></td>"
                    Else
                        ReturnStr += "</td>"
                    End If
                    ReturnStr += "</tr>"

                Next
                ReturnStr += "</table>"
            End If

            lbxAttachments.InnerHtml = ReturnStr
        Catch ex As Exception

            'Finally
            '    If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
            '        myReader.Close()
            '    End If
        End Try
    End Sub

    Private Sub DisableEditing(ByVal ctrlContainer As Control)

        For Each ctrl As Control In ctrlContainer.Controls
            If TypeOf ctrl Is TextBox Then
                CType(ctrl, TextBox).Enabled = False
            ElseIf TypeOf ctrl Is CheckBox Then
                CType(ctrl, CheckBox).Enabled = False
            ElseIf TypeOf ctrl Is FileUpload Then
                CType(ctrl, FileUpload).Enabled = False
            ElseIf TypeOf ctrl Is Button Then
                CType(ctrl, Button).Enabled = False
            ElseIf TypeOf ctrl Is DropDownList Then
                CType(ctrl, DropDownList).Enabled = False
            Else
                If ctrl.Controls.Count > 0 Then
                    DisableEditing(ctrl)
                End If
            End If
        Next

    End Sub

    Private Sub EnableEditing(ByVal ctrlContainer As Control)
        'ElseIf TypeOf ctrl Is CheckBox Then
        '    CType(ctrl, CheckBox).Enabled = True
        For Each ctrl As Control In ctrlContainer.Controls
            If TypeOf ctrl Is TextBox Then
                CType(ctrl, TextBox).Enabled = True
            ElseIf TypeOf ctrl Is FileUpload Then
                CType(ctrl, FileUpload).Enabled = True
            ElseIf TypeOf ctrl Is Button Then
                CType(ctrl, Button).Enabled = True
            ElseIf TypeOf ctrl Is DropDownList Then
                CType(ctrl, DropDownList).Enabled = True
            Else
                If ctrl.Controls.Count > 0 Then
                    EnableEditing(ctrl)
                End If
            End If
        Next
    End Sub

    Private Sub LoadDocumentIndexes(ByVal DocmID As Integer)
        Dim ReturnStr As String = String.Empty
        Dim ds As DataSet
        Dim params As SqlParameter() = Nothing

        Dim strSQL As String = "  select IndexName,DisplayName,DateType from tbl_Web_SearchIndexNames as webindex where webindex.GroupID=(select GroupID from tbl_Web_Queues as webque where webque.ID=(select QueueID from tbl_Web_DocumentAttributes where DCN=@DCN) ) Order by SortOrder  "
        params = {New SqlParameter("@DCN", DocmID)}
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")


        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)
            If Not ds Is Nothing Then
                Dim strIndexes As String = String.Empty
                Dim StrLen As Integer = ds.Tables(0).Rows.Count
                strIndexes = "<tr>"
                For Each row As System.Data.DataRow In ds.Tables(0).Rows

                    strIndexes += "<td width='15%' align='right' class='form_label_medium Reassignment' rel='" & row("IndexName") & "'>" & row("IndexName") & ":&nbsp;</td><td width='18%' align='left' class='form_input_small'><asp:TextBox ID='tb" & row("IndexName") & "' runat='server' Visible='True' MaxLength='200' Height='25px' Width='170px' TabIndex='10'/><obout:Calendar ID='Calendar" & row("IndexName") & "'  runat='server' DateFormat='MM/dd/yyyy' DatePickerMode='True' TextBoxId='tb" & row("IndexName") & "' Columns='1' EnableViewState='False'StyleFolder='styles/calendar/simple' DatePickerImagePath='styles/calendar/icon2.gif'></obout:Calendar></td>"

                    If (StrLen Mod 3 = 0) Then
                        strIndexes += "</tr><tr>"
                    End If
                Next
                strIndexes += "</tr>"
                StrListIndes = strIndexes
            End If

        Catch ex As Exception



        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try
    End Sub

    Public Shared Function GetDCNLog(ByVal DocmID As Integer) As String

        Dim strRet As String = String.Empty
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim params As SqlParameter() = Nothing

        Dim strSQL As String = "select CONVERT(VARCHAR(10), CreatedDate, 101) + ' '+CONVERT(varchar(100),CreatedDate, 8) as createtime,[UserId],isNull(Comments,'') + char(13) as Comments from tbl_Web_DCNLog where DCN=@DCN  order by CreatedDate desc"
        params = {New SqlParameter("@DCN", DocmID)}
        Dim ds As System.Data.DataSet
        Dim ReturnStr As String = String.Empty
        Try
            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)

            If Not ds Is Nothing Then
                Dim fileName As String = String.Empty
                Dim fileType As String = String.Empty
                ReturnStr = "<table border='0' cellspacing='0' cellpadding='0'>"
                For Each row As System.Data.DataRow In ds.Tables(0).Rows
                    ReturnStr += "<tr>"
                    ReturnStr += "<td width='170px' align='center'>" & row("createtime") & "</td>"
                    ReturnStr += "<td width='100px' align='center'>" & row("UserId") & "</td>"
                    ReturnStr += "<td width='380px'>" & row("Comments").Replace("\", "\\").Replace(Chr(13) + Chr(10), "<br>").Replace(Chr(13), "<br>").Replace(Chr(10), "<br>").Replace("'", "&quot&").Replace("""", "!quot!").Replace(",", "%quot%").Replace(":", "-quot-") & "</td>"
                    ReturnStr += "</tr>"
                Next
                ReturnStr += "</table>"
            End If

            strRet = ReturnStr
        Catch ex As Exception
            CommonUtilsv2.CreateErrorLog("GetDCNLog()", ex, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress())
            ' CommonUtilsv2.CreateErrorLog(ex, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress())
        End Try
        Return strRet

    End Function

    Private Shared Function IsExietSId(ByVal StrId As Integer) As Boolean
        Dim isExit As Integer = 0
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim params As SqlParameter() = Nothing
        Dim strSQL As String = "select count(*) from [dbo].[tbl_Web_DocumentAttributes]  Where Recordstatus=1 and DCN=@DCN"
        params = {New SqlParameter("@DCN", StrId)}

        isExit = StrHelp.GetInt(CommonUtilsv2.RunScalarQuery(dbKey, strSQL, CommandType.Text, params))
        Return IIf(isExit > 0, True, False)
    End Function

    Protected Sub ReplaceImage(ByVal sender As Object, ByVal e As EventArgs)
        Dim intDocId As Integer = -1
        Dim Filename As String = ""
        Dim lFSizeLimit As Long = 0
        Dim bValidateFile As Boolean = True
        Dim sbValidateMessage As StringBuilder = New StringBuilder()

        If FileUpload2.HasFile = True Then
            Filename = FileUpload2.PostedFile.FileName.Substring(FileUpload2.PostedFile.FileName.LastIndexOf("\") + 1)
            Dim strFileExt As String = FileUpload2.PostedFile.FileName.Substring(FileUpload2.PostedFile.FileName.LastIndexOf("."))
            If InStr(Application("AllowedUploadFiletypes"), strFileExt, CompareMethod.Text) = 0 Then
                sbValidateMessage.AppendLine(" The type " & strFileExt & " of the file to be uploaded is not supported. ")
                bValidateFile = False
            End If
            '      lFSizeLimit = Long.Parse(Application("AllowedUploadFileSize"))
            Long.TryParse(Application("AllowedUploadFileSize"), lFSizeLimit)

            If FileUpload2.PostedFile.ContentLength > lFSizeLimit Then
                sbValidateMessage.AppendLine(" Document size exceeds the" & Application("AllowedUploadFileSizeInMB") & " limit.")
                bValidateFile = False
            End If
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
            intDocId = UpdateImageFile()
            If intDocId = -1 Then
                lbSubmitComment.Text = "File has not been uploaded."
                lbSubmitComment.ForeColor = Drawing.Color.Red
                lbSubmitComment.Visible = True
            Else
                Response.Redirect("/Document-Detail.aspx?id=" + intDocId.ToString())
            End If
        Else
            lbSubmitComment.Text = "Please add upload image."
            lbSubmitComment.ForeColor = Drawing.Color.Red
            lbSubmitComment.Visible = True
        End If
    End Sub

    Private Function UpdateImageFile() As Integer

        Dim insSuccess As Boolean = False
        Dim intId As Integer = -1
        Dim strDocumentSize As String = ""
        Dim params As SqlParameter() = Nothing
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        If Not FileUpload2.PostedFile Is Nothing AndAlso FileUpload2.PostedFile.ContentLength > 0 Then

            Dim ImageData(FileUpload2.PostedFile.ContentLength - 1) As Byte
            Dim StrSql = "update  tbl_Web_Documents set ImageBlob=@ImageBlob,UploadedBy=@UploadedBy where DCN=@DCN"
            FileUpload2.PostedFile.InputStream.Read(ImageData, 0, FileUpload2.PostedFile.ContentLength)
            Dim rSize As Single = FileUpload2.PostedFile.ContentLength
            rSize = rSize / (1024 * 1024.0)
            strDocumentSize = FormatNumber(rSize, 2, TriState.True, TriState.False, TriState.True)

            Dim Filename As String = FileUpload2.PostedFile.FileName.Substring(FileUpload2.PostedFile.FileName.LastIndexOf("\") + 1)
            Filename = Replace(Filename, " ", "_")
            Filename = Replace(Filename, "'", "")

            params = {New SqlParameter("@ImageBlob", ImageData), New SqlParameter("@UploadedBy", Session("User")), New SqlParameter("@DCN", Session(SessionVariable_Prefix & "DCN"))}
            Try
                'Dim oldparams As SqlParameter() = Nothing
                'Dim StroldSql As String = "select * from tbl_Web_Documents where DCN=@DCN"
                'oldparams = {New SqlParameter("@DCN", Session(SessionVariable_Prefix & "DCN"))}
                'Dim oldfilename As String = CommonUtilsv2.RunScalarQuery(dbKey, StroldSql, CommandType.Text, oldparams)

                CommonUtilsv2.RunNonQuery(dbKey, StrSql, CommandType.Text, params)
                intId = Session(SessionVariable_Prefix & "DCN")
                CommonUtilsv2.CreateDCNLog(intId, CommonUtilsv2.LogType.User, " replace image,the Old image is :" & Filename, HttpContext.Current.Session("User").ToString(), HttpContext.Current.Request.UserHostAddress())
            Catch ex As SqlClient.SqlException
                '  errString = ex.Message
                errLocation = "UpdateImageFile() "
                CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            End Try
        End If

        Return intId
    End Function

    Protected Sub MoveDCNForDelete(ByVal sender As Object, ByVal e As EventArgs)

        Dim bSuccess As Boolean = True
        Dim iDestQueue As Integer = 19
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim sqlCmd As New SqlClient.SqlCommand

        sqlCmd.CommandTimeout = 0
        sqlCmd.Connection = myConn
        sqlCmd.CommandText = "prc_MoveDCNForDeletion"
        sqlCmd.CommandType = CommandType.StoredProcedure

        Try
            myConn.Open()
            sqlCmd.Parameters.Clear()
            sqlCmd.Parameters.AddWithValue("@DCN", Session(SessionVariable_Prefix & "DCN"))
            sqlCmd.Parameters.AddWithValue("@DestQId", iDestQueue)
            sqlCmd.Parameters.AddWithValue("@DeleteNotes", tbDeleteNotes.Text.Trim())
            sqlCmd.Parameters.AddWithValue("@UserID", Session("User"))

            sqlCmd.ExecuteNonQuery()
            LoadRecordDetails(Session(SessionVariable_Prefix & "DCN"))
            If CustomRoles.IsInRole("R_Delete_DCN") AndAlso CanBeDeleted = True Then
                btnDeleteDCN.Visible = True
                btnMoveDCNForDelete.Visible = False
            Else
                btnDeleteDCN.Visible = False
            End If

            If CanBeDeleted = False AndAlso IsNotCompleteAuthority = True Then
                btnDeleteDCN.Visible = False
                btnMoveDCNForDelete.Visible = True
            Else
                btnMoveDCNForDelete.Visible = False
            End If
        Catch ex As Exception
            bSuccess = False
            errString = ex.Message
            errLocation = "MoveDCNForDelete()"
            CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            myConn.Close()
            sqlCmd.Parameters.Clear()
        End Try
    End Sub

    Protected Sub DeleteDCN(ByVal sender As Object, ByVal e As EventArgs)

        Dim bSuccess As Boolean = True
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim sqlCmd As New SqlClient.SqlCommand

        sqlCmd.CommandTimeout = 0
        sqlCmd.Connection = myConn
        sqlCmd.CommandText = "prc_DeleteDCN"
        sqlCmd.CommandType = CommandType.StoredProcedure

        Try
            myConn.Open()
            sqlCmd.Parameters.Clear()
            sqlCmd.Parameters.AddWithValue("@DCN", Session(SessionVariable_Prefix & "DCN"))
            sqlCmd.Parameters.AddWithValue("@UserID", Session("User"))
            sqlCmd.ExecuteNonQuery()
            Response.Redirect("QueueRegister.aspx?qid=19", False)

        Catch ex As Exception
            bSuccess = False
            errString = ex.Message
            errLocation = "DeleteDCN()"
            CommonUtilsv2.CreateErrorLog(errLocation, ex, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            myConn.Close()
            sqlCmd.Parameters.Clear()
        End Try
    End Sub

    <WebMethod()>
    Public Shared Function GetIndexlist(ByVal DocmID As Integer) As String
        Dim StrUser As String = HttpContext.Current.Session("User")
        If StrUser Is Nothing Then
            Return "{""msg"":""You don't have access to this Function "",""status"":-1}"
        End If
        ' status 0:error,1:no data,2:success
        Dim ReturnStr As String = String.Empty
        Dim ds As DataSet
        Dim params As SqlParameter() = Nothing
        ',REPLACE(REPLACE(TT.IndexValue,char(13),''),char(10),'')  as
        Dim strSQL As String = "select IndexName,DisplayName,DateType,TT.IndexValue from tbl_Web_SearchIndexNames as webindex left join ( select ltrim(rtrim(IndexValue)) as IndexValue,[Index] from tbl_Web_DocumentAttributes as C    UNPIVOT(IndexValue for [Index] in (Index1,Index2,Index3,Index4,Index5,Index6,Index7,Index8,Index9,Index10)) AS T  where DCN=@DCN ) as TT on TT.[Index]=webindex.IndexName  where webindex.GroupID=(select GroupID from tbl_Web_Queues as webque where webque.ID=(select QueueID from tbl_Web_DocumentAttributes where DCN=@DCN) )  Order by SortOrder;select * from Tbl_Web_States;"
        params = {New SqlParameter("@DCN", DocmID)}
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")


        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)
            If Not ds Is Nothing Then
                Dim strjson As String = String.Empty
                Dim Indes7state As String = String.Empty 'Get the value of state to query county data  
                For Each row As System.Data.DataRow In ds.Tables(0).Rows
                    If (row("DateType").ToString().Trim().ToUpper() = "SELECT") Then
                        Dim Strchildjson As String = String.Empty
                        ''state
                        If (row("DisplayName").ToString().ToUpper().Contains("STATE")) Then
                            Indes7state = row("IndexValue").ToString().Trim()
                            For Each row1 As System.Data.DataRow In ds.Tables(1).Rows
                                Strchildjson += "{""Id"":""" + row1("state_name").Trim() + """,""Name"":""" + row1("state_name").Trim() + """},"
                            Next
                        End If
                        'County
                        If (row("DisplayName").ToString().ToUpper().Contains("COUNTY")) Then
                            If Not IsNothing(row("IndexValue").ToString()) Then
                                Dim dsCounty As DataSet = Nothing
                                Dim params2 As SqlParameter() = Nothing
                                strSQL = "select * from Tbl_Web_Counties where State=@State "
                                params2 = {New SqlParameter("@State", Indes7state)}
                                dsCounty = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params2)

                                If Not IsNothing(dsCounty) Then
                                    For Each row2 As System.Data.DataRow In dsCounty.Tables(0).Rows
                                        Strchildjson += "{""Id"":""" + row2("State") + """,""Name"":""" + row2("County").Trim() + """},"
                                    Next
                                End If
                            End If
                        End If

                        strjson += "{""Index"":""" + row("IndexName") + """,""Name"":""" + row("DisplayName") + """,""Type"":""" + row("DateType") + """,""Value"":""" + row("IndexValue") + """,""data"":" + "[" & Strchildjson.ToString().TrimEnd(",") & "]" + "},"
                    Else
                        Dim StrValue As String = GetFuReplace(row("IndexValue").ToString()).Trim() 'remove \r \n \t
                        ' Dim StrValue As String = row("IndexValue").ToString().Replace("'", "&quot&").Replace("""", "!quot!").Replace(",", "%quot%").Replace(":", "-quot-").Replace("\\", "#quot#").Trim()
                        strjson += "{""Index"":""" + row("IndexName").Trim() + """,""Name"":""" + row("DisplayName").Trim() + """,""Type"":""" + row("DateType").Trim() + """,""Value"":""" + StrValue + """},"
                    End If
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

    <WebMethod()>
    Public Shared Function DeleteAttac(ByVal ID As Integer) As String
        Dim StrUser As String = HttpContext.Current.Session("User")
        If StrUser Is Nothing Then
            Return "{""msg"":""You don't have access to this Function "",""status"":-1}"
        End If
        ' status 0:error,1:success
        Dim errLocation As String
        Dim errString As String
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim ReturnStr As String = String.Empty
        Dim strComments As String = String.Empty
        Dim params As SqlParameter() = Nothing

        ' Dim strSQL As String = "  UPDATE tbl_Web_Attachments SET Recordstatus=0 WHERE ID =" & ID
        Dim strSQL As String = "prc_DeleteAttachment"
        params = {New SqlParameter("@AtID", ID), New SqlParameter("@UserID", HttpContext.Current.Session("User").ToString())}

        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            CommonUtilsv2.RunNonQuery(dbKey, strSQL, CommandType.StoredProcedure, params)

            strComments = GetDCNLog(HttpContext.Current.Session("DocumentDetail_DCN").ToString().Replace("\", ""))

            ReturnStr = "{""msg"":""Update is successful."",""status"":1,""Comments"":""" & strComments.Trim() & """}"

        Catch ex As Exception
            errString = ex.Message
            errLocation = "DeleteAttac()"
            ' CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress())
            CommonUtilsv2.CreateErrorLog(errLocation, ex, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress())
            ReturnStr = "{""msg"":""Update failed. The detail error is : " & ex.ToString() + """,""status"":0}"
        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try
        Return ReturnStr
    End Function

    <WebMethod()>
    Public Shared Function GetAttachmentBase(ByVal ID As Integer, ByVal type As Integer) As String
        Dim StrUser As String = HttpContext.Current.Session("User")
        If StrUser Is Nothing Then
            Return "{""msg"":""You don't have access to this Function "",""status"":-1}"
        End If
        ' status 0:error,1:no data,2:success
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim ReturnStr As String = String.Empty
        Dim params As SqlParameter() = Nothing
        Dim ds As DataSet
        Dim strSQL As String = String.Empty
        If type = 1 Then
            strSQL = "select ImageBlob from tbl_Web_Documents where DCN=@DCN  "
            params = {New SqlParameter("@DCN", ID)}
        Else
            strSQL = "SELECT ImageBlob FROM tbl_Web_Attachments Where Recordstatus=1 and ID=@ID"
            params = {New SqlParameter("@ID", ID)}
        End If



        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)
            If Not ds Is Nothing Then
                Dim strjson As String = String.Empty

                ReturnStr = "{""msg"":""this success!"",""status"":2,""data"":""" & System.Convert.ToBase64String(ds.Tables(0).Rows(0)("ImageBlob")) + """}"

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

    <WebMethod()>
    Public Shared Function SubmintDetail(ByVal DocmId As Integer, ByVal Notes As String, ByVal Index As String, ByVal User As String) As String
        Dim StrUser As String = HttpContext.Current.Session("User")
        If StrUser Is Nothing Then
            Return "{""msg"":""You don't have access to this Function "",""status"":-1}"
        End If
        ' status 0:error,1:success
        Dim errLocation As String
        Dim sbValidateMessage As StringBuilder = New StringBuilder()
        Dim dtNow As DateTime = DateTime.UtcNow
        Dim ReturnStr As String = String.Empty
        Dim strNow As Date = Now().ToString()
        Dim strComments As String = String.Empty
        Dim StrLogSQL As String = String.Empty

        Dim StrIndexList As JObject = New JObject()
        Try
            StrIndexList = JObject.Parse(Index.Replace("\", "\\"))
        Catch ex As Exception
            StrIndexList = Nothing
        End Try
        If StrIndexList Is Nothing Then
            Return "{""msg"":""Please fill in the value"",""status"":0}"
        End If

        Dim strSQLQuery As String = "UPDATE tbl_Web_DocumentAttributes SET Comments=@Comments"
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        myConn = New SqlClient.SqlConnection(dbKey)
        Dim myComm As New SqlCommand(strSQLQuery, myConn)
        myConn.Open()
        Try
            'log  ---
            Dim paramsLog As SqlParameter() = Nothing
            Dim dsIndex As DataSet
            Dim StrInsertAttachment As StringBuilder = New StringBuilder()
            Dim strSQLAttachment As String = "select IndexName,ColumnName from tbl_Web_SearchIndexNames where  GroupID=(select GroupID from tbl_Web_Queues where ID=(select QueueID from tbl_Web_DocumentAttributes where DCN=@DCN))"
            paramsLog = {New SqlParameter("@DCN", DocmId)}
            dsIndex = CommonUtilsv2.GetDataSet(dbKey, strSQLAttachment, CommandType.Text, paramsLog)
            'If (dsIndex Is Nothing Or dsIndex.Tables(0).Rows.Count <= 0) Then
            '    ReturnStr = "{""msg"":""Document Indexes"",""status"":0}"
            'End If

            If StrIndexList IsNot Nothing Then
                Dim results As JToken = StrIndexList
                For Each item As JProperty In results
                    strSQLQuery += "," + item.Name + "=Replace(Replace(Replace(Replace(Replace(@" + item.Name + ", Char(13) + Char(10), ''), CHAR(13), ''), CHAR(10), ''),CHAR(8),''),CHAR(9),'') "
                    ' strSQLQuery += "," + item.Name + "=@" + item.Name
                    If Not String.IsNullOrEmpty(item.Value) Then

                        myComm.Parameters.AddWithValue("@" + item.Name.Trim(), item.Value.ToString().Replace("&quot&", "").Replace("!quot!", """").Replace("%quot%", ",").Replace("-quot-", ":").Replace("#quot#", "\\").Replace("-quot-", ":").Replace(vbCr, "").Replace(vbLf, "").Replace(vbCrLf, "").Replace(vbTab, "").Trim())
                        'myComm.Parameters.AddWithValue("@" + item.Name.Trim(), item.Value.ToString().Replace("&quot&", "'").Replace("!quot!", """").Replace("%quot%", ",").Replace("-quot-", ":").Replace("#quot#", "\\").Replace("-quot-", ":").Trim())

                    Else
                        myComm.Parameters.AddWithValue("@" + item.Name.Trim(), "")
                    End If
                    StrLogSQL += item.Name + "," 'log
                Next

                strSQLQuery += ", UpdatedBy =@UpdatedBy, UpdatedDate =@UpdatedDate "
                strSQLQuery += " WHERE dcn=@id"
                myComm.Parameters.AddWithValue("@UpdatedBy", User)
                myComm.Parameters.AddWithValue("@UpdatedDate", strNow)
                myComm.Parameters.AddWithValue("@id", DocmId)
                myComm.CommandText = strSQLQuery

                myComm.Parameters.AddWithValue("@Comments", System.DBNull.Value)

                ' Create log first
                Dim bLogCreated As Boolean = CommonUtilsv2.CreateAuditLog(DocmId, User)
                'Dim bLogCreated As Boolean = CommonUtilsv2.CreateAuditLog("tbl_Web_DocumentAttributes", "DCN", DocmId, "tbl_Web_DocumentAttributeslog", User)
                If bLogCreated = False Then
                    Throw New Exception("Create audit log record failed")
                End If

                ''Write log  starte
                If Not IsNothing(StrLogSQL) And Not String.IsNullOrEmpty(StrLogSQL) And StrIndexList.Count > 0 Then
                    StrLogSQL = "select " & StrLogSQL.TrimEnd(",") & " from tbl_Web_DocumentAttributes where DCN=@DCN"
                    Dim params As System.Data.SqlClient.SqlParameter() = {
                        New System.Data.SqlClient.SqlParameter("@DCN", DocmId)
                    }
                    Dim StrIndexsLog As String = ""
                    Dim StrIneexNoteLog As String = ""
                    Dim dsLog As DataSet = CommonUtilsv2.GetDataSet(dbKey, StrLogSQL, CommandType.Text, params)
                    If (Not IsNothing(dsLog)) And dsIndex.Tables(0).Rows.Count > 0 Then
                        For Each item As JProperty In results
                            Dim OldValue As String = dsLog.Tables(0).Rows(0)(item.Name).ToString()
                            Dim NewValue As String = GetYiReplace(item.Value.ToString()).Trim()
                            ' Dim NewValue As String = item.Value.ToString().Replace("&quot&", "'").Replace("!quot!", """").Replace("%quot%", ",").Replace("-quot-", ":").Replace("#quot#", "\\").Replace("!quot-", "  ")
                            If (OldValue.ToUpper().Trim() <> NewValue.ToUpper().Trim()) Then
                                Dim IndexName As String = dsIndex.Tables(0).Select(" IndexName='" & item.Name.ToString() & "'")(0)("ColumnName").ToString()
                                If (IndexName = "Notes") Then
                                    StrIneexNoteLog = NewValue & "<br>"
                                End If
                                If (String.IsNullOrEmpty(OldValue)) Then
                                    StrIndexsLog = StrIndexsLog & IndexName.Trim() & "=> Empty Changed  To " & NewValue.Trim() & ","
                                Else
                                    If (String.IsNullOrEmpty(NewValue)) Then
                                        StrIndexsLog = StrIndexsLog & IndexName.Trim() & "=>" & OldValue.Trim() & " Changed  To  Empty"
                                    Else
                                        StrIndexsLog = StrIndexsLog & IndexName.Trim() & "=>" & OldValue.Trim() & " Changed  To " & NewValue.Trim() & ","
                                    End If
                                End If
                            End If

                        Next
                    End If

                    If (Not String.IsNullOrEmpty(Notes)) Then
                        Notes = Notes & "<br>"
                    End If
                    If (Not String.IsNullOrEmpty(StrIneexNoteLog)) Then
                        Notes = Notes & StrIneexNoteLog
                    End If
                    If (Not String.IsNullOrEmpty(StrIndexsLog)) Then
                        Notes = Notes & "Update indexes:" & StrIndexsLog.TrimEnd(",")
                    End If

                End If
                ''Write log  End

                myComm.ExecuteNonQuery()

                If (Not String.IsNullOrEmpty(Notes)) Then
                    bLogCreated = CommonUtilsv2.CreateDCNLog(DocmId, CommonUtilsv2.LogType.User, GetYiReplace(Notes.TrimEnd(",")), User, HttpContext.Current.Request.UserHostAddress())
                    If bLogCreated = False Then
                        Throw New Exception("Create DCN log record failed")
                    End If
                End If

                myComm.Dispose()

                strComments = GetDCNLog(DocmId)

                ReturnStr = "{""msg"":""Update is successful."",""status"":1,""Comments"":""" + strComments.Trim() + """}"
            Else
                ReturnStr = "{""msg"":""Please fill in the value"",""status"":0}"
            End If
        Catch ex As Exception
            ' errString = ex.Message
            errLocation = "SubmintDetail() for DCN: " & DocmId.ToString & " Index : " & Index
            CommonUtilsv2.CreateErrorLog(errLocation, ex, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress())
            ReturnStr = "{""msg"":""Update failed. The detail error is : " & ex.ToString() + """,""status"":0}"
        Finally
            myComm.Dispose()
            myConn.Close()
        End Try

        Return ReturnStr
    End Function

    <WebMethod()>
    Public Shared Function ChangedOwner(ByVal ID As String, ByVal DocTypeID As String, ByVal SelectValue As String) As String
        Dim StrUser As String = HttpContext.Current.Session("User")
        If StrUser Is Nothing Then
            Return "{""msg"":""You don't have access to this Function "",""status"":-1}"
        End If
        ' status 0:error,1:success
        Dim errLocation As String
        Dim errString As String
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim ReturnStr As String = String.Empty
        Dim strComments As String = String.Empty
        Dim params As SqlParameter() = Nothing

        Dim strSQL As String = "prc_AssignOwner"
        params = {New SqlParameter("@DCN", ID), New SqlParameter("@DocOwner", SelectValue), New SqlParameter("@UserID", HttpContext.Current.Session("User").ToString())}

        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            CommonUtilsv2.RunNonQuery(dbKey, strSQL, CommandType.StoredProcedure, params)
            Dim strSQL2 As String = "Update tbl_Web_DocumentAttributes set DocType=@DocType,UpdatedBy=@UserID,UpdatedDate=getdate() where DCN=@DCN"
            Dim params2 As SqlParameter() = {New SqlParameter("@DCN", ID), New SqlParameter("@DocType", DocTypeID), New SqlParameter("@UserID", HttpContext.Current.Session("User").ToString())}
            CommonUtilsv2.RunNonQuery(dbKey, strSQL2, CommandType.Text, params2)

            CommonUtilsv2.CreateDCNLog(ID, CommonUtilsv2.LogType.User, "update owner to " & SelectValue.Replace("&quot&", "'").Replace("!quot!", """").Replace("%quot%", ",").Replace("-quot-", ":").Replace("#quot#", "\\"), HttpContext.Current.Session("User").ToString(), HttpContext.Current.Request.UserHostAddress())

            ReturnStr = "{""msg"":""Update is successful."",""status"":1}"

        Catch ex As Exception
            errString = ex.Message
            errLocation = "ChangedOwner()"
            ' CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress())
            CommonUtilsv2.CreateErrorLog(errLocation, ex, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress())
            ReturnStr = "{""msg"":""Update failed. The detail error is : " & ex.ToString() + """,""status"":0}"
        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try
        Return ReturnStr
    End Function

    <WebMethod()>
    Public Shared Function GetDCNIsExist(ByVal DCN As Integer) As String
        Dim StrUser As String = HttpContext.Current.Session("User")
        If StrUser Is Nothing Then
            Return "{""msg"":""You don't have access to this Function "",""status"":-1}"
        End If
        ' status 0:error,1:no data,2:success
        Dim ReturnStr As String = String.Empty
        Dim StrIsExist As Integer
        Dim params As SqlParameter() = Nothing
        ' Dim strSQL As String = "select Count(*) from tbl_Web_Documents where  DCN=@DCN and RecordStatus=1 "
        Dim strSQL As String = "select Count(*) from v_w_QueueRegister where  DCN=@DCN and RecordStatus=1 "
        Dim strFilter As String = CommonUtilsv2.GetDCNRecordFilterByUser(StrUser)
        If Not String.IsNullOrEmpty(strFilter) Then
            strSQL = strSQL & " AND (" & strFilter & ") "
        End If
        params = {New SqlParameter("@DCN", DCN)}
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")

        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            StrIsExist = StrHelp.GetInt(CommonUtilsv2.RunScalarQuery(dbKey, strSQL, CommandType.Text, params))
            If StrIsExist > 0 Then
                ReturnStr = "{""msg"":""this success!"",""status"":2}"

            Else
                ReturnStr = "{""msg"":""this no data!"",""status"":1}"

            End If
        Catch ex As Exception
            CommonUtilsv2.CreateErrorLog("GetDCNIsExist()", ex, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress())
            ReturnStr = "{""msg"":""" + ex.ToString() + """,""status"":0}"
        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try
        Return ReturnStr
    End Function

    <WebMethod()>
    Public Shared Function GetCountiesSelect(ByVal ID As String) As String
        Dim StrUser As String = HttpContext.Current.Session("User")
        If StrUser Is Nothing Then
            Return "{""msg"":""You don't have access to this Function "",""status"":-1}"
        End If
        ' status 0:error,1:no data,2:success,-1:redirect to  login.aspx
        Dim ReturnStr As String = String.Empty
        Dim ds As DataSet
        Dim params As SqlParameter() = Nothing

        Dim strSQL As String = "select * from Tbl_Web_Counties where  State=@State "
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        params = {New SqlParameter("@State", ID)}

        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)
            If Not ds Is Nothing Then
                Dim strjson As String = String.Empty
                For Each row As System.Data.DataRow In ds.Tables(0).Rows
                    strjson += "{""Id"":""" + row("State").ToString() + """,""Name"":""" + row("County").ToString() + """},"
                Next
                ReturnStr = "{""msg"":""this success!"",""status"":2,""data"":" + "[" & strjson.ToString().TrimEnd(",") & "]" + "}"

            Else
                ReturnStr = "{""msg"":""this no data!"",""status"":1}"

            End If
        Catch ex As Exception
            CommonUtilsv2.CreateErrorLog("GetCountiesSelect()", ex, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress())
            ReturnStr = "{""msg"":""" + ex.ToString() + """,""status"":0}"
        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try
        Return ReturnStr
    End Function

    Private Shared Function GetYiReplace(ByVal _Str As String) As String

        Return _Str.Replace("&quot&", "'").Replace("!quot!", """").Replace("%quot%", ",").Replace("-quot-", ":").Replace(vbCr, "").Replace(vbLf, "").Replace(vbCrLf, "").Replace(vbTab, "").Replace("\t", "").Replace("\b", "").Replace("\f", "").Replace("\n", "").Replace("\r", "")

    End Function

    Private Shared Function GetFuReplace(ByVal _Str As String) As String

        Return _Str.Replace("'", "&quot&").Replace("""", "!quot!").Replace(",", "%quot%").Replace(":", "-quot-").Replace(vbCr, "").Replace(vbLf, "").Replace(vbCrLf, "").Replace(vbTab, "").Replace("\t", "").Replace("\b", "").Replace("\f", "").Replace("\n", "").Replace("\r", "")

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
        function setFileName(fuID, tbID) {
            var arrFileName = document.getElementById(fuID).value.split('\\');
            document.getElementById(tbID).value = arrFileName[arrFileName.length - 1];
        }
        
        function clearFileUploader() {
            $("#NoAttachmentChange span").click();
        }

        function backfunction()
        {
            if(window.history.length > 1){
                window.history.back(-1);
            }else{
                window.opener.location.reload();
            }
        }

        $(function () {
            $("#NoAttachmentChange").attr("href", window.location.href);
            let DocmId = getUrlParam('id');
            UpdateTablehead(DocmId);
            
            $("#docIndex_block").on("change", '.selectstate', function () {
                var strid = $(this).find("option:selected").val();
                if(strid){
                    $.ajax({
                        type: "Post",
                        contentType: "application/json;charset=UTF-8",
                        url: "/Document-Detail.aspx/GetCountiesSelect",
                        data: "{ID:'" + strid + "'}",
                        success: function (result) {
                            var strjson = JSON.parse(result.d);
                            
                            if(strjson.status ==-1){
                                window.location.href="/login.aspx";
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
                }else
                {
                    var strchildoption = "<option value=''>select a value</option>";
                    $(".selectcounty").html(strchildoption);
                }
            });

            $('#ctl00_Body_tbNotes').bind('input propertychange', function() {
                $("#CommentsTextCount").text("("+$(this).val().length+")");
            });

        });

        function getUrlParam(name) {
            var reg = new RegExp("(^|&)" + name + "=([^&]*)(&|$)");
            var r = window.location.search.substr(1).match(reg);
            if (r != null) return unescape(r[2]); return null;
        }
        
        function toifram() {
            let DocmId = getUrlParam('id');
            let StrImageArchived = $("#ctl00_Body_ImageArchived").val();

            $("#myframe").attr("src", "/PDFHandler.ashx?Id=" + DocmId + "&SId=1&ImageArchived=" + StrImageArchived);
            $("#ctl00_Body_lbxAttachments tr").removeClass("lipoint");
            
        }
        function selectattac(_obj, StrImageArchived=0) {
            $(_obj).parent().parent().parent().addClass("lipoint").siblings().removeClass("lipoint");
            let StrId = $(_obj).parent().parent().parent().attr("StrId");
            $("#myframe").attr("src", "/PDFHandler.ashx?Id=" + StrId + "&SId=2&ImageArchived=" + StrImageArchived);
        }
        var $obj;
        let AttachmentID;
        function colsetattac(_obj) {
            $obj = $(_obj);
            var strwidth = $("body").width()/2-100;
            AttachmentID = $obj.parent().parent().attr("StrId");
            $(".mod-dialog").css("left", strwidth + "px");
            $(".mod-dialog").css("display", "block");
            $(".mod-dialog-bg").css("display", "block");
            $(".dialog-content").find("p").text("Are you sure you want to delete " + $obj.parent().parent().find("td:eq(0)").find("a").text());
           
        }
        function DialogCancel() {
            $(".mod-dialog").css("display", "none");
            $(".mod-dialog-bg").css("display", "none");
        }
        function DialogOK() {
            $.ajax({
                type: "Post",
                contentType: "application/json;charset=UTF-8",
                url: "/Document-Detail.aspx/DeleteAttac",
                data: "{ID:" + AttachmentID + "}",
                success: function (result) {
                    var strjson = JSON.parse(result.d.replace(/(\t)+/g,'\\t'));
                    if(strjson.status ==-1){
                        window.location.href="/login.aspx";
                    }
                    if (strjson.status == 1) {
                        $obj.parent().parent().remove();
                        $(".mod-dialog").css("display", "none");
                        $(".mod-dialog-bg").css("display", "none");

                        var strcommentlist=ReturnTrunchar(strjson.Comments)
                        var reg1 = new RegExp("&", "g");
                        var reg2 = new RegExp("<br>", "g");
                        $("#ctl00_Body_lbxAccumulatedComment").html(strcommentlist.replace(reg1, ":").replace(reg2, "\r\n"))
                       
                        setTimeout(function () { $("#ctl00_Body_lbSubmitComment").hide("3000"); }, 3000);

                    } else {
                     $("#ctl00_Body_lbSubmitComment").attr("style", "color:red;font-weight: bold;").text(strjson.msg);
                    }
                }
            });
        }
        function UpdateTablehead(DocmId) {
            $.ajax({
                type: "Post",
                contentType: "application/json;charset=UTF-8",
                url: "/Document-Detail.aspx/GetIndexlist",
                data: "{DocmID:" + DocmId + "}",
                success: function (result) {
                    var strjson = JSON.parse(result.d.replace(/\\/g, '\\\\'));
                    var strIndexes;
                    if(strjson.status ==-1){
                        window.location.href="/login.aspx";
                    }

                    if (strjson.status == 2) { 
                        var strIsNotCompleteAuthority=<%= Convert.ToInt16(IsNotCompleteAuthority)%>;
                        var strIsAuthority=<%= Convert.ToInt16(IsAuthority)%>;
                        var isstyle="";
                        var isinputReadonly="";
                        var isdateReadonly="onclick=\"WdatePicker({lang:'en',dateFmt:'MM/dd/yyyy'})\"";
                        var isselectReaonly="";
                        if (strIsNotCompleteAuthority==0 || strIsAuthority==0)//Authority
                        {
                            isinputReadonly="readonly='readonly'";
                            isstyle="background:#d8d2d2";
                            isdateReadonly="";
                            isselectReaonly="disabled='disabled'";
                            //$("#ctl00_Body_ddDocUser").attr("disabled","disabled");
                        }

                        $(strjson.data).each(function (i, dom) {
                            if (i % 3 == 0) {
                                strIndexes += "<tr>";
                            }
                            if (dom.Type.toUpperCase().trim() == "DATE") {
                                if (dom.Name.toUpperCase().trim() == "LOAD DATE") {
                                    strIndexes += "<td width='15%' align='right' class='form_label_medium' >" + dom.Name + ":&nbsp;</td><td width='18%' align='left' class='form_input_small Reassignment'><input id='tb" + dom.Index + "' MaxLength='200'  style='height:25px;width:170px;background: #d8d2d2;' readonly='readonly'  rel=" + dom.Index + "   value='" + ReturnTrunchar(dom.Value) + "'  />";

                                } else {
                                    strIndexes += "<td width='15%' align='right' class='form_label_medium' >" + dom.Name + ":&nbsp;</td><td width='18%' align='left' class='form_input_small Reassignment'><input id='tb" + dom.Index + "'  rel=" + dom.Index + "  MaxLength='200'  style='height:25px;width:170px;" + isstyle + "' readonly='readonly'  class='Wdate'  " + isdateReadonly + "  value='" + ReturnTrunchar(dom.Value) + "'  />";
                                }

                            } else if (dom.Type.toUpperCase().trim() == "SELECT") {
                                var strchildoption = "<option value=''>select a value</option>";
                                var stronclicclass = "";
                                //state
                                if (dom.Name.toUpperCase().trim().indexOf("STATE") >= 0) {
                                    stronclicclass="class='selectstate'"
                                    $(dom.data).each(function (i1, dom1) {
                                        if (dom.Value.trim() == dom1.Id) {
                                            strchildoption += "<option value='" + dom1.Id + "' selected='selected'>" + dom1.Name + "</option>";
                                        } else {
                                            strchildoption += "<option value='" + dom1.Id + "'>" + dom1.Name + "</option>";
                                        }
                                    });
                                }
                                //county
                                if (dom.Name.toUpperCase().trim().indexOf("COUNTY") >= 0) {
                                    stronclicclass = "class='selectcounty'"
                                    $(dom.data).each(function (i2, dom2) {
                                        if (dom.Value.trim() == dom2.Name) {
                                            strchildoption += "<option value='" + dom2.Name + "'  selected='selected'>" + dom2.Name + "</option>";
                                        } else {
                                            strchildoption += "<option value='" + dom2.Name + "'>" + dom2.Name + "</option>";
                                        }
                                    });
                                }

                                strIndexes += "<td width='15%' align='right' class='form_label_medium' >" + dom.Name + ":&nbsp;</td><td width='18%' align='left' class='form_input_small Reassignment ReSelect'><select id='tb" + dom.Index + "' style='height:25px;width:170px;"+isstyle+"'" + stronclicclass + "  rel='" + dom.Index + "' "+isselectReaonly+">" + strchildoption + "<select>";

                            } else {

                                strIndexes += "<td width='15%' align='right' class='form_label_medium' >" + dom.Name + ":&nbsp;</td><td width='18%' align='left' class='form_input_small Reassignment'><input id='tb" + dom.Index + "' MaxLength='200'  style='height:25px;width:170px;"+isstyle+"'  rel=" + dom.Index + "  value='" + ReturnTrunchar(dom.Value) + "' "+isinputReadonly+" />";
                            }
                           
                            strIndexes += "</td>";

                            if (i % 3 == 2) {
                                strIndexes += "</tr>";
                            }
                        });
                        $("#tableinput tbody").prepend(strIndexes);
                        
                    }
                }
            });
        }

        //&#25552;&#20132;
        function Submitdatail()
        {
            let DocmId = getUrlParam('id');
            var strdata = "";
            var insertchildrenjson={};
            $("#ctl00_Body_Table3 .Reassignment").each(function () {
                if ($(this).hasClass("ReSelect")) {
                    var keystr = $(this).find("select").attr("rel");
                    var valstr = $(this).find("select").find("option:selected").val();
                    if (valstr == null) {
                        insertchildrenjson[keystr]="";
                    } else {
                        insertchildrenjson[keystr]=valstr;
                    }

                } else {
                    var keystr = $(this).find("input").attr("rel");
                    var valstr =$(this).find("input").val();//读取时转义
                    if (valstr == null) {
                        insertchildrenjson[keystr]="";
                    } else {
                        insertchildrenjson[keystr] = ReturnToPunctuation(qudiaoAll(valstr));
                    }
                }
            });

            var strnotes = ReturnToPunctuation($("#ctl00_Body_tbNotes").val().replace(/\\r/g, "").replace(/\\n/g, "").replace(/\\t/g, "").replace(/\\b/g, "").replace(/\\f/g, ""));
            
            var Struser = $("#ctl00_Body_tableuser").val();
            $.ajax({
                type: "Post",
                contentType: "application/json;charset=UTF-8",
                url: "/Document-Detail.aspx/SubmintDetail",
                data: "{DocmId:" + DocmId + ",Notes:'" + strnotes + "',Index:'" + JSON.stringify(insertchildrenjson) + "',User:'" + Struser + "'}",
                success: function (result) {
                    console.log(result.d);
                    var strjson = JSON.parse(result.d);
                    if(strjson.status ==-1){
                        window.location.href="/login.aspx";
                    }
                    if (strjson.status == 1) {
                        $("#ctl00_Body_lbSubmitComment").show();
                        $("#ctl00_Body_lbSubmitComment").attr("style", "color:blue;font-weight: bold;").text(strjson.msg);
                       
                        $("#ctl00_Body_lbxAccumulatedComment").html(ReturnTrunchar(strjson.Comments));//提交成功后转为符号
                        setTimeout(function () { $("#ctl00_Body_lbSubmitComment").hide("3000"); }, 3000);
                        
                    } else {
                        $("#ctl00_Body_lbSubmitComment").attr("style", "color:red;font-weight: bold;").text(strjson.msg);
                    }

                    $("#ctl00_Body_tbNotes").val("");
                }
            });
        }

        function btnOwnerFunction()
        {
            $("#btnok1").css("display", "none");
            $("#btnok2").css("display", "");
            var strwidth = $("body").width() / 2 - 100;           
            $(".mod-dialog").css("left", strwidth + "px");
            $(".mod-dialog").css("display", "block");
            $(".mod-dialog-bg").css("display", "block");
            $(".dialog-content").find("p").text("Are you sure you want to change owner or Doc Type  " );
        }

        function DialogOwnerOK() {
            var strid = $("#ctl00_Body_lblDCN").text();
            var strdoctype = $("#ctl00_Body_ddDocType").find("option:selected").val();
            var strselectvalue = $("#ctl00_Body_ddDocUser").find("option:selected").text();
            $.ajax({
                type: "Post",
                contentType: "application/json;charset=UTF-8",
                url: "/Document-Detail.aspx/ChangedOwner",
                data: "{ID:'" + strid + "',DocTypeID:'" + strdoctype + "',SelectValue:'" + strselectvalue + "'}",
                success: function (result) {
                    var strjson = JSON.parse(result.d.replace(/(\t)+/g,'\\t'));
                    if(strjson.status ==-1){
                        window.location.href="/login.aspx";
                    }
                    if (strjson.status == 1) {
                        var reg1 = new RegExp("&", "g");
                        var reg2 = new RegExp("<br>", "g");
                        window.location.reload()
                    } else {
                        $("#ctl00_Body_lbSubmitComment").attr("style", "color:red;font-weight: bold;").text(strjson.msg);
                    }
                }
            });
        }

        function ReturnTrunchar(_Strobj)
        {
            if (_Strobj.indexOf("-rquotn-") != -1) {
                var reg = new RegExp("-rquotn-", "g");
                _Strobj = _Strobj.replace(reg, /\r\n/);
            }
            if (_Strobj.indexOf("rquotr") != -1) {
                var reg = new RegExp("rquotr", "g");
                _Strobj = _Strobj.replace(reg, '\r');
            }
            if (_Strobj.indexOf("nquotn") != -1) {
                var reg = new RegExp("nquotn", "g");
                _Strobj = _Strobj.replace(reg, '\n');
            }
            if(_Strobj.indexOf("&quot&")!=-1)
            {
                var reg = new RegExp("&quot&","g");
                _Strobj=_Strobj.replace(reg, '&apos;');
            }
            if(_Strobj.indexOf("%quot%")!=-1)
            {  
                var reg = new RegExp("%quot%","g");
                _Strobj=_Strobj.replace(reg, ',');
            }
            if(_Strobj.indexOf("-quot-")!=-1)
            {
                var reg = new RegExp("-quot-","g");
                _Strobj=_Strobj.replace(reg, ':');
            }

            if(_Strobj.indexOf("!quot-")!=-1)
            {
                var reg = new RegExp("!quot-","g");
                _Strobj=_Strobj.replace(reg, '<br/>');
            }
            if(_Strobj.indexOf("!quot>")!=-1)
            {
                var reg = new RegExp("!quot>","g");
                _Strobj=_Strobj.replace(reg, '&nbsp;&nbsp;');
            }
            if (_Strobj.indexOf("!quot!") != -1) {
                var reg = new RegExp("!quot!", "g");
                _Strobj = _Strobj.replace(reg, "\"");
            }
            
            
            return _Strobj;
        }

        function ReturnToPunctuation(_Strobj){
            //valstr.replace("'", "&quot&").replace("\"", "!quot!").replace(",", "%quot%").replace(":", "-quot-").replace("\\", "#quot#")
            _Strobj = _Strobj.replace(/\r\n/g, "-rquotn-");

           
            if (_Strobj.indexOf("\\r") != -1) {
                var reg = new RegExp("\\r", "g");
                _Strobj = _Strobj.replace(reg, 'rquotr');
            }
            if (_Strobj.indexOf("\\n") != -1) {
                var reg = new RegExp("\\n", "g");
                _Strobj = _Strobj.replace(reg, 'nquotn');
            }
            if (_Strobj.indexOf("'") != -1) {
                var reg = new RegExp("'", "g");
                _Strobj = _Strobj.replace(reg, '&quot&');
            }
            if (_Strobj.indexOf("'") != -1) {
                var reg = new RegExp("'", "g");
                _Strobj = _Strobj.replace(reg, '&quot&');
            }
            if(_Strobj.indexOf("'")!=-1)
            {
                var reg = new RegExp("'","g");
                _Strobj=_Strobj.replace(reg, '&quot&');
            }
            if (_Strobj.indexOf(/\"/g)!=-1)
            {
                //var reg = new RegExp("\"","g");
                _Strobj = _Strobj.replace(/\"/g, '!quot!');
            }
            if (_Strobj.indexOf("\"") != -1) {
                var reg = new RegExp("\"","g");
                _Strobj = _Strobj.replace(reg, '!quot!');
            }
            if(_Strobj.indexOf(",")!=-1)
            {  
                var reg = new RegExp(",","g");
                _Strobj=_Strobj.replace(reg, '%quot%');
            }
            if(_Strobj.indexOf(":")!=-1)
            {
                var reg = new RegExp(":","g");
                _Strobj=_Strobj.replace(reg, '-quot-');
            }           
           
            return _Strobj;
        }

        function qudiaoAll(_Strobj) {
            _Strobj = _Strobj.replace(/\\r/g, "").replace(/\\n/g, "").replace(/\\t/g, "").replace(/\\b/g, "").replace(/\\f/g, "").replace(/\\\"/g, "");
            _Strobj = _Strobj.replace(/\r/g, "").replace(/\n/g, "").replace(/\t/g, "").replace(/\b/g, "").replace(/\f/g, "").replace(/\""/g, "");

            return _Strobj;
        }
       
    </script>
    <div align="center" id="documentdetail" runat="server" >
        <a href="" id="NoAttachmentChange"><span></span></a>  <%--This label is to imitate the label to re-click to prevent "Attachment" from resubmitting data --%>
        <div class="body_title">
          <a href="javascript:void(0);" onclick="backfunction()"><span>back</span></a>
           Document-Detail
        </div>
        <input id="tableuser" type="hidden" runat="server" />
        <input id="ImageArchived" type="hidden" runat="server" />
        <%=StrdocumentCopy %>
        
        <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0">
            <tr class="header_spacer">
                <td colspan="6">&nbsp;</td>
            </tr>
            <tr>
                <td colspan="6">
                    <div title="Click to Expand" id="docSummary_expand" style="display: none;">
                        <img alt="" src="Images/expand.gif" width="12" height="12" onclick="expandBlock('docSummary');" /></div>
                    <div title="Click to Hide" id="docSummary_collapse" style="display: inline;">
                        <img alt="" src="Images/collapse.gif" width="12" height="12" onclick="collapseBlock('docSummary');" /></div>
                    <div class="form_label_medium_nobackground" style="display: inline;"><b>Document Info</b></div>
                </td>
            </tr>
            <tr class="popupbody_simple">
                <td colspan="6" align="center">
                    <div id="docSummary_block">
                        <table width="100%"  border="0" class="claimdetail_summaryblock">
                            <tr>
                                <td width="12%" align="right" class="form_label_medium">DCN#:&nbsp;</td>
                                <td width="18%" align="left" class="form_label_medium_nobackground">
                                    <a href="javascript:void(0);" onclick="toifram()" ><asp:Label ID="lblDCN" runat="server"></asp:Label></a>
                                    <asp:Label ID="lblDownloadDCN" runat="server"></asp:Label>  
                                </td>
                                <td width="40%" colspan="2" align="center" class="form_label_medium">Comment:&nbsp;</td>
                                <td width="30%" colspan="2" align="center" class="form_label_medium">Attachments:&nbsp;</td>
                            </tr>
                            <tr>
                                <td colspan="2" align="right" class="form_label_medium">
                                    <% If IsNotCompleteAuthority And IsReplaceDCNImage Then %>
                                    <asp:FileUpload ID="FileUpload2" runat="server" Style="border: solid 1px #00437a;"></asp:FileUpload>
                                    &nbsp;&nbsp;
                                    <asp:Button ID="btnReplaceImage" runat="server" Text="Replace" CssClass="owner-btn-submit" OnClick="ReplaceImage" TabIndex="105" CausesValidation="false" ValidationGroup="uploadAction" />
                                    <%End If %>
                                </td>
                                 <td colspan="2" rowspan="6" align="center" valign="top" class="form_label_medium_nobackground">
                                    <%--<asp:TextBox ID="tbAccumulatedComment" runat="server" Enabled="false" ReadOnly="true" TextMode="MultiLine" Height="120" Width="95%" Wrap="true"></asp:TextBox>--%>
                                    <div id="lbxAccumulatedComment" runat="server"></div>
                                </td>
                                <td colspan="2" rowspan="6" align="center"  valign="top" class="form_label_medium_nobackground">
                                      <div id="lbxAttachments" runat="server"></div>
                                </td>                              
                            </tr>
                            <tr>
                                <td align="right" class="form_label_medium">File Name:&nbsp;</td>
                                <td align="left" class="form_label_medium_nobackground">
                                    <asp:Label ID="lblFileName" runat="server"></asp:Label>
                                </td>
                               
                            </tr>
                            <tr>
                                <td align="right" class="form_label_medium">Queue Name:&nbsp;</td>
                                <td align="left" class="form_label_medium_nobackground">
                                    <asp:Label ID="lblQueueName" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" class="form_label_medium">Uploaded Date:&nbsp;</td>
                                <td align="left" class="form_label_medium_nobackground">
                                    <asp:Label ID="lblUploadedDate" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" class="form_label_medium">Doc Status:&nbsp;</td>
                                <td align="left" class="form_label_medium_nobackground">
                                    <asp:Label ID="lblDocStatus" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" class="form_label_medium">Doc Type:&nbsp;</td>
                                <td align="left" class="form_label_medium_nobackground">
                                    <%--<asp:Label ID="lblDocType" runat="server"></asp:Label>--%>
                                    <asp:DropDownList ID="ddDocType" runat="server" Width="70%" AppendDataBoundItems="true" TabIndex="102" Height="25px"></asp:DropDownList>
                                </td>
                            </tr>
                            <tr><td colspan="6"></td></tr>
                            <tr>
                                <td align="right" class="form_label_medium">Owner:&nbsp;</td>
                                <td align="left" class="form_label_medium_nobackground">                                    
                                    <asp:DropDownList ID="ddDocUser" runat="server" Width="70%" AppendDataBoundItems="true" TabIndex="211" Height="25px"></asp:DropDownList>
                                    <%-- <% If IsNotCompleteAuthority And IsAuthority Then %>    --%>                                   
                                        <input type="button" value="Submit" class="owner-btn-submit" onclick="btnOwnerFunction()"/>                                                             
                                   <%-- <%End If %>--%>
                                </td>
                                <td colspan="4" align="right"  valign="top" class="form_label_medium_nobackground">
                                <% If IsAuthority Then %>
                                Add Attachment:&nbsp;<asp:TextBox ID="tbNewFile" runat="server" Text="" Width="210px" Height="25px" ></asp:TextBox>
                                <label class="file-upload" style="line-height: 18px;">
                                        <span class="Submit_button">Browse</span>
                                        <asp:FileUpload ID="FileUpload1" runat="server" onchange="setFileName()"></asp:FileUpload>
                                    </label>
                                     &nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="btnUploadAttachmentc" Width="160px" runat="server" Text="Upload Attachment" CssClass="Submit_button" OnClick="UploadAttachment" TabIndex="105" CausesValidation="false" ValidationGroup="uploadAction" />&nbsp;&nbsp;&nbsp;
                                   <%End If %>   
                                </td>                                 
                            </tr>
                            <%--<% If GroupID = 1 And Not IsVisibleInDelete Then%> --%>
                            <tr>
                                <td align="right" class="form_label_medium">Delete Notes:&nbsp;</td>
                                <td align="left" colspan="3" class="form_label_medium_nobackground">                                    
                                   <asp:TextBox ID="tbDeleteNotes" runat="server"  MaxLength="1000" Height="32" Enabled="true" Width="95%" Wrap="true"></asp:TextBox>
                                </td>
                                <td colspan="2" align="left"  valign="middle" class="form_label_medium_nobackground">
                                                             
                                     &nbsp;&nbsp;&nbsp;<asp:Button ID="btnMoveDCNForDelete" Width="160px" runat="server" Text="Move DCN for Delete" CssClass="Submit_button" OnClick="MoveDCNForDelete" Enabled="true" TabIndex="108" CausesValidation="false"  />&nbsp;&nbsp;&nbsp;
                                    <asp:Button ID="btnDeleteDCN" Width="160px" runat="server" Text="Delete DCN" CssClass="Submit_button" OnClick="DeleteDCN" Enabled="true" TabIndex="109" CausesValidation="false"  />&nbsp;&nbsp;&nbsp;
              
                                </td>                                 
                            </tr>
                            <%-- <%  End If %> --%>
                        </table>
                    </div>
                </td>
            </tr>
            <tr class="header_spacer">
                <td colspan="6">&nbsp;</td>
            </tr>
            <tr>
                <td colspan="6">
                    <div title="Click to Expand" id="docIndex_expand" style="display: none;">
                        <img alt="" src="Images/expand.gif" width="12" height="12" onclick="expandBlock('docIndex');" /></div>
                    <div title="Click to Hide" id="docIndex_collapse" style="display: inline;">
                        <img alt="" src="Images/collapse.gif" width="12" height="12" onclick="collapseBlock('docIndex');" /></div>
                    <div class="form_label_medium_nobackground" style="display: inline;"><b>Document Indexes</b></div>
                </td>
            </tr>
            <tr class="popupbody_simple">
                <td colspan="6" align="center">
                    <div id="docIndex_block">
                        <table runat="server" id="Table3" width="100%" border="0" class="claimdetail_claimactionblock">
                            <tr class="popupbody_simple">
                                <td colspan="6" align="left" class="form_input" style="border-bottom: solid 1px #00658c"></td>
                            </tr>
                            <tr class="header_spacer">
                                <td colspan="6">&nbsp;</td>
                            </tr>
                            <tr>
                                <td colspan="6">
                                    <table width="100%" border="0"  id="tableinput">
                                        <tr> 
                                            <td width="15%" align="right" class="form_label_medium">Comments<label id="CommentsTextCount" style="color:red;">(0)</label>:&nbsp;</td>
                                            <td align="left" colspan="5" valign="top" width="85%"  class="popupbody_simple">
                                                <asp:TextBox ID="tbNotes" runat="server" Enabled="true" TextMode="MultiLine" Height="100" Width="95%" Wrap="true" TabIndex="24"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                             <td colspan="6" align="right" valign="bottom" class="form_label_medium">
                                              <%--<% If IsAuthority Then %>--%>
                                              <% If IsNotCompleteAuthority And IsAuthority Then %>
                                               <button id="submitIndex"  type="button"  onclick="Submitdatail()"  class="btn-datail" >Submit</button>
                                              <%End If %>
                                            </td>                                       
                                       </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr class="popupbody_simple">
                                <td align="left" valign="top" colspan="6">
                                    <asp:Label ID="lbSubmitComment" runat="server" CssClass="body_text_errors"></asp:Label>
                                </td>
                            </tr>
                        </table>
                        <%--<asp:HiddenField ID="hfOwnerName" runat="server" ClientIDMode="Static"/>
                        <asp:HiddenField ID="hfUser" runat="server" ClientIDMode="Static"/>--%>
                    </div>
                </td>
            </tr>
            <tr class="header_spacer">
                <td colspan="6">&nbsp;</td>
            </tr>
            <tr class="header_spacer">
                <td colspan="6">
                    <asp:Literal runat="server" ID="DocumentViewer" Text="Document Window" ></asp:Literal></td>
            </tr>
        </table>
    </div>
    <div class="mod-dialog-bg" style="display: none;"></div>
    <div class="mod-dialog" style=" top: 368.5px; display: none;"><div class="dialog-nav"><span class="dialog-title">Confirm</span><a href="#" onclick="return false" class="dialog-close"></a></div><div class="dialog-main"><div class="dialog-content"><p></p></div><div class="dialog-console clearfix_new" style="margin-top: 20px;"><a class="console-btn-confirm" href="#" onclick="DialogOK()"  id="btnok1">OK</a><a class="console-btn-confirm" href="#" onclick="DialogOwnerOK()" id="btnok2" style="display:none;">OK</a><a class="console-btn-cancel" href="#" onclick="DialogCancel()">Cancel </a></div></div></div>
    <div id="nodatadiv"  runat="server" style="width: 100%;height: 400px;line-height: 400px;text-align: center;font-size: 30px;color: red;font-weight: bold;">This DCN is not found. </div>
</asp:Content>
