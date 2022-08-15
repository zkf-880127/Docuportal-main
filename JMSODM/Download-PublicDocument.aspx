<%@ Page Language="VB" MasterPageFile="PageMaster.master" AutoEventWireup="false" EnableEventValidation="false" ASPCOMPAT="TRUE" %>
<%@ MasterType VirtualPath="PageMaster.master" %>
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
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        pageName = Request.RawUrl.ToString
        Try
            Master.SetCurrentMenuItem = System.IO.Path.GetFileName(Request.RawUrl.ToString)
        Catch ex As Exception
            Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
            Exit Sub
        End Try
        Dim anID As String = Request.QueryString.Get("Document")
        
        If CustomRoles.RolesForPageLoad() Then
            If CommonUtilsv2.Validate(anID, CommonUtilsv2.DataTypes.Int, True, True, True) Then
                DownloadDocument(anID)
            Else
                If String.Compare(System.IO.Path.GetFileName(Webapps.Utils.ApplicationSettings.Homepage), System.IO.Path.GetFileName(Request.RawUrl.ToString), True) = 0 Then
                    ' Do Nothing
                Else
                    Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)     'QueryString is not valid; Redirect to web.config defined home page
                End If
            End If
           
        Else
            CustomRoles.TransferIfNotInRole()
            Exit Sub
        End If
    End Sub
    
    Protected Sub DownloadDocument(ByVal strDocumentID As String)
        Dim written As Boolean = False
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
        Dim sqlCmd As New SqlClient.SqlCommand
        Dim sqldr As SqlClient.SqlDataReader
        sqlCmd.Connection = sqlCnn
        sqlCmd.CommandTimeout = 0
        sqlCmd.CommandText = "prc_Download_Public_Document"
        sqlCmd.CommandType = CommandType.StoredProcedure
        sqlCmd.Parameters.Add(New SqlClient.SqlParameter("@ID", strDocumentID))
        Dim stopme As Boolean = True
        Try
            sqlCnn.Open()
            sqldr = sqlCmd.ExecuteReader()
            While sqldr.Read
                If sqldr.Item(1) Is System.DBNull.Value Then
                    errString = "Cannot find or display Public document " & strDocumentID
                    errLocation = "DownloadDocument()"
                     CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
                Else
                    Dim Filename As String
                    Filename = sqldr.Item(0)
                    'Dim decompressed As New Aced.Compression.AcedInflator
                    'Dim bytearr As Byte() = decompressed.Decompress(DirectCast(sqldr.Item(1), Byte()), 0, 0, 0)
                    Dim bytearr As Byte() = DirectCast(sqldr.Item(1), Byte())
                    Response.AddHeader("Content-disposition", "attachment; filename=" & ControlChars.Quote & Filename & ControlChars.Quote)
                    Response.ContentType = "application/octet-stream"
                    'Write Array to Browser
                    Response.BinaryWrite(bytearr)
                    written = True
                    Response.End()
                End If
            End While
            sqldr.Close()
            sqlCnn.Close()
        Catch ex As Exception
            errString = ex.Message
            errLocation = "DownloadDocument"
            If Not String.IsNullOrEmpty(errString) AndAlso Not errString.ToLower().Contains("thread was being aborted") Then
                CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress())
            End If
        Finally
            sqlCnn.Dispose()
            sqlCmd.Dispose()
            sqlCnn = Nothing
            sqlCmd = Nothing
        End Try
    End Sub
</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
    <div align="center">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="wrapper">
		    <tr>
	            <td class="body" colspan="2">
                    <table width="100%">
		                <tr class="body_title">
		                    <td colspan="7" align="left" class="body_title">Document Download</td>
		                </tr>
		                <tr class="body_plain" colspan="2">
                            <td align="left"><br />
                                We're sorry . . . that document is not available right now.&nbsp;&nbsp;<b>Please try back later.</b> 
                                <br />
                                <br />
                                A message has been sent to the system administrator to notify them of this condition.
                                <br />
                                <br />
                            </td>
                        </tr>
                    </table>
                 </td>
		      </tr>        
           </table> 
        </div>
</asp:Content>