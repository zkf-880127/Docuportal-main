<%@ Page Language="VB" MasterPageFile="PageMaster.master" AutoEventWireup="false" EnableEventValidation="false" viewStateEncryptionMode="Auto" ASPCOMPAT="TRUE" ValidateRequest="True" %>
<%@ MasterType VirtualPath="PageMaster.master" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="Webapps.Utils" %>


<script language="VB" runat="server">

    Dim pageName As String
    Dim strPageName As String = ""
    Dim errString As String
    Dim errLocation As String
    '----------------------------------------------------------------------------
    '
    '----- Custom / Page specific Code Section --------------------------------------------
    
    
    '----- Standardized Code Section -----------------------------------------------------------
    '----- Note:  These code blocks generally do not need to be modifed-------------------------
    '-------------------------------------------------------------------------------------------
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        pageName = Request.RawUrl.ToString
        Try
            Master.SetCurrentMenuItem = System.IO.Path.GetFileName(Request.RawUrl.ToString)
        Catch ex As Exception
            Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
            Exit Sub
        End Try
        
        If CommonUtilsv2.Validate(Request.QueryString.Get("name"), CommonUtilsv2.DataTypes.String, True, True, True) Then   ' Do Nothing
        Else
            If String.Compare(System.IO.Path.GetFileName(Webapps.Utils.ApplicationSettings.Homepage), System.IO.Path.GetFileName(Request.RawUrl.ToString), True) = 0 Then
                
            Else
                Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)     'QueryString is not valid; Redirect to web.config defined home page
            End If
        End If
        If CustomRoles.RolesForPageLoad() Then  'If security/role access is good, do nothing else transfer
        Else
            CustomRoles.TransferIfNotInRole()
            Exit Sub
        End If
        strPageName = Request.QueryString.Get("name")
        If CommonUtilsv2.Validate(strPageName, CommonUtilsv2.DataTypes.String, True, True, False, 20) Then
        Else
            strPageName = "Home"
        End If
        GetCurrentHomePageContent()
    End Sub
       
    Protected Sub GetCurrentHomePageContent()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
        Dim sqlCmd As New SqlCommand("SELECT TOP 1 ID, htmlPageContent, PageTitle FROM tbl_Web_Page_Content WHERE RecordStatus = 1 AND PageName = @PageName ORDER BY DateLastUpdated DESC", sqlCnn)
        If String.IsNullOrEmpty(strPageName) Then
            sqlCmd.Parameters.AddWithValue("@PageName", "Home")
        Else
            sqlCmd.Parameters.AddWithValue("@PageName", strPageName)
        End If      
        Dim htmlContent As String = ""
        Dim myReader3 As SqlDataReader = Nothing
        Try
            sqlCnn.Open()
            myReader3 = sqlCmd.ExecuteReader()
            While myReader3.Read()
                If myReader3.GetValue(1) Is System.DBNull.Value Then    '---------- Current Page Content ----------
                    htmlContent = ""
                Else
                    htmlContent = myReader3.GetString(1)
                End If
                If myReader3.GetValue(2) Is System.DBNull.Value Then    '---------- Current Page Title ----------
                    LiteralBodyTitle.Text = ""
                Else
                    LiteralBodyTitle.Text = myReader3.GetString(2)
                End If
            End While
        Catch ex As Exception
            errString = ex.Message
            errLocation = "Get Current Home Page Content"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myReader3 Is Nothing AndAlso Not myReader3.IsClosed Then
                myReader3.Close()
            End If
            sqlCmd.Dispose()
            sqlCmd = Nothing
            sqlCnn.Close()
        End Try
        HomePageContent.Text = htmlContent
    End Sub
</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
<script type="text/javascript">
    function backfunction() {
        if (window.history.length > 1) {
            window.history.go(-1);
        } else {
            window.opener.location.reload();
        }
    }
</script>
    <div align="center">
        <table width="100%" runat="server" border="0" cellpadding="0" cellspacing="0">
              <tr>
                  <td colspan="1" align="left" class="body_title" >
                     <a href="javascript:void(0);" onclick="backfunction()"><span>back</span></a>
                      <asp:Literal ID="LiteralBodyTitle" runat="server"></asp:Literal>
                   </td>
              </tr>
              <tr>
                  <td align="left">
                    <br />
                    <asp:Literal runat="server" ID="HomePageContent" Text="No Content Available." Mode="Transform"></asp:Literal>
                  </td>
              </tr>
        </table>
    </div>
</asp:Content>