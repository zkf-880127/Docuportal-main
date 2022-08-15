<%@ Page Language="VB" MasterPageFile="PageMaster.master" AutoEventWireup="false" EnableEventValidation="false" viewStateEncryptionMode="Auto" ASPCOMPAT="TRUE" %>
<%@ MasterType VirtualPath="PageMaster.master" %>
<%@ Register TagPrefix="obout" Namespace="Obout.Grid" Assembly="obout_Grid_NET" %>
<%@ Register TagPrefix="oint" Namespace="Obout.Interface" Assembly="obout_Interface" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.Page" %>
<%@ Import Namespace="Webapps.Utils" %>

<script language="VB" runat="server">
    '************* Error logging Section ***********************
    Dim errLocation As String
    Dim errString As String
    Dim pageName As String

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
        If CustomRoles.RolesForPageLoad() Then
            CreateGrid()
        Else
            CustomRoles.TransferIfNotInRole()
            Exit Sub
        End If
    End Sub

    Dim strCommandBase As String = "SELECT ID, SiteID, IPAddress, UserID, ErrorMessage, ErrorDateTime, ErrorLocation, ErrorPage " & _
                               " FROM tbl_Web_ErrorLog ORDER BY ErrorDateTime DESC"

    Private Sub CreateGrid()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim sqlDR As SqlDataReader
        Try
            sqlDR = CommonUtilsv2.GetDataReader(dbKey, strCommandBase, CommandType.Text)
            gridErrorLog.DataSource = sqlDR
            gridErrorLog.DataBind()
        Catch ex As Exception
            errString = ex.Message
            errLocation = "CreateGrid()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try
    End Sub

</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
    <script type="text/javascript">

    </script>
    <div align="center">
        <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0" class="wrapper">
          <tr>
             <td colspan="5" align="left" class="body_title" >
                View Error Log
             </td>
          </tr>
          <tr><td>&nbsp;</td></tr>
                      <tr>
                <td align="left">
                    <obout:Grid id="gridErrorLog" runat="server" CallbackMode="true" Serialize="true" AutoGenerateColumns="false"
                        PageSizeOptions="1,5,10,15,20" PageSize="20"
                        Width="100%" FolderStyle="styles/style_13" EnableRecordHover="true" AllowSorting="true"
                        AllowRecordSelection="false" AllowMultiRecordSelection="false" KeepSelectedRecords="false"
                        AllowAddingRecords="false" AllowFiltering="true" ShowLoadingMessage="true"
                        AllowGrouping="true" ShowGroupsInfo="false" ShowColumnsFooter="false"
                        ShowGroupFooter="false">
                        <Columns>
                            <obout:Column ID="Column5" DataField="ID" Width="0%" Visible="false" />
                            <obout:Column ID="Column1" HeaderText="Site ID" DataField="SiteID" Align="left" Width="10%" Wrap="true" Visible="true" AllowFilter="false" />
                            <obout:Column ID="Column2" HeaderText="IP Address" DataField="IPAddress" Align="left" Width="10%" Visible="true" AllowFilter="true" />
                            <obout:Column ID="Column3" HeaderText="User ID" DataField="UserID" Align="left" Width="10%" Visible="true" AllowFilter="true" />
                            <obout:Column ID="Column6" HeaderText="Message" DataField="ErrorMessage" Align="left" Width="40%" Visible="true" AllowFilter="true" Wrap="true" />
                            <obout:Column ID="Column7" HeaderText="Date" DataField="ErrorDateTime" Align="left" Width="10%" Visible="true" AllowFilter="true" Wrap="true" />
                            <obout:Column ID="Column8" HeaderText="Location" DataField="ErrorLocation" Align="left" Width="10%" Visible="true" AllowFilter="true" Wrap="true" />
                            <obout:Column ID="Column4" HeaderText="Page" DataField="ErrorPage" Align="left" Width="10%" Visible="true" AllowFilter="true" Wrap="true" />
                        </Columns>
                    </obout:Grid>
                </td>
            </tr>
        </table>
    </div>
</asp:Content>