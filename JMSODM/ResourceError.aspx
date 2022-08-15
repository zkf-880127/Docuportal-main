<%@ Page Language="VB" MasterPageFile="PageMaster.Master" AutoEventWireup="false" EnableEventValidation="false" ASPCOMPAT="TRUE" %>
<%@ MasterType VirtualPath="PageMaster.Master" %>
<%@ Import Namespace="Webapps.Utils" %>

<script language="VB" runat="server">
    '---- Error logging -----
    Dim pageName As String
    Dim errLocation As String
    Dim errString As String
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        pageName = Request.RawUrl.ToString  'initialize the page name
        Master.SetCurrentMenuItem = "NoMenuSelected"
    End Sub
</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
    <div align="center">
        <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0" class="wrapper">
	       <tr>
	            <td class="body" colspan="2">
                    <table width="100%">
		              <tr class="body_title">
		                <td colspan="1" align="left"><font color="red">Application Resource Error</font><hr /></td>
		              </tr>
		              <tr>
		                <td align="left"><br /><p>The resource you requested is currently unavailable.  Please ensure the URL you entered is correct or you are authorized to access the resource.</p>

		                </td>
		              </tr>
                    </table>
                </td>
	        </tr>      
        </table>
    </div>
</asp:Content>