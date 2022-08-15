<%@ Page Language="VB" MasterPageFile="PageMaster.master" AutoEventWireup="false" EnableEventValidation="false" ASPCOMPAT="TRUE" %>
<%@ MasterType VirtualPath="PageMaster.master" %>
<%@ Register TagPrefix="obout" Namespace="Obout.Grid" Assembly="obout_Grid_NET" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.Page" %>
<%@ Import Namespace="Webapps.Utils" %>

<script language="VB" runat="server">
    '************* Error logging Section ***********************
    Dim pageName As String
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        pageName = Request.RawUrl.ToString  'initialize the page name
        Master.SetCurrentMenuItem = "NoMenuSelected"
        Dim strLogout As String = "False"
        strLogout = Request.QueryString.Get("logout")
        If strLogout = "True" Then
            Session("Level") = "0"
        End If
        Master.SetCurrentMenuItem = System.IO.Path.GetFileName(Request.RawUrl.ToString)
        If CustomRoles.RolesForPageLoad() Then

        Else
            CustomRoles.TransferIfNotInRole()
            Exit Sub
        End If
    End Sub
</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
    <div align="center">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
		                <tr class="body_title">
		                    <td colspan="7" align="left" class="body_title">Submission Successful</td>
		                </tr>
		                <tr >
	                        <td class="body_plain" align="left" colspan="2">
	                            <br /><br />
	                            <b>Thank you.  Your information has been submitted, and will be reviewed as soon as possible.</b>
	                        </td>
			            </tr>
			            <tr >
			              <td class="body_simple" align="left" colspan="2">
			                <br /><br />
			                &nbsp;&nbsp;Click <a href="page.aspx">here</a> to return to the home page of this site.
			              </td>
			            </tr>      
           </table> 
        </div>
</asp:Content>