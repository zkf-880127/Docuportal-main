<%@ Page Language="VB" MasterPageFile="PageMaster.Master" AutoEventWireup="false" EnableEventValidation="false" viewStateEncryptionMode="Auto" ASPCOMPAT="TRUE" %>
<%@ MasterType VirtualPath="PageMaster.Master" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.Page" %>
<%@ Import Namespace="Webapps.Utils" %>

<script language="VB" runat="server">
    Dim bRequirePWChange As Boolean = False
    Dim iNumberOfPriorPWNotbeUsed As Integer = ApplicationSettings.NumberOfPriorPWsNotBeUsed 'CommonUtilsv2.GetNumberOfPriorPWsNotBeUsed()
    
    Dim strTempMsg As String = ""
    
    '---- Error logging
    Dim pageName As String = "ChangePassword.aspx"
    Dim strRecepient As String = Webapps.Utils.ApplicationSettings.ErrorNoticeEmails
    Dim strFrom As String = Webapps.Utils.ApplicationSettings.ApplicationSourceEmail
    Dim environment As String = Webapps.Utils.ApplicationSettings.Environment
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
        If Not String.IsNullOrEmpty(Session("@@UserIDForChangePW")) Then
            bRequirePWChange = True
        End If

        If bRequirePWChange Then
            If String.Compare("Public", Session("User"), True) <> 0 And Not String.IsNullOrEmpty(Session("@@UserIDForChangePW")) Then
                If iNumberOfPriorPWNotbeUsed > 0 Then
                    If Not IsPostBack() Then
                        lblMessage.Text = lblMessage.Text & "</br>  Password cannot be one of the " & iNumberOfPriorPWNotbeUsed.ToString & " previous passwords."
                    End If
                End If
            Else
                Response.Redirect("Login.aspx", True)
            End If
        ElseIf String.Compare("Public", Session("User"), True) <> 0 And Not String.IsNullOrEmpty(Session("User")) Then
            If iNumberOfPriorPWNotbeUsed > 0 Then
                If Not IsPostBack() Then
                    lblMessage.Text = lblMessage.Text & "</br>  Password cannot be one of the " & iNumberOfPriorPWNotbeUsed.ToString & " previous passwords."
                End If
            End If
        Else
            Response.Redirect("Login.aspx", True)
        End If
    End Sub
    
    Public Sub ButtonResetPassword_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ButtonChangePassword.Click
        lblMessage.Visible = True
        Dim userName As String = Session("User")
        If bRequirePWChange Then
            userName = Session("@@UserIDForChangePW")
        End If
        Dim oldPassword As String = tbCurrentPassword.Text.Trim()
        Dim newPassword As String = tbNewPassword.Text.Trim()
        
        If String.IsNullOrEmpty(newPassword) Or CommonUtilsv2.Validate(newPassword, CommonUtilsv2.DataTypes.String, True, False, False) = False Then
            lblMessage.Text = "Password entered is invalid."
            lblMessage.ForeColor = Drawing.Color.Red
            Exit Sub
        End If
        
        If String.IsNullOrEmpty(userName) Then
            lblMessage.Text = "Invalid username. Please log into the website again."
            lblMessage.ForeColor = Drawing.Color.Red
            Exit Sub
        End If

        If String.Compare(tbNewPassword.Text, tbCurrentPassword.Text, True) = 0 Then
            lblMessage.Text = "New password cannot be the same as the current password. "
            lblMessage.ForeColor = Drawing.Color.Red
            Exit Sub
        End If
 
        If Not String.Compare(newPassword, tbConfirmPassword.Text.Trim(), True) = 0 Then
            lblMessage.Text = "Confirm password must be the same."
            lblMessage.ForeColor = Drawing.Color.Red
            Exit Sub
        End If
        '---- Create instance of login class
        Dim processLogin As login = New login
        Dim iRet As Integer = processLogin.ChangePassword(userName, newPassword, oldPassword)
        Dim sb As StringBuilder = New StringBuilder()
        
        Select Case iRet
            Case 0
                ' unknown such as server error.  Ask user to retry.  Login page
                sb.Append(" An error has occured while attempt to change password, and administrator has been notified. ")
                lblMessage.Text = sb.ToString
                lblMessage.Visible = True
                SendEmail()
                Exit Sub
            Case 7
                'invalid new pw.  Login page  
                If iNumberOfPriorPWNotbeUsed > 0 Then
                    sb.Append(" New password cannot be the same as one of the recent " & iNumberOfPriorPWNotbeUsed.ToString & " previous passwords. ")
                Else
                    sb.Append(" The new password entered is invalid. ")
                End If
                lblMessage.Text = sb.ToString
                lblMessage.Visible = True
                Exit Sub
            Case 9
                'pw change success
                If String.IsNullOrEmpty(Session("User")) Then
                    Session("User") = Session("@@UserIDForChangePW")
                End If

                Session("@@UserIDForChangePW") = ""
                Session("@@FullNameForChangePW") = ""
                Session("FullName") = processLogin.FullName
                Session("Roles") = Nothing
                lblMessage.Text = "Password has been changed successfully. Your new password will expire in " & ApplicationSettings.DaysPermanentPWExipres.ToString() & " days. Please <a href='" & Webapps.Utils.ApplicationSettings.Homepage & "'>click here</a> to continue.. "
                lblMessage.ForeColor = Drawing.Color.Green
                ButtonChangePassword.Enabled = False
                tbCurrentPassword.Enabled = False
                tbNewPassword.Enabled = False
                tbConfirmPassword.Enabled = False
            Case Else
                ' invalid id/pw entered              
                sb.Append(" Invalid Username or/and current password are entered. ")
                lblMessage.Text = sb.ToString
                lblMessage.Visible = True
                Exit Sub
        End Select
        
    End Sub
    
    Private Sub SendEmail()
        Dim userid As String = Session("User")
        If bRequirePWChange Then
            userid = Session("@@UserIDForChangePW")
        End If
        
        Dim datetimestamp As String = (DateTime.Now).ToString("MMMM dd, yyyy @ hh:mm")
        Dim strSubj As String = " Password change unsuccessful for web site: " + Webapps.Utils.ApplicationSettings.SiteTitle
        If String.Compare("", Webapps.Utils.ApplicationSettings.Environment, True) = 0 Then
        Else
            strSubj = strSubj + " in " + CommonUtilsv2.GetEnvironmentAndHost()
        End If
        Dim strBody As String = "<font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
        strBody += "<b> Attempt to change password on " + datetimestamp + " by User " + userid + " was not successful " + "</b>"
        strBody += "<br/><br/><b> Invalid Username or/and current password are entered. " + "</b>"
        strBody += "</font><br/><br/>"
        CommonUtilsv2.SendEMail(strFrom, strRecepient, strBody, strSubj)

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
    <div align="center" style="height: 100%">
            <div class="body_title">
                <a href="javascript:void(0);" onclick="backfunction()"><span>back</span></a>
                Change Password
            </div>
            <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0" class="wrapper">
             <tr>      
                <td width="100%" align="left" valign="top" class="form_login">
                    <table width="100%" border="0">
		              <tr>
		                <td align="left" width="100%">
		                  <div>
                          <fieldset style="width:96%; height:100%; margin-left:10px;">
                                <legend class="form_login" style="border-style:none">
                                    &nbsp;&nbsp;<b>Change Password</b>&nbsp&nbsp
                                </legend>
                                <table style="width: 100%" border="0">
                                    <tr>
                                        <td align="left" colspan="2">
                                        <asp:Label runat="server" ID="lblMessage" Visible="true" ForeColor="Red" Text="Password must contain a minimum of 8 characters with at least one lower case letter, one upper case letter, and one digit." />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="2">&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td align="right" width="25%">
		                                    Current Password:
		                                </td>     
		                                <td align="left" width="75%">    
		                                    <asp:TextBox ID="tbCurrentPassword" runat="server" TextMode="Password" MaxLength="50"
                                                ValidationGroup="change_password" Width="250px"  ></asp:TextBox>
                                            <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" 
                                                ControlToValidate="tbCurrentPassword" Display="Dynamic" 
                                                ErrorMessage="You must enter a current password." 
                                                ValidationGroup="change_password">*</asp:RequiredFieldValidator>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right">
		                                    New Password:
		                                </td>     
		                                <td align="left">    
		                                    <asp:TextBox ID="tbNewPassword" runat="server" TextMode="Password" MaxLength="50"
                                                ValidationGroup="change_password" Width="250px" ></asp:TextBox>
                                            <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" 
                                                ControlToValidate="tbNewPassword" Display="Dynamic" 
                                                ErrorMessage="You must enter a new password." 
                                                ValidationGroup="change_password">*</asp:RequiredFieldValidator>
                                            <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" 
                                                ControlToValidate="tbNewPassword" Display="Dynamic"                                               
                                                ErrorMessage="Password must contain a minimum of 8 characters with at least one lower case letter, one upper case letter, and one digit." 
                                                ValidationExpression="^.*(?=.{8,})(?=.*\d)(?=.*[a-z])(?=.*[A-Z]).*$"                                                
                                                ValidationGroup="change_password">*</asp:RegularExpressionValidator>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right">
                                            Confirm Password: 
		                                </td>     
		                                <td align="left">
		                                    <asp:TextBox ID="tbConfirmPassword" runat="server" TextMode="Password" Width="250px" MaxLength="50" ValidationGroup="change_password"></asp:TextBox>
                                             <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" 
                                                ControlToValidate="tbConfirmPassword" Display="Dynamic" 
                                                ErrorMessage="You must enter a confirm password." 
                                                ValidationGroup="change_password">*</asp:RequiredFieldValidator>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right">
                                        </td>     
		                                <td align="left">
                                        <asp:Button ID="ButtonChangePassword" runat="server" Text="Change Password" Width="154px" ValidationGroup="change_password" />
                                        </td>
                                    </tr>
                                </table>
                           </fieldset>     
                          </div>
                        </td>                   
		              </tr>
		              <tr>
		              <td align="left">
		              </td>
		              </tr>
		              <tr>
		              <td>
		              <br />
		              
		              </td>
		              </tr>
                    </table>
                </td>            
            </tr> 
        </table>       
    </div>
</asp:Content>
