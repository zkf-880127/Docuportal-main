<%@ Page Language="VB" MasterPageFile="PageMaster.Master" AutoEventWireup="false" EnableEventValidation="false" viewStateEncryptionMode="Auto" ASPCOMPAT="TRUE" %>
<%@ MasterType VirtualPath="PageMaster.Master" %>
<%@ Register Assembly="Obout.Ajax.UI" Namespace="Obout.Ajax.UI.Captcha" TagPrefix="obout" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.Page" %>
<%@ Import Namespace="Webapps.Utils" %>
<%@ Import Namespace="BCrypt" %>
<%@ Import Namespace="ExpertPdf.HtmlToPdf" %>


<script language="VB" runat="server">

    '---- login globals
    Dim bSendLoginEmail As Boolean = False

    '---- Error logging
    Dim pageName As String = "login.aspx"
    Dim strRecepient As String = Webapps.Utils.ApplicationSettings.ErrorNoticeEmails
    Dim strFrom As String = Webapps.Utils.ApplicationSettings.ApplicationSourceEmail
    Dim errLocation As String
    Dim errString As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim strURL As String = Webapps.Utils.ApplicationSettings.SiteURL

        pageName = Request.RawUrl.ToString
        Try
            Master.SetCurrentMenuItem = System.IO.Path.GetFileName(Request.RawUrl.ToString)
        Catch ex As Exception
            Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
            Exit Sub
        End Try
        If Not IsPostBack Then
            'This block to clear of session even if user gets here from browse back button without logout
            Session("Level") = 0
            Session("FullName") = ""
            Session("Display") = ""
            Session("User") = ""
            Session("UserType") = ""
            Session("VerifcationCaseID") = ""
            Session("CaseID_For_CaseUser") = ""
            Dim strStoredRedirectURL = Session("RedirectURLAfterLogin")
            Session.Clear()
            Session("RedirectURLAfterLogin") = strStoredRedirectURL
        End If        'Response.Cache.SetExpires(DateTime.Parse(DateTime.Now.ToString()))
        'Response.Cache.SetCacheability(HttpCacheability.Private)
        'Response.Cache.SetCacheability(HttpCacheability.NoCache)
        'Response.Cache.SetNoStore()

        Response.AppendHeader("Pragma", "no-cache")
        Response.Cache.SetNoStore()
        Response.Cache.AppendCacheExtension("no-cache")
        Response.Expires = 0
        Page.MaintainScrollPositionOnPostBack = False
        Page.SetFocus(tbUsername)

    End Sub

    Protected Sub btn_Submit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '---- clear session (removing any straggler session variables)
        'Session.Clear()
        Dim bValid As Boolean = True
        Dim sb1 As StringBuilder = New StringBuilder

        If String.IsNullOrEmpty(tbCaptchaIn.Text.Trim()) OrElse CommonUtilsv2.Validate(tbCaptchaIn.Text.Trim(), CommonUtilsv2.DataTypes.String, False, True, False, 20) = False Then
            bValid = False
            ' sb1.AppendLine("Image Code must be entered.")
            tbCaptchaIn.BackColor = Drawing.Color.LightPink
        ElseIf Captcha1.TestText(tbCaptchaIn.Text.Trim()) Then
            tbCaptchaIn.BackColor = Drawing.Color.White
        Else
            bValid = False
            '   sb1.AppendLine("Image Code entered is invalid.")
            tbCaptchaIn.BackColor = Drawing.Color.LightPink
        End If

        If String.IsNullOrEmpty(tbUsername.Text.Trim()) OrElse CommonUtilsv2.Validate(tbUsername.Text.Trim(), CommonUtilsv2.DataTypes.String, True, False, False, 50) = False Then
            bValid = False
            sb1.AppendLine("Please enter a valid username. <br/><br/>")
            tbUsername.BackColor = Drawing.Color.LightPink
        Else
            tbUsername.BackColor = Drawing.Color.White
        End If
        If String.IsNullOrEmpty(tbPassword.Text.Trim()) OrElse CommonUtilsv2.Validate(tbPassword.Text.Trim(), CommonUtilsv2.DataTypes.String, True, False, False, 50) = False Then
            bValid = False
            sb1.AppendLine("Please enter a valid password. <br/><br/>")
            tbPassword.BackColor = Drawing.Color.LightPink
        Else
            tbPassword.BackColor = Drawing.Color.White
        End If
        Dim strStoredRedirectURL = Session("RedirectURLAfterLogin")
        Session.Clear()
        Session("RedirectURLAfterLogin") = strStoredRedirectURL

        If Not bValid Then
            Dim err As New CustomValidator
            err.ValidationGroup = "vg_Login"
            err.ErrorMessage = sb1.ToString
            err.IsValid = False
            Page.Validators.Add(err)
            Exit Sub
        End If


        Dim sb As StringBuilder = New StringBuilder()
        '----Validate the input
        Dim userID As String = tbUsername.Text.Trim()
        Dim pass As String = tbPassword.Text.Trim()
        Dim bRequirePWChange As Boolean = False
        '---- Create instance of login class
        Dim processLogin As login = New login
        Dim iRet As Integer = processLogin.Login(userID, pass, Request.UserHostAddress())

        Select Case iRet
            Case 0
                ' unknown such as server error.  Ask user to retry.  Login page
                sb.Append(" An error has occured with this login, and a message has been sent to the site administrator. ")
                'lbl_loginErrors.Text = sb.ToString
                'lbl_loginErrors.Visible = True
                SendEmail("Unknow error", processLogin.NumFailedLogins)
                ' Exit Sub
            Case 1
                'invalid ID.  Login page                                         
                sb.Append(" The combined id/pw is invalid. Try again. ")
                'lbl_loginErrors.Text = sb.ToString
                'lbl_loginErrors.Visible = True
                ' Exit Sub
            Case 2
                'invalid pw Login page 
                sb.Append(" The combined id/pw is invalid. Try again. ")
                'lbl_loginErrors.Text = sb.ToString
                'lbl_loginErrors.Visible = True
                SendEmail("2: Invalid password", processLogin.NumFailedLogins)
                ' Exit Sub
            Case 3
                'Account disabled.  Contact administrator
                sb.Append(" Your account is not active.  Please contact Administrator. ")
                'lbl_loginErrors.Text = sb.ToString
                'lbl_loginErrors.Visible = True
                SendEmail("3: Account disabled", processLogin.NumFailedLogins)
                Session("User") = ""
                btn_Submit.Enabled = False
                tbUsername.Enabled = False
                tbPassword.Enabled = False
                '   Exit Sub
            Case 4
                'Account Locked.  Password Reset request page
                sb.Append(" Your account is locked. Please <a href='forgotpassword.aspx'>click here</a> to submit a request to unlock your account.")
                'lbl_loginErrors.Text = sb.ToString
                'lbl_loginErrors.Visible = True
                SendEmail("4: Account Locked", processLogin.NumFailedLogins)
                Session("User") = ""
                btn_Submit.Enabled = False
                tbUsername.Enabled = False
                tbPassword.Enabled = False
                ' Exit Sub
            Case 5
                'Required to change pw.  Change Password page
                sb.Append("You must change your password before proceeding. Please <a href='ChangePassword.aspx'>click here</a> to change your password.")
                'lbl_loginErrors.Text = sb.ToString
                'lbl_loginErrors.Visible = True
                Session("@@UserIDForChangePW") = userID 'Session("User")
                Session("User") = ""
                Session("FullName") = processLogin.FullName
                btn_Submit.Enabled = False
                tbUsername.Enabled = False
                tbPassword.Enabled = False
                '  Exit Sub
            Case 6
                'Password expired and cannot change.  Password Reset request page
                sb.Append(" Your account is expired. Please <a href='forgotpassword.aspx'>click here</a> to submit a request to reactivate your account.")
                'lbl_loginErrors.Text = sb.ToString
                'lbl_loginErrors.Visible = True
                Session("User") = ""
                Session("FullName") = ""
                btn_Submit.Enabled = False
                tbUsername.Enabled = False
                tbPassword.Enabled = False
                '   Exit Sub
            Case 9
                'Login success
                Session("User") = tbUsername.Text.Trim()
                Session("FullName") = processLogin.FullName

                Dim strLandingPage As String = ""

                If String.IsNullOrEmpty(strLandingPage) Then
                    strLandingPage = Webapps.Utils.ApplicationSettings.Homepage
                End If
                If Not Session("RedirectURLAfterLogin") Is Nothing Then
                    strLandingPage = Session("RedirectURLAfterLogin")
                    If strLandingPage.Contains("DownloadClaimADR") Then
                        Session("DownloadADRPageRedirect") = True
                    Else
                        Session("DownloadADRPageRedirect") = False
                    End If
                    Session.Remove("RedirectURLAfterLogin")
                End If
                If Session("DownloadADRPageRedirect") = True Then
                    Session("RedirectDownloadURL") = strLandingPage
                    Dim iDocID As Integer = Integer.Parse(strLandingPage.Split("=")(1))
                    Response.Redirect(CommonUtilsv2.GetADRPage(iDocID))
                Else
                    Response.Redirect(strLandingPage, False)
                End If

                Response.Redirect(strLandingPage, False)

                'Dim strLandingPage As String = Webapps.Utils.ApplicationSettings.Homepage
                'If String.IsNullOrEmpty(strLandingPage) Then
                '    strLandingPage = System.Configuration.ConfigurationManager.AppSettings("Homepage")
                'End If
                'Response.Redirect(strLandingPage, False)
        End Select

        If iRet < 9 Then
            Dim err As New CustomValidator
            err.ValidationGroup = "vg_Login"
            err.ErrorMessage = sb.ToString
            err.IsValid = False
            Page.Validators.Add(err)
        End If
    End Sub

    Private Function GetLandingPage() As String
        Dim strRet As String = ""
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim strSQL As String = " SELECT A.KeyDescription FROM tbl_ROLES_UserRoles UR INNER JOIN tbl_Apps_AppSettings A ON UR.Role_ID=A.KeyValue WHERE UR.User_ID=@ID "
        Dim myreader As SqlDataReader = Nothing
        Dim params As SqlParameter() = {New SqlParameter("@ID", Session("User"))}
        Try
            myreader = CommonUtilsv2.GetDataReader(dbKey, strSQL, CommandType.Text, params)
            While myreader.Read
                strRet = If(IsDBNull(myreader(0)), "", myreader(0).ToString)
            End While
        Catch ex As Exception
            errString = ex.Message()
            errLocation = "GetLandingPage "
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myreader Is Nothing AndAlso Not myreader.IsClosed Then
                myreader.Close()
            End If
        End Try

        Return strRet
    End Function

    Private Sub SendEmail(ByVal loginResult As String, ByVal iFailedAttmpts As Integer)

        Dim strUserID As String = tbUsername.Text.Trim()
        If String.IsNullOrEmpty(strUserID) Then strUserID = Request.UserHostAddress()
        Dim strSubject As String = ""
        Dim strBody As String = ""
        Dim rightNow As DateTime = DateTime.Now
        Dim strDateTime As String
        strDateTime = rightNow.ToString("yyyyMMdd_hhmmss")
        Dim strFullDateTime As String
        strFullDateTime = rightNow.ToString("MMMM dd, yyyy @ hh:mm:ss")
        strSubject = Webapps.Utils.ApplicationSettings.SiteTitle & " - " & CommonUtilsv2.GetEnvironmentAndHost() & " - Login Code: " + loginResult + " for " + strUserID
        strBody = "<b>" + strUserID + "</b> had a login problem on "
        strBody = strBody + strFullDateTime + "<br /><br />"
        strBody = strBody + "<hr />"
        strBody = strBody + "<table>"
        strBody = strBody + "<tr><td>UserID: </td><td><b>" + strUserID + "</b></td></tr>"
        strBody = strBody + "<tr><td>LoginResult: </td><td><b>" + loginResult + "</b></td></tr>"
        strBody = strBody + "<tr><td>LoginError: </td><td><b>" + lbl_loginErrors.Text + "</b></td></tr>"
        strBody = strBody + "<tr><td>Failed Attempts: </td><td><b>" + iFailedAttmpts.ToString + "</b></td></tr>"
        strBody = strBody + "</table><hr />"
        strBody = strBody + "<b>NOTE: Please do not reply to this email - this is a notification email only.</b>"
        CommonUtilsv2.SendEMail(strFrom, strRecepient, strBody, strSubject)

    End Sub

    Function GetUserMail(ByVal strUserID As String) As String
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
        Dim eMailAddress As String = ""
        Dim strSQL As String = "Select eMail, ID from tbl_usr_Details WHERE UserID = @UserID "
        Dim myComm As New SqlCommand(strSQL, myConn)
        myComm.Parameters.AddWithValue("@UserID", strUserID)
        Dim myReader As SqlDataReader = Nothing
        Try
            myConn.Open()
            myReader = myComm.ExecuteReader()
            While myReader.Read()
                If myReader.GetValue(0) Is System.DBNull.Value Then eMailAddress += "" Else eMailAddress += myReader(0).ToString
                If Not myReader.GetValue(1) Is System.DBNull.Value Then Session("User_RowID") = myReader.GetValue(1)
            End While
            myReader.Close()
        Catch ex As Exception
            errString = ex.Message
            errLocation = "GetUserEMail()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, tbUsername.Text.Trim(), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myreader Is Nothing AndAlso Not myreader.IsClosed Then
                myreader.Close()
            End If
            myComm.Dispose()
            myConn.Close()
        End Try
        Return eMailAddress
    End Function


    Private Sub UpdatePermPassword()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strSQL As String = "Update tbl_usr_Logins Set PermPasswordExpire=@PermPasswordExpire Where Userid=@Userid"
        Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
        Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
        Dim permDate As Date = Now.AddDays(60)
        sqlCmd.CommandType = CommandType.Text
        sqlCmd.Parameters.AddWithValue("@Userid", tbUsername.Text.Trim())
        sqlCmd.Parameters.AddWithValue("@PermPasswordExpire", permDate)
        Try
            sqlCnn.Open()
            sqlCmd.ExecuteNonQuery()
        Catch ex As Exception
            errString = ex.Message
            errLocation = "Update Perm PW Expire"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, tbUsername.Text.Trim(), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            sqlCmd.Dispose()
            sqlCnn.Close()
            sqlCnn.Dispose()
        End Try
    End Sub

 </script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
    <div align="center" style="height: 100%">
        <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0" class="reg_table_style" style="padding-left:10px;" >
            <tr><td>&nbsp;</td></tr>
            <tr >
                <td align="left" class="body_title"><%--Secure Login--%></td>
		    </tr>
		    <tr>
		        <td align="left" width="100%">
		            <div>
                        <table style="width: 100%" border="0">
                            <tr style="padding-top:5px; padding-bottom:5px;">
                                <td align="left" colspan="2">
                                    <asp:Label runat="server" ID="lbl_loginErrors" Visible="false" ForeColor="Red" />
                                </td>
                            </tr>
                            <tr style="padding-top:5px; padding-bottom:5px; height:2rem;">
                                <td colspan="2">&nbsp;</td>
                            </tr>
                            <tr style="padding-top:5px; padding-bottom:5px; height:2rem;">
                                <td width="25%" align="right">Username:&nbsp;</td>
                                <td width="75%" align="left"><asp:TextBox ID="tbUsername" runat="server" MaxLength="50" Width="300px" autocomplete="off" TabIndex="1" Height="26px"></asp:TextBox></td> 
                                <asp:RequiredFieldValidator runat="server" ID="RequiredFieldValidator1" ControlToValidate="tbUsername" ErrorMessage="Username is required." Text="*" ValidationGroup="vg_Login" />        
                            </tr>
                            <tr style="padding-top:5px; padding-bottom:5px; height:2rem;">
                                <td align="right">Password:&nbsp;</td>
                                <td><asp:TextBox ID="tbPassword" runat="server" Width="300px" MaxLength="50" TextMode="Password" autocomplete="off" TabIndex="2" Height="26px"></asp:TextBox>&nbsp;&nbsp;<a href="ForgotPassword.aspx" tabindex="6">Forgot Password?</a>&nbsp;&nbsp;</td>    
                                <asp:RequiredFieldValidator runat="server" ID="RequiredFieldValidator2" ControlToValidate="tbPassword" ErrorMessage="Password is required." Text="*" ValidationGroup="vg_Login" />
                            </tr>
                            <tr style="padding-top:5px; padding-bottom:5px; height:2rem;">
                                <td align="right">
                                        <a href="javascript: $find('<%= Captcha1.ClientID %>').getNewImage();">Refresh Image</a>
                                </td>
                                <td align="left">
                                    <obout:CaptchaImage ID="Captcha1" runat="server" FontWarpLevel="Low" > </obout:CaptchaImage>
                                </td>    
                            </tr>
                            <tr style="padding-top:5px; padding-bottom:5px; height:2rem;">
                                <td align="right">
                                    <asp:Label ID="Label1" runat="server" Text="Enter Image Code:"></asp:Label>
                                </td>    
                                <td align="left">
                                    <asp:TextBox ID="tbCaptchaIn" runat="server" Width="300px" MaxLength="100" TabIndex="3" Height="26px"></asp:TextBox>
                                    <asp:RequiredFieldValidator runat="server" ID="RequiredFieldValidator8" ControlToValidate="tbCaptchaIn" ErrorMessage="Image Code is required." Text="*" ValidationGroup="vg_Login" />
                                    <obout:CaptchaValidator ID="CaptchaValidator1" runat="server" ValidationGroup="vg_Login" ControlToValidate="tbCaptchaIn" CaptchaImageID="Captcha1"
                                        ErrorMessage="The image code you entered is invalid." Display="Dynamic"></obout:CaptchaValidator>
                                </td>
                            </tr>                               
                            <tr><td>&nbsp;</td></tr>
                            <tr style="padding-top:5px; padding-bottom:5px; height:2rem;">
                                <td align="right">&nbsp;</td>
                                <td align="left"><asp:Button ID="btn_Submit" runat="server" OnClick="btn_Submit_Click" Text="Submit" TabIndex="4" Width="154px" CausesValidation="true"  ValidationGroup="vg_Login" CssClass="Submit_button"/></td>
                            </tr>
                            <tr style="padding-top:5px; padding-bottom:5px; height:2rem;">
                                <td align="center" valign="top" colspan="2">
                                    <asp:ValidationSummary ID="ValidationSummary" runat="server" Width="90%"
                                        HeaderText="Errors:" ShowSummary="true" DisplayMode="List" BackColor="#f2dede" ValidationGroup="vg_Login" />
                                </td>
                            </tr>
                            <tr><td colspan="2">&nbsp;</td></tr>
                            <tr><td colspan="2">&nbsp;</td></tr>
                            <tr><td colspan="2">&nbsp;</td></tr>
                            <tr>
                                <td colspan="2" class="body_disclaimer">
                                    <%--<b>NOTICE:</b>This computer system belongs to WebApps.
                                    &nbsp;&nbsp;Access to this private computer system is for authorized users only. 
                                    &nbsp;&nbsp;Unauthorized and/or inappropriate use, including exceeding authorization, is strictly prohibited and may subject you to civil and criminal penalties. 
                                    &nbsp;&nbsp;By using this computer system, you understand and consent to the following: 
                                    &nbsp;&nbsp;(i) you have no reasonable expectation of privacy regarding communications or data transiting or stored on this computer system; 
                                    &nbsp;&nbsp;(ii) at any time, and for any lawful purpose, the computer system may be monitored and recorded and any communication or data transiting or stored on this computer system may be intercepted, searched, and seized; 
                                    &nbsp;&nbsp;(iii) and any communication or data transiting or stored on this computer system may be disclosed or used for any lawful purpose.  
                                    &nbsp;&nbsp;In addition, WebApps reserves the right to consent to a valid law enforcement request to search the computer system.
                                    <br /><br />
                                    By logging on, you hereby agree to the foregoing terms and conditions.  If you do not agree to the foregoing terms and condition, do not log on and do not use this computer system.--%>
					            </td>
                            </tr>
                        </table>
                    </div>
                </td>                   
		    </tr> 
        </table>       
    </div>
     <script type="text/javascript">
         scrollTo(0, 0); 
       //  document.getElementById('home1').scrollIntoView(true);
     </script>

</asp:Content>
