<%@ Page Language="VB" MasterPageFile="PageMaster.Master" AutoEventWireup="false" EnableEventValidation="false" ASPCOMPAT="TRUE" %>
<%@ Register Assembly="Obout.Ajax.UI" Namespace="Obout.Ajax.UI.Captcha" TagPrefix="obout" %>
<%@ MasterType VirtualPath="PageMaster.Master" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.Page" %>
<%@ Import Namespace="Webapps.Utils" %>

<script language="VB" runat="server">
    '---- Error logging -----
    Dim pageName As String
    Dim strRecepient As String = Webapps.Utils.ApplicationSettings.ErrorNoticeEmails
    Dim strFrom As String = Webapps.Utils.ApplicationSettings.ApplicationSourceEmail
    Dim errLocation As String
    Dim errString As String
    Dim strUserNoticeAddress As String = ApplicationSettings.UserAccountNoticeEmails
    Dim strClientUserNoticeAddress As String = ApplicationSettings.ClientUserAccountNoticeEmails
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        pageName = Request.RawUrl.ToString
        Try
            Master.SetCurrentMenuItem = System.IO.Path.GetFileName(Request.RawUrl.ToString)
        Catch ex As Exception
            Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
            Exit Sub
        End Try

        If CustomRoles.RolesForPageLoad() Then
            
        Else
            CustomRoles.TransferIfNotInRole()
            Exit Sub
        End If
    End Sub
    
    Public Sub ButtonResetPassword_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ButtonResetPassword.Click
        lblMessage.Visible = True
        Dim userName As String = tbUserID.Text
        Dim eMail As String = tbEmail.Text
        Dim bValid As Boolean = True
        Dim sb As StringBuilder = New StringBuilder
        
        If String.IsNullOrEmpty(tbCaptchaIn.Text.Trim()) OrElse CommonUtilsv2.Validate(tbCaptchaIn.Text.Trim(), CommonUtilsv2.DataTypes.String, False, True, False, 20) = False Then
            bValid = False
            ' sb.AppendLine("Image Code must be entered.")
            tbCaptchaIn.BackColor = Drawing.Color.LightPink
        ElseIf Captcha1.TestText(tbCaptchaIn.Text.Trim()) Then
            tbCaptchaIn.BackColor = Drawing.Color.White
        Else
            bValid = False
            '   sb.AppendLine("Image Code entered is invalid.")
            tbCaptchaIn.BackColor = Drawing.Color.LightPink
        End If

        If String.IsNullOrEmpty(tbUserID.Text.Trim()) OrElse CommonUtilsv2.Validate(tbUserID.Text.Trim(), CommonUtilsv2.DataTypes.String, True, True, False, 50) = False Then
            bValid = False
            sb.AppendLine("A valid User ID must be entered.")
            tbUserID.BackColor = Drawing.Color.LightPink
        Else
            tbUserID.BackColor = Drawing.Color.White
        End If
        
        Dim eMailRegex As String = "^[\w-\.]+@([\w-]+\.)+[\w-]{2,3}$"
        If String.IsNullOrEmpty(tbEmail.Text.Trim()) OrElse CommonUtilsv2.Validate(tbEmail.Text.Trim(), CommonUtilsv2.DataTypes.String, True, False, False, 120) = False OrElse (Regex.IsMatch(tbEmail.Text, eMailRegex)) = False Then
            bValid = False
            sb.AppendLine("A valid E-mail Address must be entered.")
            tbEmail.BackColor = Drawing.Color.LightPink
        Else
            tbEmail.BackColor = Drawing.Color.White
        End If
                
        If Not bValid Then
            Dim err As New CustomValidator
            err.ValidationGroup = "vg_ResetPW"
            err.ErrorMessage = sb.ToString
            err.IsValid = False
            Page.Validators.Add(err)
            'lblMessage.ForeColor = Drawing.Color.Red
            'lblMessage.Text = sb.ToString
        Else
            GenerateAndSend(userName, eMail)
        End If
    End Sub
    
    Public Sub GenerateAndSend(ByVal User As String, ByVal Email As String)
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strSQL As String = "Select Email From tbl_usr_Details Where UserID=@UserID And Email=@Email"
        Dim myReader As SqlDataReader
        Dim params As SqlParameter() = { _
            New SqlParameter("@UserID", User), _
            New SqlParameter("@Email", Email) _
            }
        Try
            myReader = CommonUtilsv2.GetDataReader(dbKey, strSQL, CommandType.Text, params)
            If myReader.Read Then
                SendEmailRequest(User, Email)
            Else
                SendEmailRequestNoMatch(User, Email)
            End If
            lblMessage.ForeColor = Drawing.Color.Green
            lblMessage.Text = "Your password request has been sent."
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    
    Public Sub SendEmailRequest(ByVal userName As String, ByVal email As String)
        Dim datetimestamp As String = (DateTime.Now).ToString("MMMM dd, yyyy @ hh:mm")
        Dim strSubject As String = "Password Change Request for UserName: " + userName
        Dim strBody As String = "<font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
        strBody += "<b>A password reset request was made on " + datetimestamp + " for UserName: " + userName + " in web application: " + Webapps.Utils.ApplicationSettings.SiteTitle & " - " & CommonUtilsv2.GetEnvironmentAndHost() & "</b>"
        strBody += "</font>"
        strBody += "<table width='100%'><font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
        strBody += "<tr><td colspan=3><hr /></td></tr>"
        strBody += "<tr><td width='20%' align='right'><u>UserName</u>&nbsp;&nbsp;</td><td width='40%' align='left'><u>E-mail Address</u></td><td width='40%'><u>Web Application</u></td></tr>"
        strBody += "<tr><td align='right'>" + userName + "&nbsp;&nbsp;</td><td align='left'>" + email + "</td><td>" + Webapps.Utils.ApplicationSettings.SiteURL + "</td></tr>"
        strBody += "<tr><td colspan=3><hr /></td></tr>"
        strBody += "</font></table>"
        If String.Compare("NonClientUser", GetUserType(email), True) = 0 Then
         
            strUserNoticeAddress = ApplicationSettings.UserAccountNoticeEmails
            Else
            strUserNoticeAddress = ApplicationSettings.ClientUserAccountNoticeEmails
        End If
        CommonUtilsv2.SendEMail(strFrom, strUserNoticeAddress, strBody, strSubject)
     
    End Sub
    
    Public Sub SendEmailRequestNoMatch(ByVal userName As String, ByVal email As String)
        Dim datetimestamp As String = (DateTime.Now).ToString("MMMM dd, yyyy @ hh:mm")
        Dim strSubject As String = "Password Change Request for UserName: " + userName + " NO MATCH IN DATABASE"
        Dim strBody As String = "<font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
        strBody += "<b>A password reset request was made on " + datetimestamp + " for UserName: " + userName + " in web application: " + Webapps.Utils.ApplicationSettings.SiteTitle & " - " & CommonUtilsv2.GetEnvironmentAndHost() & "</b>"
        strBody += "</font>"
        strBody += "<table width='100%'><font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
        strBody += "<tr><td colspan=3><hr /></td></tr>"
        strBody += "<tr><td width='20%' align='right'><u>UserName</u>&nbsp;&nbsp;</td><td width='40%' align='left'><u>E-mail Address</u></td><td width='40%'><u>Web Application</u></td></tr>"
        strBody += "<tr><td align='right'>" + userName + "&nbsp;&nbsp;</td><td align='left'>" + email + "</td><td>" + Webapps.Utils.ApplicationSettings.SiteTitle + "</td></tr>"
        strBody += "<tr><td align='left' colspan='3'>Username and Email did not produce a match in the database."
        strBody += "<tr><td colspan=3><hr /></td></tr>"
        strBody += "</font></table>"
        CommonUtilsv2.SendEMail(strFrom, strUserNoticeAddress, strBody, strSubject)
    End Sub
    
     
    Function GetUserType(ByVal userEmail As String) As String

        Dim strRet As String = "NonClientUser"
    
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")

        Dim sql As String = "SELECT UserID, UserType, Email FROM [v_w_UserTypeList] WHERE Email = @Email "

        Dim params As SqlParameter() = { _
            New SqlParameter("@Email", userEmail)
            }

        Dim myReader As SqlDataReader = Nothing
        Try
            myReader = CommonUtilsv2.GetDataReader(dbKey, sql, CommandType.Text, params)

            If Not myReader Is Nothing Then

                While myReader.Read()

                    If myReader.GetValue(1) Is System.DBNull.Value Then
                       
                    Else
                        strRet = myReader.GetString(1)
                    End If
                End While

            End If

        Catch ex As Exception
            strRet = "NonClientUser"
            errString = ex.Message
            errLocation = "GetUserType()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myReader Is Nothing Then
                myReader.Close()
            End If
        End Try
        Return strRet
        
    End Function
    
    
</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
    <script type="text/javascript">
        function backfunction() {
            window.location.href = "/login.aspx";
        }
    </script>
    <div align="center" style="height: 100%">
          <div class="body_title">
            <a href="javascript:void(0);" onclick="backfunction()"><span>back</span></a>
            Request Password Reset
        </div>
            <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0" class="wrapper">
            <%--<tr >
            <td align="left" class="body_title">Request Password Reset</td>
		    </tr>--%>
             <tr>      
                <td width="100%" align="left" valign="top" class="form_login">
                    <table width="100%" border="0">
		              <tr>
		                <td align="left" width="100%">
		                  <div>
                                <table style="width: 100%" border="0" class="reg_table_style">
                                    <tr>
                                        <td align="left" colspan="2">
                                        <asp:Label runat="server" ID="lblMessage" Visible="true" ForeColor="Red" Text="Please submit your request for password reset below. The site administrator will review and complete your request promptly." />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="2">&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td width="25%" align="right">
                                            <asp:Label ID="lblUserID" runat="server" Text="User ID:"></asp:Label>
                                        </td>    
                                        <td width="75%" align="left">
                                            <asp:TextBox ID="tbUserID" runat="server" Width="300px" MaxLength="50"></asp:TextBox>
                                            <asp:RequiredFieldValidator runat="server" ID="RequiredFieldValidator2" ControlToValidate="tbUserID" ErrorMessage="User ID is required." Text="*" ValidationGroup="vg_ResetPW" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right">
                                            <asp:Label ID="lblEmail" runat="server" Text="E-mail Address:"></asp:Label>
                                        </td>    
                                        <td align="left">
                                            <asp:TextBox ID="tbEmail" runat="server" Width="300px" MaxLength="100"></asp:TextBox>
                                            <asp:RequiredFieldValidator runat="server" ID="RequiredFieldValidator1" ControlToValidate="tbEmail" ErrorMessage="Email address is required." Text="*" ValidationGroup="vg_ResetPW" />
                                        </td>
                                    </tr>

                                    <tr>
                                       <td align="right">
                                             <a href="javascript: $find('<%= Captcha1.ClientID %>').getNewImage();">Refresh Image</a>
                                        </td>
                                        <td align="left">
                                            <obout:CaptchaImage ID="Captcha1" runat="server" FontWarpLevel="Low" > </obout:CaptchaImage>
                                        </td>    
                                    </tr>

                                    <tr>
                                        <td align="right">
                                            <asp:Label ID="Label1" runat="server" Text="Enter Image Code:"></asp:Label>
                                        </td>    
                                        <td align="left">
                                            <asp:TextBox ID="tbCaptchaIn" runat="server" Width="300px" MaxLength="100"></asp:TextBox>
                                            <asp:RequiredFieldValidator runat="server" ID="RequiredFieldValidator8" ControlToValidate="tbCaptchaIn" ErrorMessage="Image Code is required." Text="*" ValidationGroup="vg_ResetPW" />
                                            <obout:CaptchaValidator ID="CaptchaValidator1" runat="server" ValidationGroup="vg_ResetPW" ControlToValidate="tbCaptchaIn" CaptchaImageID="Captcha1"
                                                ErrorMessage="The image code you entered is invalid." Display="Dynamic"></obout:CaptchaValidator>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="right">
                                        </td>     
		                                <td align="left">
                                        <asp:Button ID="ButtonResetPassword" runat="server" Text="Submit" Width="154px" CausesValidation="true"  ValidationGroup="vg_ResetPW" CssClass="Submit_button"/>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center" valign="top" colspan="2">
                                            <asp:ValidationSummary ID="ValidationSummary" runat="server" Width="500px"
                                                HeaderText="Errors:" ShowSummary="true" DisplayMode="List" BackColor="#f2dede" ValidationGroup="vg_ResetPW" />
                                        </td>
                                    </tr>
                                </table>   
                          </div>
                        </td>                   
		              </tr>
		              <tr>
		              <td align="left">
		              </td>
		              </tr>
                    </table>
                </td>            
            </tr> 
        </table>       
    </div>
</asp:Content>
