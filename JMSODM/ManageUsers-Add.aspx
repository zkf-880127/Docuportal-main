<%@ Page Language="VB" MasterPageFile="PageMaster.Master" AutoEventWireup="false" EnableEventValidation="false" viewStateEncryptionMode="Auto" ASPCOMPAT="TRUE" Inherits="OboutInc.oboutAJAXPage" %>
<%@ MasterType VirtualPath="PageMaster.Master" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.Page" %>
<%@ Import Namespace="Webapps.Utils" %>

<script language="VB" runat="server">
    '---- Error logging -----

    Dim errLocation As String
    Dim errString As String
    Dim strHoursTempPasswordExpire As String
    Dim strDaysPermPasswordExpire As String
    Dim hash As String = ""
    Dim pageName As String

    Function HashValue(ByVal strValue As String) As String
        If strValue <> Nothing Or strValue <> "" Then
            hash = BCrypt.Net.BCrypt.HashPassword(strValue, 12)
        Else
            hash = Nothing
        End If
        Return hash
    End Function

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
        lblUserCreated.Text = ""
        strHoursTempPasswordExpire = ApplicationSettings.HoursTempPWExpires 'CommonUtilsv2.GetHoursTempPWExpires()
        strDaysPermPasswordExpire = ApplicationSettings.DaysPermanentPWExipres ' CommonUtilsv2.GetDaysPermanentPWExipres()

        If CustomRoles.RolesForPageLoad() Then
            If Not Page.IsPostBack Then
                UserDropDownLists()
            End If
        Else
            CustomRoles.TransferIfNotInRole()
            Exit Sub
        End If
    End Sub

    Public Sub ButtonResetPassword_Click(ByVal userid As String)
        GenerateAndSend(userid, CommonUtilsv2.GetUserEmailByLoginID(userid))

    End Sub

    Public Sub GenerateAndSend(ByVal User As String, ByVal Email As String)
        Dim bSuccess As Boolean = False
        Dim strPassword As String = ""

        Try
            strPassword = RandomPassword(10)
            HashValue(strPassword)

            UpdatePassword(User, hash)
            SendPassword(User, Email, strPassword)

            ShowAlert("Password successfully reset. Selected user will receive an e-mail shortly.")
        Catch ex As Exception
            errString = ex.Message
            errLocation = "GenerateAndSend()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())

            ShowAlert("An error ocurred during password reset. An e-mail has been sent.")
        End Try
    End Sub

    Public Sub UpdatePassword(ByVal userName As String, ByVal passwordHash As String)
        Dim lookupValue As Double = Convert.ToDouble(strHOursTempPasswordExpire)
        Dim passwordExpire As Date = Now()
        passwordExpire = passwordExpire.AddHours(lookupValue)
        Dim permPasswordExpire As Date = Now()
        permPasswordExpire = permPasswordExpire.AddDays(Integer.Parse(strDaysPermPasswordExpire))
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strUpdate As String = "Update tbl_usr_logins Set Authentication_Value=@Hash, ChangePassword='True', TempPasswordExpire=@TempPasswordExpire, PermPasswordExpire=@PermPasswordExpire, NumFailedLogins=0 Where UserID=@UserID"
        Dim params As SqlParameter() = { _
            New SqlParameter("@UserID", userName), _
            New SqlParameter("@Hash", passwordHash), _
            New SqlParameter("@TempPasswordExpire", passwordExpire), _
            New SqlParameter("@PermPasswordExpire", permPasswordExpire) _
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strUpdate, CommandType.Text, params)
        Catch ex As Exception
            errString = ex.Message
            errLocation = "UpdatePassword()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try
    End Sub

    Public Sub SendPassword(ByVal userName As String, ByVal email As String, ByVal password As String)
        Dim datetimestamp As String = (DateTime.Now).ToString("MMMM dd, yyyy @ hh:mm")
        Dim strSubject As String = "Password Change Request for UserName: " + userName + " in web application: " + Webapps.Utils.ApplicationSettings.SiteTitle
        If String.Compare("", Webapps.Utils.ApplicationSettings.Environment, True) = 0 Then
        Else
            strSubject = strSubject + " in " + CommonUtilsv2.GetEnvironmentAndHost()
        End If
        Dim strBody As String = "<font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
        strBody += "<b>A password reset request was made on " + datetimestamp + " for UserName: " + userName + "</b>"
        strBody += "</font>"
        strBody += "<table width='100%'><font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
        strBody += "<tr><td colspan=2><hr /></td></tr>"
        strBody += "<tr><td width='25%' align='right'><u>UserName</u>&nbsp;&nbsp;</td><td width='75%' align='left'><u>Temporary Password</u></td></tr>"
        strBody += "<tr><td align='right'>" + userName + "&nbsp;&nbsp;</td><td align='left'>" + password + "</td></tr>"
        strBody += "<tr><td align='left' colspan=2>Please visit " + ApplicationSettings.SiteURL + " to log in.</td></tr>"
        strBody += "<tr><td align='left' colspan=2>Note: You will be required to change your password at next login.</td></tr>"
        strBody += "<tr><td colspan=2><hr /></td></tr>"
        strBody += "</font></table>"
        CommonUtilsv2.SendEMail(ApplicationSettings.ApplicationSourceEmail, email, strBody, strSubject)
    End Sub

    Public Sub SendNewPassword(ByVal userName As String, ByVal email As String, ByVal password As String, ByVal bcc As String)
        Dim datetimestamp As String = (DateTime.Now).ToString("MMMM dd, yyyy @ hh:mm")
        Dim strSubject As String = "New User Created: " + userName + " in web application: " + Webapps.Utils.ApplicationSettings.SiteTitle
        If String.Compare("", Webapps.Utils.ApplicationSettings.Environment, True) = 0 Then
        Else
            strSubject = strSubject + " in " + CommonUtilsv2.GetEnvironmentAndHost()
        End If
        Dim strBody As String = "<font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
        strBody += "<b>A new account was created on " + datetimestamp + " for UserName: " + userName + ".<br />Login credentials below.</b>"
        strBody += "</font>"
        strBody += "<table width='100%'><font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
        strBody += "<tr><td colspan=2><hr /></td></tr>"
        strBody += "<tr><td width='25%' align='right'><u>UserName</u>&nbsp;&nbsp;</td><td width='75%' align='left'><u>Temporary Password</u></td></tr>"
        strBody += "<tr><td align='right'>" + userName + "&nbsp;&nbsp;</td><td align='left'>" + password + "</td></tr>"
        strBody += "<tr><td align='left' colspan=2>Please visit " + ApplicationSettings.SiteURL + " to log in.</td></tr>"
        strBody += "<tr><td align='left' colspan=2>Note: Password must be changed upon initial login.</td></tr>"
        strBody += "<tr><td colspan=2><hr /></td></tr>"
        strBody += "</font></table>"
        If bcc = "no" Then
            CommonUtilsv2.SendEMailBCC(ApplicationSettings.ApplicationSourceEmail, email, strBody, strSubject, ApplicationSettings.ClientUserAccountNoticeEmails)
        Else
            CommonUtilsv2.SendEMailBCC(ApplicationSettings.ApplicationSourceEmail, email, strBody, strSubject, ApplicationSettings.ClientUserAccountNoticeEmails)
        End If
    End Sub

    Public Function RandomPassword(ByVal Length As Integer) As String
        Dim strPassword As String = ""

        strPassword = Membership.GeneratePassword(Length, 0)
        Dim isMatch As Match = Regex.Match(strPassword, "(?!^[0-9]*$)(?!^[a-km-zA-Z]*$)^([a-km-zA-Z0-9]{8,15})$")
        While isMatch.Success = False
            strPassword = Membership.GeneratePassword(Length, 0)
            isMatch = Regex.Match(strPassword, "(?!^[0-9]*$)(?!^[a-km-zA-Z]*$)^([a-km-zA-Z0-9]{8,15})$")
        End While

        Return strPassword
    End Function

    Private Sub ButtonCreateUser_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ButtonCreateUser.Click
        'Validate inputs
        If CommonUtilsv2.Validate(tbFirstName.Text, CommonUtilsv2.DataTypes.String, False, True, False, 100) = False OrElse CommonUtilsv2.Validate(tbLastName.Text, CommonUtilsv2.DataTypes.String, False, True, False, 100) = False Then
            Exit Sub
        End If
        If Not String.IsNullOrEmpty(tbCompany.Text.Trim) Then
            If Not CommonUtilsv2.Validate(tbCompany.Text, CommonUtilsv2.DataTypes.String, True, True, False, 100) Then
                Exit Sub
            End If
        End If
        If Not String.IsNullOrEmpty(tbPosition.Text.Trim) Then
            If Not CommonUtilsv2.Validate(tbPosition.Text, CommonUtilsv2.DataTypes.String, True, True, False, 100) Then
                Exit Sub
            End If
        End If
        If Not String.IsNullOrEmpty(tbeMail.Text.Trim) Then
            If Not CommonUtilsv2.Validate(tbeMail.Text, CommonUtilsv2.DataTypes.String, True, True, False, 100) Then
                Exit Sub
            End If
        End If
        If Not String.IsNullOrEmpty(tbRequestEmail.Text.Trim) Then
            If Not CommonUtilsv2.Validate(tbRequestEmail.Text, CommonUtilsv2.DataTypes.String, True, True, False, 100) Then
                Exit Sub
            End If
        End If
        If Not String.IsNullOrEmpty(tbPhone.Text.Trim) Then
            If Not CommonUtilsv2.Validate(tbPhone.Text, CommonUtilsv2.DataTypes.String, True, True, False, 50) Then
                Exit Sub
            End If
            Exit Sub
        End If
        If Not String.IsNullOrEmpty(tbRequest.Text.Trim) Then
            If Not CommonUtilsv2.Validate(tbRequest.Text, CommonUtilsv2.DataTypes.String, True, True, False, 100) Then
                Exit Sub
            End If
        End If

        If UserNameInUse(tbNewUsername.Text.Trim) OrElse EmailAlreadyUsed(tbeMail.Text.Trim) Then
            lblUserCreated.Visible = True
            lblUserCreated.ForeColor = Drawing.Color.Red
            lblUserCreated.Text = "User name and Email must be unique. New user has not been created. <br/>"
            Exit Sub
        Else
            lblUserCreated.Visible = False
            lblUserCreated.Text = ""
        End If

        Dim lookupValue As Double = Convert.ToDouble(strHoursTempPasswordExpire)
        Dim lookupValue2 As Double = Convert.ToDouble(strDaysPermPasswordExpire)
        Dim passwordExpire As Date = Now()
        passwordExpire = passwordExpire.AddHours(lookupValue)
        Dim permPasswordExpire As Date = Now()
        permPasswordExpire = permPasswordExpire.AddDays(lookupValue2)
        Dim strUserName As String = tbNewUsername.Text.Trim()
        Dim strNewPassword As String = RandomPassword(8)
        Dim strPassword As String = HashValue(strNewPassword)
        Dim strFName As String = tbFirstName.Text.Trim()
        Dim strLName As String = tbLastName.Text.Trim()
        Dim strCompany As String = tbCompany.Text.Trim()
        Dim strPosition As String = tbPosition.Text.Trim()
        Dim strEmail As String = tbeMail.Text.Trim()
        Dim strReqEmail As String = tbRequestEmail.Text.Trim()
        Dim strPhone As String = tbPhone.Text.Trim()
        Dim strCreatedBy As String = tbRequest.Text.Trim()

        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strSQL As String = "Insert Into tbl_usr_Logins (UserID, Authentication_Value, " &
                               " DateCreated, DateLastUpdated, RecordStatus, ChangePassword, TempPasswordExpire, PermPasswordExpire) " &
                               " Values (@UserID, @Authentication_Value, " &
                               " @DateCreated, @DateLastUpdated, @RecordStatus, @ChangePassword, @TempPasswordExpire, @PermPasswordExpire) "
        Dim strSQL2 As String = "Insert Into tbl_usr_Details (UserID, FirstName, LastName, " &
                                " Company, Position, Email, Phone, RequestedBy, DateLastUpdated) " &
                                " Values (@UserID, @FirstName, @LastName, " &
                                " @Company, @Position, @Email, @Phone, @RequestedBy, @DateLastUpdated) "
        Dim params As SqlParameter() = {
            New SqlParameter("@UserID", strUserName),
            New SqlParameter("@Authentication_Value", strPassword),
            New SqlParameter("@DateCreated", Now()),
            New SqlParameter("@DateLastUpdated", Now()),
            New SqlParameter("@RequestedBy", strCreatedBy),
            New SqlParameter("@RecordStatus", 1),
            New SqlParameter("@ChangePassword", 1),
            New SqlParameter("@TempPasswordExpire", passwordExpire),
            New SqlParameter("@PermPasswordExpire", permPasswordExpire)
            }
        Dim params2 As SqlParameter() = {
            New SqlParameter("@UserID", strUserName),
            New SqlParameter("@FirstName", strFName),
            New SqlParameter("@LastName", strLName),
            New SqlParameter("@Company", strCompany),
            New SqlParameter("@Position", strPosition),
            New SqlParameter("@Email", strEmail),
            New SqlParameter("@Phone", strPhone),
            New SqlParameter("@RequestedBy", strCreatedBy),
            New SqlParameter("@DateLastUpdated", Now())
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strSQL, CommandType.Text, params)
            CommonUtilsv2.RunNonQuery(dbKey, strSQL2, CommandType.Text, params2)
            If cbEmailUser.Checked Then 'E-mail User
                SendNewPassword(strUserName, strEmail, strNewPassword, ApplicationSettings.UserAccountNoticeEmails)
            End If
            If cbEmailReq.Checked Then  'E-mail Requestor
                SendNewPassword(strUserName, strReqEmail, strNewPassword, ApplicationSettings.UserAccountNoticeEmails)
            End If
            If cbEmailUser.Checked = False And cbEmailReq.Checked = False Then
                Dim UsereMail As String = CommonUtilsv2.GetCurrentUserEamil(Session("User"))
                If (UsereMail.Contains("@jmsassoc.com")) Then
                    SendNewPassword(strUserName, ApplicationSettings.UserAccountNoticeEmails, strNewPassword, "no")
                Else
                    SendNewPassword(strUserName, ApplicationSettings.ClientUserAccountNoticeEmails, strNewPassword, "no")
                End If
            End If


            lblUserCreated.Visible = True
            lblUserCreated.ForeColor = Drawing.Color.Green
            lblUserCreated.Text = "User Successfully Created."
        Catch ex As Exception
            errString = ex.Message
            errLocation = "Create New User"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            lblUserCreated.Visible = True
            lblUserCreated.ForeColor = Drawing.Color.Red
            lblUserCreated.Text = "Error creating user."
        End Try
    End Sub

    Private Function EmailAlreadyUsed(ByVal emailAddress As String) As Boolean
        Dim bRet As Boolean = False
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim strULsql As String = " SELECT * FROM tbl_usr_Logins ul INNER JOIN tbl_usr_Details ud on ul.UserID = ud.UserID WHERE ud.Email = @emailAddress OR ul.UserID = @emailAddress AND ul.RecordStatus = 1 "
        Dim params1 As SqlParameter() = {New SqlParameter("@emailAddress", emailAddress)}
        Dim myReader1 As SqlDataReader = Nothing

        Try
            myReader1 = CommonUtilsv2.GetDataReader(dbKey, strULsql, CommandType.Text, params1)

            If (Not myReader1 Is Nothing AndAlso myReader1.HasRows) Then
                bRet = True
            End If
        Catch ex As Exception
            errString = ex.Message
            errLocation = "EmailAlreadyUsed()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            Throw ex
        Finally
            If Not myReader1 Is Nothing Then
                myReader1.Close()
            End If
        End Try
        Return bRet
    End Function

    Private Function UserNameInUse(ByVal userName As String) As Boolean
        Dim bRet As Boolean = False
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim strULsql As String = " SELECT * FROM tbl_usr_Logins ul WHERE ul.UserID = @userName AND ul.RecordStatus = 1 "
        Dim params1 As SqlParameter() = {New SqlParameter("@userName", userName)}
        Dim myReader1 As SqlDataReader = Nothing

        Try
            myReader1 = CommonUtilsv2.GetDataReader(dbKey, strULsql, CommandType.Text, params1)

            If (Not myReader1 Is Nothing AndAlso myReader1.HasRows) Then
                bRet = True
            End If
        Catch ex As Exception
            errString = ex.Message
            errLocation = "UserNameInUse()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            Throw ex
        Finally
            If Not myReader1 Is Nothing Then
                myReader1.Close()
            End If
        End Try
        Return bRet
    End Function

    'start role manager
    Protected Sub DropDownListUserRoleSection_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If String.IsNullOrEmpty(tbNewUsername.Text) Then
            lbl_loginErrors.Text = "Please fill in the username before assigning roles."
            lbl_loginErrors.Visible = True
            Exit Sub
        End If


        LabelUserRoleStatus.Visible = False
        LabelUserRoleStatus.Text = ""
        MultiViewUserRoles.Visible = True

        Select Case DropDownListUserRoleSection.SelectedItem.Value

            Case "AssignUnassign"

                'DropDownListUsers.SelectedIndex = 0
                lbxToValues.Items.Clear()
                lbxAssignedToValues.Items.Clear()
                'LabelUser.Visible = True
                'DropDownListUsers.Visible = True
                MultiViewUserRoles.SetActiveView(ViewAssignUnassignUserRoles)
                'move user refresh here, assuming username textbox has a value
                RefreshUserRolesList()
                UserDropDownLists()
            Case "Clone"
                'DropDownListUsers.SelectedIndex = 0
                'LabelUser.Visible = True
                'DropDownListUsers.Visible = True
                MultiViewUserRoles.SetActiveView(ViewCloneUserRoles)
                'move user refresh here, assuming username textbox has a value
                RefreshUserRolesList()
                UserDropDownLists()
            Case Else
                MultiViewUserRoles.Visible = False

        End Select

    End Sub

    Sub RefreshUserRolesList()
        Dim strUsrname As String = tbNewUsername.Text.Trim()
        If Not String.IsNullOrEmpty(strUsrname) Then
            If Not CommonUtilsv2.Validate(strUsrname, CommonUtilsv2.DataTypes.String, True, True, False, 100) Then
                Exit Sub
            End If
        End If
        Dim strSQL As String
        Dim UsereMail As String = tbeMail.Text.Trim()
        'If (UsereMail.Contains("@jmsassoc.com")) Then
        '    strSQL = "SELECT distinct [Role_ID]  FROM tbl_ROLES_UserRoles ur WHERE ur.User_ID = @UserID"

        'Else
        ' strSQL = "SELECT distinct [Role_ID] FROM tbl_ROLES_UserRoles ur WHERE ur.User_ID = @UserID and (NOT (Role_ID = 'Public')) AND (NOT (Role_ID LIKE 'R_Admin_%'))"
        strSQL = "SELECT distinct ur.[Role_ID] , r.RoleName FROM tbl_ROLES_UserRoles ur INNER  JOIN tbl_ROLES_Roles r ON ur.Role_ID =r.Role_ID  WHERE ur.User_ID = @UserID and (NOT (ur.Role_ID = 'Public')) AND (NOT (ur.Role_ID LIKE 'R_Admin_%'))ORDER BY  r.RoleName "
        
        'End If
        LoadDropDownValues(strSQL, "RoleName", "Role_ID", Me.lbxAssignedToValues, strUsrname)

        'If (UsereMail.Contains("@jmsassoc.com")) Then
        '    strSQL = "SELECT distinct [Role_ID] FROM tbl_ROLES_UserRoles where [Role_ID] NOT IN (select [Role_ID] from [dbo].[tbl_ROLES_UserRoles] where [User_ID] = @UserID) ORDER by [Role_ID]"
        'Else
        'strSQL = "SELECT distinct [Role_ID] FROM v_w_ClientUserRoles where [Role_ID] NOT IN (select [Role_ID] from [dbo].[tbl_ROLES_UserRoles] where [User_ID] = @UserID) ORDER by [Role_ID]"
        strSQL = "SELECT distinct [Role_ID], RoleName FROM v_w_ClientUserRoles where [Role_ID] NOT IN (select [Role_ID] from [dbo].[tbl_ROLES_UserRoles] where [User_ID] = @UserID) ORDER by RoleName "
        'End If

        LoadDropDownValues(strSQL, "RoleName", "Role_ID", Me.lbxToValues, strUsrname)

    End Sub

    Protected Sub LoadDropDownValues(ByVal strSQL As String, ByVal textField As String, ByVal dataField As String, ByVal oList As Object, Optional ByVal userName As String = Nothing)
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim myComm As New SqlCommand(strSQL, myConn)
        Dim myReader As SqlDataReader = Nothing
        If Not userName Is Nothing Then
            myComm.Parameters.AddWithValue("@UserID", userName)
        End If
        Try
            myConn.Open()
            myReader = myComm.ExecuteReader()
            oList.DataSource = myReader
            oList.DataTextField = textField
            oList.DataValueField = dataField
            oList.DataBind()
            myReader.Close()

        Catch ex As Exception
            errString = ex.Message
            errLocation = "LoadDropDownValues()for " & dataField
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
            myComm.Dispose()
            myConn.Close()
        End Try
    End Sub


    Private Sub MoveRight(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not String.IsNullOrEmpty(tbNewUsername.Text) Then

            Dim selectedToRStage As String = ""
            If (lbxToValues.SelectedIndex <> -1) Then

                Dim toRefresh As Boolean = False
                Dim showErrorMsg As Boolean = False

                Dim errorMsg As String = ""
                Try
                    For i As Integer = 0 To lbxToValues.Items.Count - 1
                        selectedToRStage = ""
                        If lbxToValues.Items(i).Selected = True Then

                            AddRoleToUser(tbNewUsername.Text, lbxToValues.Items(i).Value)
                        End If
                    Next i

                    RefreshUserRolesList()

                Catch ex As Exception
                    errorMsg = errorMsg + " Could not add the role to the user: " + Server.HtmlEncode(ex.Message)
                    showErrorMsg = True
                End Try
                lbxToValues.SelectedIndex = -1
                If toRefresh = True Then
                    'RefreshListBoxes(selectedUserRole, selectedFromRStage)
                End If
            End If
        End If
    End Sub

    Private Sub MoveLeft(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not String.IsNullOrEmpty(tbNewUsername.Text) Then

            Dim selectedAssignedRoleRStageID As Integer = -1
            If (lbxAssignedToValues.SelectedIndex <> -1) Then

                Dim toRefresh As Boolean = False
                Dim showErrorMsg As Boolean = False

                Dim errorMsg As String = ""
                Try
                    For i As Integer = 0 To lbxAssignedToValues.Items.Count - 1
                        selectedAssignedRoleRStageID = -1
                        If lbxAssignedToValues.Items(i).Selected = True Then
                            DeleteRoleForUser(tbNewUsername.Text, lbxAssignedToValues.Items(i).Value)
                        End If
                    Next i

                    RefreshUserRolesList()

                Catch ex As Exception
                    errorMsg = errorMsg + " Could not add the role to the user: " + Server.HtmlEncode(ex.Message)
                    showErrorMsg = True
                End Try
                lbxAssignedToValues.SelectedIndex = -1
                If toRefresh = True Then
                    'RefreshListBoxes(selectedUserRole, selectedFromRStage)
                End If
            End If
        End If
    End Sub

    Sub DeleteRoleForUser(ByVal user As String, ByVal role As String)

        If Not String.IsNullOrEmpty(user) Then

            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")

            Dim sql As String = "DELETE FROM [dbo].[tbl_ROLES_UserRoles] WHERE Role_ID = @Role_ID and User_ID = @User_ID;"

            Dim params As SqlParameter() = { _
                New SqlParameter("@User_ID", user), _
                New SqlParameter("@Role_ID", role) _
                }

            Try
                CommonUtilsv2.RunNonQuery(dbKey, sql, CommandType.Text, params)

            Catch ex As Exception
                errString = ex.Message
                errLocation = "DeleteRoleForUser()"
                CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            Finally

            End Try

        End If

    End Sub

    Sub AddRoleToUser(ByVal userName As String, ByVal roleName As String)

        If Not DoesUserBelongToRole(userName, roleName) Then
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")

            Dim sql As String = "INSERT INTO [dbo].[tbl_ROLES_UserRoles] ([User_ID],[Role_ID]) VALUES(@User_ID, @Role_ID)"

            Dim params As SqlParameter() = { _
                New SqlParameter("@User_ID", userName), _
                New SqlParameter("@Role_ID", roleName) _
                }

            Try
                CommonUtilsv2.RunNonQuery(dbKey, sql, CommandType.Text, params)

            Catch ex As Exception
                errString = ex.Message
                errLocation = "AddRoleToUser()"
                CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            Finally

            End Try
        End If

    End Sub

    Function DoesUserBelongToRole(ByVal userName As String, ByVal roleName As String) As Boolean

        Dim count As Integer = 0

        If Not String.IsNullOrEmpty(roleName) Then

            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")

            Dim sql As String = "SELECT COUNT(*) FROM [dbo].[tbl_ROLES_UserRoles] WHERE User_ID = @User_ID and @Role_ID = Role_ID"

            Dim params As SqlParameter() = { _
                New SqlParameter("@User_ID", userName), _
                New SqlParameter("@Role_ID", roleName) _
                }

            Dim myReader As SqlDataReader = Nothing
            Try
                myReader = CommonUtilsv2.GetDataReader(dbKey, sql, CommandType.Text, params)

                If Not myReader Is Nothing Then

                    While myReader.Read()

                        If myReader.GetValue(0) Is System.DBNull.Value Then
                            count = 0
                        Else
                            count = myReader.GetInt32(0)
                        End If

                    End While

                End If

            Catch ex As Exception
                errString = ex.Message
                errLocation = "DoesUserBelongToRole()"
                CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            Finally
                If Not myReader Is Nothing Then
                    myReader.Close()
                End If
            End Try

        End If

        If count > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Protected Sub ButtonCloneRolesFromUserToUser_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        If Not String.IsNullOrEmpty(tbNewUsername.Text) Then

            LabelUserRoleStatus.Visible = False
            LabelUserRoleStatus.Text = ""

            Try

                CloneRoles()

                LabelUserRoleStatus.Visible = True
                LabelUserRoleStatus.ForeColor = Drawing.Color.Green
                LabelUserRoleStatus.Text = "Cloned Successfully"

            Catch ex As Exception
                LabelUserRoleStatus.Visible = True
                LabelUserRoleStatus.ForeColor = Drawing.Color.Red
                LabelUserRoleStatus.Text = ex.Message
                errString = ex.Message
                errLocation = "ButtonCloneRolesFromUserToUser_Click()"
                CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            Finally
            End Try

        End If

    End Sub

    Sub CloneRoles()

        For Each lstItem As ListItem In ListboxUsersCloneRoles.Items
            If lstItem.Selected Then
                CloneRolesFromUserToUser(lstItem.Value, tbNewUsername.Text)
            End If
        Next

    End Sub

    Sub CloneRolesFromUserToUser(ByVal sourceUser As String, ByVal destUser As String)

        ' THIS WILL DELETE ALL ROLES FROM destUser BEFORE THE CLONE. destUser will then have only the list of roles assigned to sourceUser
        DeleteAllUserRolesForUser(destUser)

        Dim count As Integer = 0

        If Not String.IsNullOrEmpty(sourceUser) Then

            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")

            Dim sql As String = "SELECT Role_ID FROM [dbo].[tbl_ROLES_UserRoles] WHERE User_ID = @User_ID"

            Dim params As SqlParameter() = { _
                New SqlParameter("@User_ID", sourceUser) _
                }

            Dim myReader As SqlDataReader = Nothing
            Try
                myReader = CommonUtilsv2.GetDataReader(dbKey, sql, CommandType.Text, params)

                If Not myReader Is Nothing Then

                    While myReader.Read()

                        If myReader.GetValue(0) Is System.DBNull.Value Then
                            count = 0
                        Else
                            ' myReader.GetString(0)
                            AddRoleToUser(destUser, myReader.GetString(0))
                        End If

                    End While

                End If

            Catch ex As Exception
                errString = ex.Message
                errLocation = "CloneRolesFromUserToUser()"
                CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            Finally
                If Not myReader Is Nothing Then
                    myReader.Close()
                End If
            End Try

        End If

    End Sub

    Sub DeleteAllUserRolesForUser(ByVal sourceUser As String)

        If Not String.IsNullOrEmpty(sourceUser) Then

            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")

            Dim sql As String = "DELETE FROM [dbo].[tbl_ROLES_UserRoles] WHERE User_ID = @User_ID;"

            Dim params As SqlParameter() = { _
                New SqlParameter("@User_ID", sourceUser) _
                }

            Try
                CommonUtilsv2.RunNonQuery(dbKey, sql, CommandType.Text, params)

            Catch ex As Exception
                errString = ex.Message
                errLocation = "DeleteAllRolesForUser()"
                CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            Finally

            End Try

        End If

    End Sub

    Sub UserDropDownLists()
        Dim strSQL As String = ""
        Dim UsereMail As String = tbeMail.Text.Trim()
        'If (Not String.IsNullOrEmpty(UsereMail)) Then
        'If (UsereMail.Contains("@jmsassoc.com")) Then
        '    strSQL = "SELECT distinct [UserID] FROM [dbo].[tbl_usr_Logins] WHERE (([disabled] IS NULL) OR ([disabled]<>1)) ORDER BY [UserID]"
        'Else
        strSQL = "SELECT distinct [UserID] FROM v_w_ClientUsers WHERE (([disabled] IS NULL) OR ([disabled]<>1)) ORDER BY [UserID]"
        'End If

        'Else
        '    strSQL = "SELECT distinct [UserID] FROM v_w_ClientUsers WHERE (([disabled] IS NULL) OR ([disabled]<>1)) ORDER BY [UserID]"
        'End If
        LoadDropDownValues(strSQL, "UserID", "UserID", Me.ListboxUsersCloneRoles)
    End Sub
    ' end role manager

</script>
<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
<script type="text/javascript">

</script>
    <div align="center" style="height: 100%; width:100%;">
            <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0" class="wrapper">
            <%--<tr class="body_title">
                <td colspan="7" align="left">User Manager</td>
            </tr>--%>
            <tr><td>&nbsp;</td></tr>
             <tr>      
                <td width="100%" colspan="2" align="left" valign="top" class="form_login">
                    <!-- Add Users -->
                    <fieldset runat="server" id="fsAdd" style="width: 100%">
                        <legend class="form_login" style="border-style: none">&nbsp;&nbsp;<b>Create User:
                        </b>&nbsp&nbsp </legend>
                        <table width="100%">
                          <tr>
                              <td align="left" colspan="4">
                              &nbsp;
                              </td>
                          </tr>
                          <tr>
                              <td align="left" colspan="4">
                                    <asp:Label runat="server" ID="lblUserCreated" Visible="false" />
                              </td>
                          </tr>
                          <tr>
                              <td align="left" valign="top" rowspan="1" colspan="4"> 
                                    <asp:Label ID="lbl_loginErrors" runat="server" EnableViewState="False" Visible="False" ForeColor="Red"></asp:Label>
                                    &nbsp;&nbsp;
                              </td>
                          </tr>
                          <tr valign="middle">
                                <td align="right" width="20%">Username: </td>
                                <td align="left" width="30%"><asp:TextBox ID="tbNewUsername" runat="server" MaxLength="50" Width="80%"></asp:TextBox>
                                <asp:RequiredFieldValidator runat="server" ID="v_tbNewUsername" ControlToValidate="tbNewUsername" ErrorMessage="Username is required" Text="*" />
                                </td>
                                <td align="right" width="15%">&nbsp;</td>
                                <td align="left" width="30%">&nbsp;</td>
                          </tr>
                          <tr>
                                <td align="right">First Name: </td>
                                <td align="left"><asp:TextBox ID="tbFirstName" runat="server" MaxLength="100" Width="80%"></asp:TextBox>
                                <asp:RequiredFieldValidator runat="server" ID="v_tbFirstName" ControlToValidate="tbFirstName" ErrorMessage="First Name is required" Text="*" />
                                </td> 
                                <td align="right">Last Name: </td>
                                <td align="left"><asp:TextBox ID="tbLastName" runat="server" MaxLength="100" Width="80%"></asp:TextBox>
                                <asp:RequiredFieldValidator runat="server" ID="v_tbLastName" ControlToValidate="tbLastName" ErrorMessage="Last Name is required" Text="*" />
                                </td>
                          </tr>
                          <tr>
                                <td align="right">Company: </td>
                                <td align="left"><asp:TextBox ID="tbCompany" runat="server" MaxLength="100" Width="80%"></asp:TextBox></td> 
                                <td align="right">Position: </td>
                                <td align="left"><asp:TextBox ID="tbPosition" runat="server" MaxLength="100" Width="80%"></asp:TextBox></td>  
                          </tr>
                          <tr>
                                <td align="right">eMail: </td>
                                <td align="left"><asp:TextBox ID="tbeMail" runat="server" MaxLength="100" Width="80%"></asp:TextBox>
                                <asp:RequiredFieldValidator runat="server" ID="v_tbeMail" ControlToValidate="tbeMail" ErrorMessage="E-mail is required" Text="*" />
                                </td> 
                                <td align="right">Phone: </td>
                                <td align="left"><asp:TextBox ID="tbPhone" runat="server" MaxLength="50" Width="80%"></asp:TextBox></td>  
                          </tr>
                          <tr>
                                <td align="right">Requested By: </td>
                                <td align="left"><asp:TextBox ID="tbRequest" runat="server" MaxLength="50" Width="80%"></asp:TextBox></td>
                                <td align="right">Requestor E-mail: </td>
                                <td align="left"><asp:TextBox ID="tbRequestEmail" runat="server" MaxLength="100" Width="80%"></asp:TextBox></td>
                          </tr>
                          <tr>
                                <td align="right">E-mail Requestor: </td>
                                <td align="left"><asp:CheckBox runat="server" ID="cbEmailReq" /></td>
                                <td align="right"></td>
                                <td align="left"></td>
                          </tr>
                          <tr>
                                <td align="right">E-mail User: </td>
                                <td align="left"><asp:CheckBox runat="server" ID="cbEmailUser" /></td>
                                <td align="right"></td>
                                <td align="left"></td>
                          </tr>
                        </table>
                    </fieldset>
                </td>            
            </tr>
            <tr>
                <td width="100%" valign="top" class="form_login">
                    <fieldset runat="server" id="fsRoles" style="width: 100%">
                        <legend class="form_login" style="border-style: none">&nbsp;&nbsp;<b>Add/Remove Role
                            To User:</b>&nbsp&nbsp </legend>
                        <asp:Label ID="LabelUserRoleStatus" runat="server" ></asp:Label>
                        <table width="100%">
                            <tr>
                                <td align="left" width="100%">Section:<asp:DropDownList 
                                        ID="DropDownListUserRoleSection" runat="server" AutoPostBack="True" 
                                        onselectedindexchanged="DropDownListUserRoleSection_SelectedIndexChanged">
                                        <asp:ListItem Value="Select Section" Text = "Select Section" Selected="True" />
                                        <asp:ListItem Value="AssignUnassign" Text="Assign Unassign User Roles" />
                                        <asp:ListItem Value="Clone" Text="Clone User Roles" />
                                    </asp:DropDownList>
                                    <br /></td>
                            </tr>
                            <tr>
                                <td align="right" width="100%"><br /></td>
                            </tr>
                        </table>
                        <asp:MultiView ID="MultiViewUserRoles" runat="server">
                            <asp:View ID="ViewAssignUnassignUserRoles" runat="server">
                                <table>
                                    <tr>
                                        <td align="center" style="width: 300px" >Available roles: </td>
                                        <td align="center" colspan="1" valign="middle" style="width: 100px"></td>
                                        <td align="center" style="width: 300px">Roles assigned to selected user: </td>
                                    </tr>                                        
                                     <tr>
                                        <td align="right" rowspan="15" style="width: 300px">
                                            <asp:listbox id="lbxToValues" runat="server" SelectionMode="Multiple" Rows="10" HorizontalScrollbar="true" style="min-width: 350px">
                                            </asp:listbox>
                                        </td>                          
                                        <td align="center" colspan="1" rowspan="10" valign="middle" style="width: 150px">
                                            <br />
                                            <asp:Button ID="btnMoveRight" runat="server" CssClass="Submit_button" Width="50px" Text=" > " OnClick="MoveRight" />
                                            <br />
                                            
                                            <br />
                                            <asp:Button ID="btnMoveLeft" runat="server" CssClass="Submit_button" Width="50px" Text=" < " OnClick="MoveLeft" />
                                            <br />
                                            <br />
                                        </td>                            
                                        <td align="left" rowspan="15" style="width: 300px">
                                            <asp:listbox id="lbxAssignedToValues" runat="server" SelectionMode="Multiple" Rows="10" HorizontalScrollbar="true" style="min-width: 350px">
                                            </asp:listbox>
                                        </td>
                                    </tr>                              
                                </table>
                            </asp:View>
                            <asp:View ID="ViewCloneUserRoles" runat="server">
                                    <br />
                                    <table style="width: 30%" border="0">
                                        <tr>
                                            <td width="80%">
                                            Source User:
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="right">
                                                <asp:listbox id="ListboxUsersCloneRoles" runat="server" SelectionMode="Single" Rows="10" style="width: 100%">
                                                </asp:listbox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="right">
                                                <asp:Button ID="ButtonCloneRolesFromUserToUser" runat="server" 
                                                    onclick="ButtonCloneRolesFromUserToUser_Click" Text="Clone" CssClass="Submit_button" />
                                            </td>
                                        </tr>
                                    </table>
                                    <br />
                                    <br />
                            </asp:View>
                        </asp:MultiView>
                        <br />
                    </fieldset>
                </td>
            </tr> 
          <tr>
            <td align="right"> 
                <asp:Button ID="ButtonCreateUser" runat="server" Text="Create Account" CssClass="Submit_button" />
            </td>
          </tr>
        </table>       
    </div>
</asp:Content>
