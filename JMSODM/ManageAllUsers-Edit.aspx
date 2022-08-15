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
    '  Dim strUserNoticeAddress As String
    Dim hash As String = ""
    Dim pageName As String
    Dim SessionPrefix As String = "ManageUsers-Edit_"
    Dim strUID As String

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
        lblMessage.Text = ""
        '  strUserNoticeAddress = CommonUtilsv2.GetUserAccountNoticeEmails()
        strHoursTempPasswordExpire = ApplicationSettings.HoursTempPWExpires 'CommonUtilsv2.GetHoursTempPWExpires()
        strDaysPermPasswordExpire = ApplicationSettings.DaysPermanentPWExipres 'CommonUtilsv2.GetDaysPermanentPWExipres()
        If CustomRoles.RolesForPageLoad() Then
            strUID = Request.QueryString.Get("UID")
            If Not Page.IsPostBack Then
                If Not String.IsNullOrEmpty(strUID) Then
                    If Not CommonUtilsv2.Validate(strUID, CommonUtilsv2.DataTypes.String, True, True, True, 100) Then
                        Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
                        Exit Sub
                    Else
                        Session(SessionPrefix & "UID") = strUID
                    End If
                End If

                ddlUserID.Items.Add("Please Select")
                ddlUserID.AppendDataBoundItems = True
                loadUserDD()
                UserDropDownLists()
                If String.IsNullOrEmpty(strUID) Then
                Else
                    LoadUser(strUID)
                End If

            End If
        Else
            CustomRoles.TransferIfNotInRole()
            Exit Sub
        End If
    End Sub

    Public Sub loadUserDD()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim strSQL As String = "Select UserID From tbl_usr_Logins WHERE  (([disabled] IS NULL) OR ([disabled]<>1)) Order By UserID"

        CommonUtilsv2.PopulateDropDownBox(dbKey, strSQL, ddlUserID, "UserID", "UserID")
    End Sub

    Public Sub ButtonResetPassword_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Dim eMail As String = ""
        lblMessage.Visible = True
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim strSQL As String = "Select Email From tbl_usr_Details Where UserID=@UserID"
        Dim myReader As SqlDataReader = Nothing
        Dim params As SqlParameter() = { _
            New SqlParameter("@UserID", ddlUserID.SelectedValue) _
            }
        Try
            myReader = CommonUtilsv2.GetDataReader(dbKey, strSQL, CommandType.Text, params)
            myReader.Read()
            eMail = myReader(0)
            GenerateAndSend(ddlUserID.SelectedValue, eMail)
            lblMessage.ForeColor = Drawing.Color.Green
            lblMessage.Text = "User account password reset completed."
        Catch ex As Exception
            errString = ex.Message
            errLocation = "Get e-mail address from userid"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try
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
        Dim lookupValue As Double = Convert.ToDouble(strHoursTempPasswordExpire)
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
        CommonUtilsv2.SendEMailBCC(ApplicationSettings.ApplicationSourceEmail, email, strBody, strSubject, ApplicationSettings.UserAccountNoticeEmails)
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
            CommonUtilsv2.SendEMailBCC(ApplicationSettings.ApplicationSourceEmail, email, strBody, strSubject)
        Else
            CommonUtilsv2.SendEMailBCC(ApplicationSettings.ApplicationSourceEmail, email, strBody, strSubject, bcc)
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

    Private Sub ResetEditField()
        tbCompanyEdit.Text = ""
        tbEmailEdit.Text = ""
        tbFirstNameEdit.Text = ""
        tbLastNameEdit.Text = ""
        tbPhoneEdit.Text = ""
        tbPositionEdit.Text = ""
        tbRequestEdit.Text = ""
    End Sub

    Private Sub btnEditUser_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEditUser.Click
        'validate inputs
        If CommonUtilsv2.Validate(tbFirstNameEdit.Text, CommonUtilsv2.DataTypes.String, False, True, False, 100) = False OrElse CommonUtilsv2.Validate(tbLastNameEdit.Text, CommonUtilsv2.DataTypes.String, False, True, False, 100) = False Then
            Exit Sub
        End If
        If Not String.IsNullOrEmpty(tbCompanyEdit.Text.Trim) Then
            If Not CommonUtilsv2.Validate(tbCompanyEdit.Text, CommonUtilsv2.DataTypes.String, True, True, False, 100) Then
                Exit Sub
            End If
        End If
        If Not String.IsNullOrEmpty(tbPositionEdit.Text.Trim) Then
            If Not CommonUtilsv2.Validate(tbPositionEdit.Text, CommonUtilsv2.DataTypes.String, True, True, False, 100) Then
                Exit Sub
            End If
        End If
        If Not String.IsNullOrEmpty(tbEmailEdit.Text.Trim) Then
            If Not CommonUtilsv2.Validate(tbEmailEdit.Text, CommonUtilsv2.DataTypes.String, True, True, False, 100) Then
                Exit Sub
            End If
        End If
        If Not String.IsNullOrEmpty(tbPhoneEdit.Text.Trim) Then
            If Not CommonUtilsv2.Validate(tbPhoneEdit.Text, CommonUtilsv2.DataTypes.String, True, True, False, 50) Then
                Exit Sub
            End If
        End If
        If Not String.IsNullOrEmpty(tbRequestEdit.Text.Trim) Then
            If Not CommonUtilsv2.Validate(tbRequestEdit.Text, CommonUtilsv2.DataTypes.String, True, True, False, 100) Then
                Exit Sub
            End If
        End If

        lblMessage.Visible = True
        Dim strUserName As String = ddlUserID.SelectedValue
        Dim strFName As String = tbFirstNameEdit.Text.Trim()
        Dim strLName As String = tbLastNameEdit.Text.Trim()
        Dim strCompany As String = tbCompanyEdit.Text.Trim()
        Dim strPosition As String = tbPositionEdit.Text.Trim()
        Dim strEmail As String = tbEmailEdit.Text.Trim()
        Dim strPhone As String = tbPhoneEdit.Text.Trim()
        Dim strCreatedBy As String = tbRequestEdit.Text.Trim()

        Dim strSQL As String = "Update tbl_usr_Details Set "
        strSQL += " FirstName=@FirstName, LastName=@LastName, Company=@Company, Position=@Position, Email=@Email, Phone=@Phone, DateLastUpdated=@DateLastUpdated, "
        strSQL += " RequestedBy=@RequestedBy "
        strSQL += " Where UserID=@UserID "

        Dim myConn As New SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced"))
        Dim myComm As New SqlClient.SqlCommand(strSQL, myConn)
        Try
            myConn.Open()
            myComm.Parameters.AddWithValue("@UserID", strUserName)
            myComm.Parameters.AddWithValue("@FirstName", strFName)
            myComm.Parameters.AddWithValue("@LastName", strLName)
            myComm.Parameters.AddWithValue("@Company", strCompany)
            myComm.Parameters.AddWithValue("@Position", strPosition)
            myComm.Parameters.AddWithValue("@Email", strEmail)
            myComm.Parameters.AddWithValue("@Phone", strPhone)
            myComm.Parameters.AddWithValue("@DateLastUpdated", Now())
            myComm.Parameters.AddWithValue("@RequestedBy", strCreatedBy)
            myComm.ExecuteNonQuery()
            lblMessage.ForeColor = Drawing.Color.Green
            lblMessage.Text = "User Successfully Updated."
        Catch ex As Exception
            errString = ex.Message
            errLocation = "Update User"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
            lblMessage.ForeColor = Drawing.Color.Red
            lblMessage.Text = "Error creating user."
        Finally
            myComm.Dispose()
            myConn.Close()
        End Try
    End Sub

    Private Sub LoadUser(ByVal strUID As String)

        If CommonUtilsv2.Validate(strUID, CommonUtilsv2.DataTypes.String, True, True, False, 100) = False Then
            Exit Sub
        End If
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strSQL As String = "Select FirstName, LastName, Company, Position, Email, Phone, RequestedBy From tbl_usr_Details Where UserID=@UserID"
        Dim myReader As SqlDataReader = Nothing
        Dim params As SqlParameter() = { _
            New SqlParameter("@UserID", strUID) _
            }
        Try
            myReader = CommonUtilsv2.GetDataReader(dbKey, strSQL, CommandType.Text, params)
            While myReader.Read
                ddlUserID.SelectedValue = strUID
                If Not IsDBNull(myReader(0)) Then
                    tbFirstNameEdit.Text = myReader(0)
                Else
                    tbFirstNameEdit.Text = ""
                End If
                If Not IsDBNull(myReader(1)) Then
                    tbLastNameEdit.Text = myReader(1)
                Else
                    tbLastNameEdit.Text = ""
                End If
                If Not IsDBNull(myReader(2)) Then
                    tbCompanyEdit.Text = myReader(2)
                Else
                    tbCompanyEdit.Text = ""
                End If
                If Not IsDBNull(myReader(3)) Then
                    tbPositionEdit.Text = myReader(3)
                Else
                    tbPositionEdit.Text = ""
                End If
                If Not IsDBNull(myReader(4)) Then
                    tbEmailEdit.Text = myReader(4)
                Else
                    tbEmailEdit.Text = ""
                End If
                If Not IsDBNull(myReader(5)) Then
                    tbPhoneEdit.Text = myReader(5)
                Else
                    tbPhoneEdit.Text = ""
                End If
                If Not IsDBNull(myReader(6)) Then
                    tbRequestEdit.Text = myReader(6)
                Else
                    tbRequestEdit.Text = ""
                End If
            End While
        Catch ex As Exception
            errString = ex.Message
            errLocation = "LoadUser"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try
        'update 
        RefreshUserRolesList()
        If Not ddlUserID.SelectedItem.Text = "Please Select" Then
            btnEditUser.Enabled = True
            btnReset.Enabled = True
            btnDisableUser.Enabled = True
        Else
            btnEditUser.Enabled = False
            btnReset.Enabled = False
            btnDisableUser.Enabled = False
            ResetEditField()
        End If
    End Sub
    Private Sub ddlUserID_IndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUserID.SelectedIndexChanged
        If CommonUtilsv2.Validate(ddlUserID.SelectedValue, CommonUtilsv2.DataTypes.String, True, True, True, 50) Then
            LoadUser(ddlUserID.SelectedValue)
        Else
            Exit Sub
        End If
    End Sub

    'start role manager
    Protected Sub DropDownListUserRoleSection_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlUserID.SelectedIndex = 0 Then
            lblMessage.Text = "Please select a userid before assigning roles."
            lblMessage.ForeColor = Drawing.Color.Red
            lblMessage.Visible = True
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
            Case "Clone"
                'DropDownListUsers.SelectedIndex = 0
                'LabelUser.Visible = True
                'DropDownListUsers.Visible = True
                MultiViewUserRoles.SetActiveView(ViewCloneUserRoles)

                'Case "Templates"

                '    DropDownListUserTemplates.SelectedIndex = 0
                '    LabelUser.Visible = False
                '    DropDownListUsers.Visible = False
                '    MultiViewUserRoles.SetActiveView(ViewUserTemplates)

            Case Else
                MultiViewUserRoles.Visible = False
                'DropDownListUsers.SelectedIndex = 0
                'LabelUser.Visible = False
                'DropDownListUsers.Visible = False

        End Select

    End Sub


    Sub RefreshUserRolesList()
        Dim strSelected = ddlUserID.SelectedValue
        If Not String.IsNullOrEmpty(strSelected) Then
            If Not CommonUtilsv2.Validate(strSelected, CommonUtilsv2.DataTypes.String, True, True, False, 100) Then
                Exit Sub
            End If
        End If

        Dim strSQL As String = "SELECT distinct [Role_ID]  FROM [dbo].[tbl_ROLES_UserRoles] ur WHERE ur.User_ID = @UserID"
        LoadDropDownValues(strSQL, "Role_ID", "Role_ID", Me.lbxAssignedToValues, strSelected)

        strSQL = "SELECT distinct [Role_ID] FROM [dbo].[tbl_ROLES_Roles] where [Role_ID] NOT IN (select [Role_ID] from [dbo].[tbl_ROLES_UserRoles] where [User_ID] = @UserID) ORDER BY [Role_ID]"
        LoadDropDownValues(strSQL, "Role_ID", "Role_ID", Me.lbxToValues, strSelected)

    End Sub

    Protected Sub LoadDropDownValues(ByVal strSQL As String, ByVal textField As String, ByVal dataField As String, ByVal oList As Object, Optional ByVal userName As String = Nothing)
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim myComm As New SqlCommand(strSQL, myConn)
        If Not userName Is Nothing Then
            myComm.Parameters.AddWithValue("@UserID", userName)
        End If
        Try
            myConn.Open()
            Dim myReader As SqlDataReader = myComm.ExecuteReader()
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
            myComm.Dispose()
            myConn.Close()
        End Try
    End Sub


    Private Sub MoveRight(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not ddlUserID.SelectedIndex = 0 Then

            Dim selectedToRStage As String = ""
            If (lbxToValues.SelectedIndex <> -1) Then

                Dim toRefresh As Boolean = False
                Dim showErrorMsg As Boolean = False

                Dim errorMsg As String = ""
                Try
                    For i As Integer = 0 To lbxToValues.Items.Count - 1
                        selectedToRStage = ""
                        If lbxToValues.Items(i).Selected = True Then

                            AddRoleToUser(ddlUserID.SelectedValue, lbxToValues.Items(i).Text)
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
        If Not ddlUserID.SelectedIndex = 0 Then

            Dim selectedAssignedRoleRStageID As Integer = -1
            If (lbxAssignedToValues.SelectedIndex <> -1) Then

                Dim toRefresh As Boolean = False
                Dim showErrorMsg As Boolean = False

                Dim errorMsg As String = ""
                Try
                    For i As Integer = 0 To lbxAssignedToValues.Items.Count - 1
                        selectedAssignedRoleRStageID = -1
                        If lbxAssignedToValues.Items(i).Selected = True Then
                            DeleteRoleForUser(ddlUserID.SelectedValue, lbxAssignedToValues.Items(i).Value)
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

        If Not ddlUserID.SelectedIndex = 0 Then

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
                CloneRolesFromUserToUser(lstItem.Value, ddlUserID.SelectedValue)
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
        Dim strSQL As String = "SELECT distinct [UserID] FROM [dbo].[tbl_usr_Logins] ORDER BY [UserID]"
        LoadDropDownValues(strSQL, "UserID", "UserID", Me.ListboxUsersCloneRoles)
    End Sub
    ' end role manager

    Public Sub ButtonDisableUser_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDisableUser.Click
        Dim eMail As String = ""
        lblMessage.Visible = True
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strUpdate As String = "Update tbl_usr_logins Set Disabled=1, RecordUpdatedBy=@RecordUpdatedBy, DateLastUpdated=@DateLastUpdated Where UserID=@UserID"
        Dim params As SqlParameter() = { _
            New SqlParameter("@UserID", ddlUserID.SelectedValue), _
            New SqlParameter("@RecordUpdatedBy", Session("User")), _
            New SqlParameter("@DateLastUpdated", Now()) _
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strUpdate, CommandType.Text, params)

            lblMessage.Visible = True
            lblMessage.ForeColor = Drawing.Color.Green
            lblMessage.Text = "User account has been disabled."

        Catch ex As Exception
            errString = ex.Message
            errLocation = "ButtonDisableUser_Click()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try

        'send email notification
        Dim datetimestamp As String = (DateTime.Now).ToString("MMMM dd, yyyy @ hh:mm")
        Dim strSubject As String = "Login Account for user " + ddlUserID.SelectedValue + " has been disabled in web application: " + Webapps.Utils.ApplicationSettings.SiteTitle
        If String.Compare("", Webapps.Utils.ApplicationSettings.Environment, True) = 0 Then
        Else
            strSubject = strSubject + " in " + CommonUtilsv2.GetEnvironmentAndHost()
        End If
        Dim strBody As String = "<font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
        strBody += "Login Account for user <b>" + ddlUserID.SelectedValue + "</b> has been disabled on <b>" + datetimestamp + "</b> for website <b>" + Webapps.Utils.ApplicationSettings.SiteTitle + ".</b>"

        CommonUtilsv2.SendEMailBCC(ApplicationSettings.ApplicationSourceEmail, ApplicationSettings.UserAccountNoticeEmails, strBody, strSubject)

    End Sub

    Public Sub ButtonAddLookup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddLookup.Click
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strInsert As String = "prc_InsertUserLookup"
        Dim UserId As String = ddlUserID.SelectedValue
        Dim params As SqlParameter() = {New SqlParameter("@UserId", UserId)}
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strInsert, CommandType.StoredProcedure, params)

        Catch ex As Exception
            errString = ex.Message
            errLocation = "AddLookup_Click()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try
    End Sub
</script>
<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
    <script type="text/javascript">
</script>
    <div align="center" style="height: 100%">
            <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0" class="wrapper">
           <%-- <tr class="body_title">
                <td colspan="7" align="left">User Manager</td>
            </tr>--%>
            <tr><td>&nbsp;</td></tr>
             <tr>      
                <td width="100%" colspan="2" align="left" valign="top" class="form_login">
                    <!-- Edit Users -->
                    <fieldset runat="server" id="fsEdit" style="width: 100%">
                        <legend class="form_login" style="border-style: none">&nbsp;&nbsp;<b>Edit User / Reset Password:
                        </b>&nbsp&nbsp </legend>
                        <table width="100%">
                          <tr>
                            <td>&nbsp;</td>
                          </tr>
                          <tr>
                              <td align="left" colspan="4">
                                    <asp:Label runat="server" ID="lblMessage" />
                              </td>
                          </tr>
                        <tr valign="middle">
                            <td align="right" width="20%">
                                <asp:Label ID="lblUserID" runat="server" Text="User ID:"></asp:Label>
                            </td>    
                            <td align="left" width="30%">
                                <asp:DropDownList ID="ddlUserID" runat="server" Width="200px" AutoPostBack="true" />
                                &nbsp;Password:&nbsp;<asp:Button runat="server" ID="btnReset" Text="Reset and Notify" Width="127px" Enabled="false" CssClass="Submit_button" />
                            </td>
                            <td width="20%"></td>
                            <td width="30%"></td>
                        </tr>
                          <tr>
                                <td align="right">First Name: </td>
                                <td align="left"><asp:TextBox ID="tbFirstNameEdit" runat="server" MaxLength="100" Width="80%"></asp:TextBox>
                                <asp:RequiredFieldValidator runat="server" ID="v_tbFirstNameEdit" ControlToValidate="tbFirstNameEdit" ErrorMessage="First Name is required" Text="*" />
                                </td> 
                                <td align="right">Last Name: </td>
                                <td align="left"><asp:TextBox ID="tbLastNameEdit" runat="server" MaxLength="100" Width="80%"></asp:TextBox>
                                <asp:RequiredFieldValidator runat="server" ID="v_tbLastNameEdit" ControlToValidate="tbLastNameEdit" ErrorMessage="Last Name is required" Text="*" />
                                </td>
                          </tr>
                          <tr>
                                <td align="right">Company: </td>
                                <td align="left"><asp:TextBox ID="tbCompanyEdit" runat="server" MaxLength="100" Width="80%"></asp:TextBox></td> 
                                <td align="right">Position: </td>
                                <td align="left"><asp:TextBox ID="tbPositionEdit" runat="server" MaxLength="100" Width="80%"></asp:TextBox></td>  
                          </tr>
                          <tr>
                                <td align="right">eMail: </td>
                                <td align="left"><asp:TextBox ID="tbEmailEdit" runat="server" MaxLength="100" Width="80%"></asp:TextBox>
                                <asp:RequiredFieldValidator runat="server" ID="v_tbEmailEdit" ControlToValidate="tbEmailEdit" ErrorMessage="E-mail is required" Text="*" />
                                </td> 
                                <td align="right">Phone: </td>
                                <td align="left"><asp:TextBox ID="tbPhoneEdit" runat="server" MaxLength="50" Width="80%"></asp:TextBox></td>  
                          </tr>
                          <tr>
                                <td align="right">Requested By: </td>
                                <td align="left"><asp:TextBox ID="tbRequestEdit" runat="server" MaxLength="50" Width="80%"></asp:TextBox></td>
                                <td align="right"></td>
                                <td align="left"></td>
                          </tr>
                          <tr>
                            <td align="left">                                
                            </td>
                            <td align="left" colspan="3">
                                <asp:Button runat="server" ID="btnEditUser" Text="Update Account" Width="127px" Enabled="false" CssClass="Submit_button" />
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Button runat="server" ID="btnDisableUser" Text="Disable Login" Width="127px" Enabled="false" CssClass="Submit_button" />
                                <asp:Button runat="server" ID="btnAddLookup" Text="Add LookUp" Width="127px" Enabled="True" CssClass="Submit_button" />
                            </td>
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
                                            <asp:listbox id="lbxToValues" runat="server" SelectionMode="Multiple" Rows="10"  HorizontalScrollbar="true" style="min-width: 350px">
                                            </asp:listbox>
                                        </td>                          
                                        <td align="center" colspan="1" rowspan="10" valign="middle" style="width: 150px">
                                            <br />
                                            <asp:Button ID="btnMoveRight" runat="server" Text=" > " CssClass="Submit_button" Width="50px" OnClick="MoveRight" />
                                            <br />
                                            
                                            <br />
                                            <asp:Button ID="btnMoveLeft" runat="server" Text=" < " CssClass="Submit_button" Width="50px" OnClick="MoveLeft" />
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
        </table>       
    </div>
</asp:Content>