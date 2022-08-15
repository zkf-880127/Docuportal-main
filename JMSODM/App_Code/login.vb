Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data
Imports Webapps.Utils
Imports BCrypt

Public Class login
    'declare datamembers
    Dim iCaseID As Integer = -1
    Dim strUserType As String = "NA"
    Dim strFullName As String
    Dim strID As String
    Dim strPW As String
    Dim strPWHash As String
    Dim bDisabled As Boolean = False
    Dim bAccountLocked As Boolean = False
    Dim bLoggedIn As Boolean = False
    Dim iNumFailedLogins As Integer = 0
    Dim bRequirePasswordChange As Boolean = False
    Dim bUserCanChangePW As Boolean = True

    Dim bValidUserID As Boolean = True
    Dim bValidUserPW As Boolean = True

    Dim bValidNewUserPW As Boolean = True

    Dim dTempPass As DateTime = Now()
    Dim dPermPass As DateTime = Now()

    Dim iDaysPermPwExpires As Integer = Webapps.Utils.ApplicationSettings.DaysPermanentPWExipres
    Dim bPermanentPasswordExpires As Boolean = Webapps.Utils.ApplicationSettings.RequirePWChange ' 
    Dim iDaysAfterPermPwExpiresUserCannotChangePW As Integer = Webapps.Utils.ApplicationSettings.DaysAfterPWExpirationPWCannotBeChanged
    Dim iNumberOfFailedLoginsBeforeAccountLockout As Integer = Webapps.Utils.ApplicationSettings.NumberOfFailedLoginsBeforeAccountLockout
    Dim iNumberOfPriorPWNotbeUsed As Integer = Webapps.Utils.ApplicationSettings.NumberOfPriorPWsNotBeUsed

    Dim strIP As String = ""

    Public Sub New(ByVal id As String, ByVal pw As String)
        strID = id
        strPW = pw
    End Sub

    Public Sub New(ByVal id As String, ByVal pw As String, ByVal ip As String)
        strID = id
        strPW = pw
        strIP = ip
    End Sub

    Public Sub New()
        strID = ""
        strPW = ""
        strPWHash = ""
        strIP = ""
        dTempPass = Now()
        dPermPass = Now()
    End Sub

    Public ReadOnly Property NumFailedLogins() As Integer
        Get
            Return iNumFailedLogins
        End Get
    End Property

    Public ReadOnly Property RequirePasswordChange() As Boolean
        Get
            Return bRequirePasswordChange
        End Get
    End Property

    Public ReadOnly Property AccountDisabled() As Boolean
        Get
            Return bDisabled
        End Get
    End Property

    Public ReadOnly Property AccountLocked() As Boolean
        Get
            Return bAccountLocked
        End Get
    End Property


    Public ReadOnly Property LoggedIn() As Boolean
        Get
            Return bLoggedIn
        End Get
    End Property

    Public ReadOnly Property CanChangePW() As Boolean
        Get
            Return bUserCanChangePW
        End Get
    End Property

    Public ReadOnly Property IsIDValid() As Boolean
        Get
            Return bValidUserID
        End Get
    End Property

    Public ReadOnly Property IsPWValid() As Boolean
        Get
            Return bValidUserPW
        End Get
    End Property

    Public ReadOnly Property IsNewPWValid() As Boolean
        Get
            Return bValidNewUserPW
        End Get
    End Property

    Public ReadOnly Property NumberOfPreviousPasswordNotReusable() As Integer
        Get
            Return iNumberOfPriorPWNotbeUsed
        End Get
    End Property

    Public ReadOnly Property FullName() As String
        Get
            Return strFullName
        End Get
    End Property
    Public ReadOnly Property UserType() As String
        Get
            Return strUserType
        End Get
    End Property

    Public ReadOnly Property CaseID() As Integer
        Get
            Return iCaseID
        End Get
    End Property


    Public Function Login(ByVal id As String, ByVal pw As String, ByVal ip As String) As Integer
        strID = id
        strPW = pw
        strIP = ip
        AccountVerify()

        If bDisabled Then
            'update failed login
            LogFailedLogin(strID, strIP)
        End If

        If bValidUserID Then
            If bLoggedIn Then
                ' updat Success login
                LogSuccesLogin(strID, strIP)
                UpdatePermanentPasswordExpiration(strID)
            Else
                'update failed login
                LogFailedLogin(strID, strIP)
            End If
            UpdateFailedLoginAttempts(strID, iNumFailedLogins)
        End If

        Return GetReturnCode()
    End Function


    Public Function ChangePassword(ByVal id As String, ByVal newPw As String, ByVal oldPw As String) As Integer
        Dim iRet As Integer = 0
        'If String.IsNullOrEmpty(newPw) Or String.IsNullOrEmpty(oldPw) Or String.IsNullOrEmpty(id) Then
        '    'Invalid info
        '    Return 7
        'End If
        strID = id
        strPW = oldPw
        AccountVerify()
        If bDisabled Or Not bLoggedIn Then
            Return GetReturnCode()
        End If
        bValidNewUserPW = VerifyNewPassword(newPw)
        If Not bValidNewUserPW Then
            'invalid new pw. Cannot be reused
            Return 7
        End If

        Dim bRet As Boolean = False
        Dim strHashOfNewPw As String = BCrypt.Net.BCrypt.HashPassword(newPw, 12)
        'Dim iDaysPermPwExpires As Integer = CommonUtilsv2.GetDaysPermanentPWExipres()
        Dim dPermPWExpiresDate As Date = Now.AddDays(iDaysPermPwExpires)

        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strUpdate As String = "Update tbl_usr_Logins Set Authentication_Value=@PWHash, PermPasswordExpire=@PermPasswordExpiresDate, DateLastUpdated=@DateLastUpdated, ChangePassword='False' Where [UserID]=@UserID"
        Dim params As SqlParameter() = { _
            New SqlParameter("@UserID", strID), _
            New SqlParameter("@PWHash", strHashOfNewPw), _
            New SqlParameter("@PermPasswordExpiresDate", dPermPWExpiresDate), _
            New SqlParameter("@DateLastUpdated", Now()) _
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strUpdate, CommandType.Text, params)
            ' insert into PW change Log
            LogPasswordChange(strID, strHashOfNewPw)
            bRet = True
            iRet = 9
        Catch ex As Exception
            bRet = False
            iRet = 0
            Throw New Exception("Error in Change password.  The detail is: " & ex.ToString())
        End Try
        Return iRet

    End Function

    Private Function GetReturnCode() As Integer
        Dim iRet As Integer = 0

        If IsIDValid = False Then
            ' invalid user id
            Return 1
        End If
        If bValidUserPW = False Then
            ' invalid pw
            If bAccountLocked = True Then
                ' Account Locked
                Return 4
            Else
                Return 2
            End If
        End If
        If bDisabled = True Then
            ' Account disabled
            Return 3
        End If

        If bAccountLocked = True Then
            ' Account Locked
            Return 4
        End If
        If bLoggedIn = True Then
            If bRequirePasswordChange = True Then
                If bUserCanChangePW Then
                    ' user is required to change password
                    iRet = 5
                Else
                    'user cannot change password.  Need send pw change request
                    iRet = 6
                End If
            Else
                iRet = 9
            End If
        End If
        Return iRet
    End Function

    Private Sub AccountVerify()

        Dim bLoginExist As Boolean = False
        Dim bVerify As Boolean = False
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim strSQLAccount As String = "select ul.changepassword, ul.temppasswordexpire, ul.permpasswordexpire, ul.disabled, ul.Authentication_Value, ul.numfailedlogins, ud.firstname + ' ' + ud.lastname as FullName from tbl_usr_Logins ul inner join tbl_usr_Details ud on ul.userid=ud.userid where ul.[UserID] = @userid"
        Dim myreader As SqlDataReader = Nothing
        Dim acctparams As SqlParameter() = { _
            New SqlParameter("@userid", strID) _
            }
        Try
            myreader = CommonUtilsv2.GetDataReader(dbKey, strSQLAccount, CommandType.Text, acctparams)
            If myreader.HasRows Then
                bLoginExist = True
                myreader.Read()

                If Not IsDBNull(myreader(0)) Then
                    bRequirePasswordChange = myreader(0)
                End If
                If Not IsDBNull(myreader(1)) Then
                    dTempPass = myreader(1)
                End If
                If Not IsDBNull(myreader(2)) Then
                    dPermPass = myreader(2)
                End If
                If Not IsDBNull(myreader(3)) Then
                    bDisabled = myreader(3)
                End If
                If Not IsDBNull(myreader(4)) Then
                    strPWHash = myreader(4)
                End If
                If Not IsDBNull(myreader(5)) Then
                    iNumFailedLogins = myreader(5)
                End If
                If Not IsDBNull(myreader(6)) Then
                    strFullName = myreader(6)
                End If
                If String.IsNullOrEmpty(strFullName) Then
                    strFullName = strID
                End If
                'If Not IsDBNull(myreader(7)) Then
                '    strUserType = myreader(7)
                'End If

                If iNumFailedLogins > iNumberOfFailedLoginsBeforeAccountLockout Then
                    bAccountLocked = True
                Else
                    bAccountLocked = False
                End If
                bVerify = BCrypt.Net.BCrypt.Verify(strPW, strPWHash)
                If bVerify Then
                    bValidUserPW = True
                    If bAccountLocked Then
                        bLoggedIn = False
                        iNumFailedLogins = iNumFailedLogins + 1
                    Else
                        bLoggedIn = True
                        iNumFailedLogins = 0
                    End If
                Else
                    bValidUserPW = False
                    bLoggedIn = False
                    iNumFailedLogins = iNumFailedLogins + 1
                End If
                If bRequirePasswordChange Then
                    ' temporary pw
                    If Now() > dTempPass Then
                        bUserCanChangePW = False
                    Else
                        bUserCanChangePW = True
                    End If
                Else
                    ' perm pw 
                    If Now() > dPermPass Then
                        bRequirePasswordChange = True

                        If Now() > dPermPass.AddDays(iDaysAfterPermPwExpiresUserCannotChangePW) Then
                            bUserCanChangePW = False
                        Else
                            bUserCanChangePW = True
                        End If
                    Else
                        bRequirePasswordChange = False
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If Not myreader Is Nothing AndAlso Not myreader.IsClosed Then
                myreader.Close()
            End If
        End Try

        If Not bLoginExist Then
            ' Invalid ID  bAccountLocked = false
            bLoggedIn = False
            bValidUserID = False
        End If

    End Sub

    Private Function VerifyNewPassword(ByVal newPW As String) As Boolean
        If String.IsNullOrEmpty(newPW) Or String.IsNullOrEmpty(strID) Then
            Return False
        End If
        If iNumberOfPriorPWNotbeUsed < 1 Then
            Return True
        End If

        Dim bRet As Boolean = True
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim strSQLAccount As String = "SELECT TOP " & iNumberOfPriorPWNotbeUsed.ToString & " Authentication_Value FROM  tbl_usr_PW_Log  WHERE UserID = @userid ORDER BY RowID DESC "
        Dim myreader As SqlDataReader = Nothing
        Dim acctparams As SqlParameter() = { _
            New SqlParameter("@userid", strID) _
            }
        Try
            myreader = CommonUtilsv2.GetDataReader(dbKey, strSQLAccount, CommandType.Text, acctparams)
            If myreader.HasRows Then
                While myreader.Read
                    If BCrypt.Net.BCrypt.Verify(newPW, myreader(0).ToString) Then
                        bRet = False
                        Exit While
                    End If
                End While
            Else
                bRet = True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If Not myreader Is Nothing AndAlso Not myreader.IsClosed Then
                myreader.Close()
            End If
        End Try
        Return bRet
    End Function

    Private Function UpdatePermanentPasswordExpiration(ByVal strUserID As String) As Boolean
        If String.Compare("CaseUser", strUserType, True) = 0 Then
            Return True
        End If
        Dim bRet As Boolean = False
        Dim strUpdate As String = "Update tbl_usr_Logins Set DateLastLoggedIn=@DateLastLoggedIn, DateLastUpdated=@DateLastUpdated "
        If bPermanentPasswordExpires Or bUserCanChangePW = False Then
        Else
            strUpdate = strUpdate + " ,PermPasswordExpire=@PermPasswordExpiresDate "
        End If
        strUpdate = strUpdate + " Where [UserID]=@Userid"

        Dim iDaysPermPwExpires As Integer = Webapps.Utils.ApplicationSettings.DaysPermanentPWExipres 'CommonUtilsv2.GetDaysPermanentPWExipres()
        Dim dPermPWExpiresDate As Date = Now.AddDays(iDaysPermPwExpires)
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        ' Dim strUpdate As String = "Update tbl_usr_Logins Set PermPasswordExpire=@PermPasswordExpiresDate, DateLastUpdated=@DateLastUpdated  Where [UserID]=@Userid"
        Dim params As SqlParameter() = { _
            New SqlParameter("@UserID", strUserID), _
            New SqlParameter("@DateLastLoggedIn", Now()), _
            New SqlParameter("@DateLastUpdated", Now()), _
            New SqlParameter("@PermPasswordExpiresDate", dPermPWExpiresDate) _
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strUpdate, CommandType.Text, params)
            bRet = True
        Catch ex As Exception
            bRet = False
            Throw New Exception("Error in updating permanent password expiration date.  The detail is: " & ex.ToString())
        End Try
        Return bRet
    End Function

    Private Function LogPasswordChange(ByVal strUserID As String, ByVal strNewPWHash As String) As Boolean
        Dim bRet As Boolean = False
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strUpdate As String = "INSERT INTO tbl_usr_PW_Log (UserID, Authentication_Value, RecordUpdatedBy, DateLastUpdated) VALUES (@Userid, @NewPWHash, @Userid, @DateLastUpdated) "
        Dim params As SqlParameter() = { _
            New SqlParameter("@UserID", strUserID), _
            New SqlParameter("@NewPWHash", strNewPWHash), _
            New SqlParameter("@DateLastUpdated", Now()) _
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strUpdate, CommandType.Text, params)
            bRet = True
        Catch ex As Exception
            bRet = False
            Throw New Exception("Error in logging password change.  The detail is: " & ex.ToString())
        End Try
        Return bRet
    End Function

    Private Function LogSuccesLogin(ByVal strUserID As String, ByVal strIP As String) As Boolean

        Dim bRet As Boolean = False
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strUpdate As String = "INSERT INTO tbl_WEB_Successful_Logins (UserID, IP_Address, DateAttempted ) VALUES (@Userid, @IPAddress, @DateAttempted) "
        Dim params As SqlParameter() = { _
            New SqlParameter("@UserID", strUserID), _
            New SqlParameter("@IPAddress", strIP), _
            New SqlParameter("@DateAttempted", Now()) _
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strUpdate, CommandType.Text, params)
            bRet = True
        Catch ex As Exception
            bRet = False
            Throw New Exception("Error in logging Success Login.  The detail is: " & ex.ToString())
        End Try
        Return bRet
    End Function

    Private Function LogFailedLogin(ByVal strUserID As String, ByVal strIP As String) As Boolean
        Dim bRet As Boolean = False
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strUpdate As String = "INSERT INTO tbl_WEB_Failed_Logins (UserID, IP_Address, DateAttempted ) VALUES (@Userid, @IPAddress, @DateAttempted) "
        Dim params As SqlParameter() = { _
            New SqlParameter("@UserID", strUserID), _
            New SqlParameter("@IPAddress", strIP), _
            New SqlParameter("@DateAttempted", Now()) _
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strUpdate, CommandType.Text, params)
            bRet = True
        Catch ex As Exception
            bRet = False
            Throw New Exception("Error in logging failed Login.  The detail is: " & ex.ToString())
        End Try
        Return bRet
    End Function

    Private Function UpdateFailedLoginAttempts(ByVal strUserID As String, ByVal iFalileCount As Integer) As Boolean
        If String.Compare("CaseUser", strUserType, True) = 0 Then
            Return True
        End If
        Dim bRet As Boolean = False
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strUpdate As String = "Update tbl_usr_Logins Set NumFailedLogins=@NumFailedLogins, DateLastUpdated=@DateLastUpdated  Where [UserID]=@Userid"
        Dim params As SqlParameter() = { _
            New SqlParameter("@UserID", strUserID), _
            New SqlParameter("@NumFailedLogins", iFalileCount), _
            New SqlParameter("@DateLastUpdated", Now()) _
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strUpdate, CommandType.Text, params)
            bRet = True
        Catch ex As Exception
            bRet = False
            Throw New Exception("Error in updating failed login attempt count.  The detail is: " & ex.ToString())
        End Try
        Return bRet
    End Function

End Class
