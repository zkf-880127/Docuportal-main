Imports System.Data.SqlClient
Imports System.Data
Imports Webapps.Utils

Namespace Webapps.Utils

    Public NotInheritable Class ApplicationSettings
        'declare datamembers
        Public Shared RequirePWChange As Boolean = True
        Public Shared SendErrorEmail As Boolean = True
        Public Shared DaysPermanentPWExipres As Integer
        Public Shared HoursTempPWExpires As Integer
        Public Shared DaysAfterPWExpirationPWCannotBeChanged As Integer
        Public Shared NumberOfFailedLoginsBeforeAccountLockout As Integer = 5
        Public Shared NumberOfPriorPWsNotBeUsed As Integer
        Public Shared SiteURL As String = ""
        Public Shared SiteTitle As String = ""
        Public Shared ApplicationSourceEmail As String
        Public Shared UserAccountNoticeEmails As String = ""
        Public Shared ClientUserAccountNoticeEmails As String = ""
        Public Shared ClaimsUpdateNoticeEmails As String = ""
        Public Shared ClaimsADRBCCEmails As String = ""
        Public Shared ErrorNoticeEmails As String = ""
        '   Public Shared CaseLoginNoticeEmails As String = ""
        Public Shared ContactUSEmails As String = ""
        Public Shared MailRelay As String = ""
        Public Shared EmailHost As String = ""
        Public Shared EmailPort As String = ""
        Public Shared EmailAccount As String = ""
        Public Shared EmailAccountPW As String = ""
        Public Shared Environment As String = ""
        Public Shared DisplayTimeZone As String = ""
        Public Shared Version As String = ""
        Public Shared CaseIdentifierPrefix As String = ""
        Public Shared Homepage As String = ""
        Public Shared PWMinimumLength As Integer = 8
        Public Shared PWPattern As String = "^(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[!@#$%^&*()., ?])[a-zA-Z0-9!@#$%^&*()., ?]{8,}$"
        Public Shared PWDescription As String = "Password must contain a minimum of 8 characters with at least one lower case letter, one upper case letter, one digit, and one following characters (! @ # $ % ^ & * ( ) . ,? and space)"


        Shared Sub New() 'LoadSettings()
            LoadSettings()
        End Sub

        Public Shared Sub LoadSettings()
            Dim dbKey1 As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim strSQL As String = "Select KeyName, KeyValue From tbl_Apps_AppSettings ORDER BY KeyName "
            Dim myReader As SqlDataReader = Nothing
            Dim strKeyname As String = Nothing
            UserAccountNoticeEmails = ""
            ClaimsUpdateNoticeEmails = ""
            ClaimsADRBCCEmails = ""
            ErrorNoticeEmails = ""
            ContactUSEmails = ""
            '    CaseLoginNoticeEmails = ""
            Try
                myReader = CommonUtilsv2.GetDataReader(dbKey1, strSQL, CommandType.Text)
                While myReader.Read
                    strKeyname = myReader.GetValue(0)
                    Select Case strKeyname
                        Case "RequirePWChange"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                Try
                                    RequirePWChange = Convert.ToBoolean(0 + myReader.GetValue(1))
                                Catch ex As Exception
                                    RequirePWChange = True
                                End Try
                            End If
                        Case "SendErrorEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                If String.Compare("1", myReader.GetValue(1), True) = 0 Then
                                    SendErrorEmail = True
                                Else
                                    SendErrorEmail = False
                                End If
                            Else
                                SendErrorEmail = False
                            End If
                        Case "DaysPermanentPWExipres"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                DaysPermanentPWExipres = myReader.GetValue(1)
                            End If
                        Case "DaysAfterPWExpirationPWCannotBeChanged"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                DaysAfterPWExpirationPWCannotBeChanged = myReader.GetValue(1)
                            End If
                        Case "HoursTempPWExpires"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                HoursTempPWExpires = myReader.GetValue(1)
                            End If
                        Case "NumberOfPriorPWsNotBeUsed"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                NumberOfPriorPWsNotBeUsed = myReader.GetValue(1)
                            End If
                        Case "NumberOfFailedLoginsBeforeAccountLockout"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                NumberOfFailedLoginsBeforeAccountLockout = myReader.GetValue(1)
                            End If
                        Case "SiteTitle"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                SiteTitle = myReader.GetValue(1)
                            End If
                        Case "SiteURL"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                SiteURL = myReader.GetValue(1)
                            End If
                        Case "MailRelay"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                MailRelay = myReader.GetValue(1)
                            End If
                        Case "EmailHost"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                EmailHost = myReader.GetValue(1)
                            End If
                        Case "EmailPort"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                EmailPort = myReader.GetValue(1)
                            End If
                        Case "EmailAccount"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                EmailAccount = myReader.GetValue(1)
                            End If
                        Case "EmailAccountPW"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                EmailAccountPW = myReader.GetValue(1)
                            End If
                        Case "Environment"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                Environment = myReader.GetValue(1)
                            End If
                        Case "ApplicationSourceEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                ApplicationSourceEmail = myReader.GetValue(1)
                            End If
                        Case "UserAccountNoticeEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                UserAccountNoticeEmails += myReader.GetValue(1) & ";"
                            End If
                        Case "ClientUserAccountNoticeEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                ClientUserAccountNoticeEmails += myReader.GetValue(1) & ";"
                            End If
                        Case "ClaimsUpdateNoticeEmails"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                ClaimsUpdateNoticeEmails += myReader.GetValue(1) & ";"
                            End If
                        Case "ClaimsADRBCCEmails"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                ClaimsADRBCCEmails += myReader.GetValue(1) & ";"
                            End If
                        Case "ErrorNoticeEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                ErrorNoticeEmails += myReader.GetValue(1) & ";"
                            End If
                        Case "ContactUSEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                ContactUSEmails += myReader.GetValue(1) & ";"
                            End If
                            'Case "CaseLoginNoticeEmail"
                            '    If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                            '        CaseLoginNoticeEmails += myReader.GetValue(1) & ";"
                            '    End If
                        Case "DisplayTimeZone"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                DisplayTimeZone = myReader.GetValue(1)
                            End If
                        Case "Version"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                Version = myReader.GetValue(1)
                            End If
                        Case "CaseIdentifierPrefix"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                CaseIdentifierPrefix = myReader.GetValue(1)
                            End If
                        Case "Homepage"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                Homepage = myReader.GetValue(1)
                            End If
                        Case "PWMinimumLength"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                PWMinimumLength = myReader.GetValue(1)
                            End If
                        Case "PWPattern"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                PWPattern = myReader.GetValue(1)
                            End If
                        Case "PWDescription"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                PWDescription = myReader.GetValue(1)
                            End If
                    End Select
                End While

                If Not String.IsNullOrEmpty(UserAccountNoticeEmails) Then
                    UserAccountNoticeEmails = UserAccountNoticeEmails.Remove(UserAccountNoticeEmails.Length - 1)
                End If
                If Not String.IsNullOrEmpty(ClientUserAccountNoticeEmails) Then
                    ClientUserAccountNoticeEmails = ClientUserAccountNoticeEmails.Remove(ClientUserAccountNoticeEmails.Length - 1)
                End If
                If Not String.IsNullOrEmpty(ClaimsUpdateNoticeEmails) Then
                    ClaimsUpdateNoticeEmails = ClaimsUpdateNoticeEmails.Remove(ClaimsUpdateNoticeEmails.Length - 1)
                End If
                If Not String.IsNullOrEmpty(ClaimsADRBCCEmails) Then
                    ClaimsADRBCCEmails = ClaimsADRBCCEmails.Remove(ClaimsADRBCCEmails.Length - 1)
                End If
                If Not String.IsNullOrEmpty(ErrorNoticeEmails) Then
                    ErrorNoticeEmails = ErrorNoticeEmails.Remove(ErrorNoticeEmails.Length - 1)
                End If
                If Not String.IsNullOrEmpty(ContactUSEmails) Then
                    ContactUSEmails = ContactUSEmails.Remove(ContactUSEmails.Length - 1)
                End If
                'If Not String.IsNullOrEmpty(CaseLoginNoticeEmails) Then
                '    CaseLoginNoticeEmails = CaseLoginNoticeEmails.Remove(CaseLoginNoticeEmails.Length - 1)
                'End If

            Catch ex As Exception
                ' Throw ex
                ' Dim msg As String = ex.ToString
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
            End Try
        End Sub

        Public Shared Sub ReLoadSettings()
            Dim dbKey1 As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim strSQL As String = "Select KeyName, KeyValue From tbl_Apps_AppSettings ORDER BY KeyName "
            Dim myReader As SqlDataReader = Nothing
            Dim strKeyname As String = Nothing
            Dim strUserAccountNoticeEmails As String = ""
            Dim strClientUserAccountNoticeEmails As String = ""
            Dim strClaimsUpdateNoticeEmailss As String = ""
            Dim strClaimsADRBCCEmails As String = ""
            Dim strErrorNoticeEmails As String = ""
            Dim strContactUSEmails As String = ""
            'Dim strCaseLoginNoticeEmails As String = ""
            Try
                myReader = CommonUtilsv2.GetDataReader(dbKey1, strSQL, CommandType.Text)
                While myReader.Read
                    strKeyname = myReader.GetValue(0)

                    Select Case strKeyname
                        Case "NumberOfPriorPWsNotBeUsed"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                NumberOfPriorPWsNotBeUsed = myReader.GetValue(1)
                            End If
                        Case "NumberOfFailedLoginsBeforeAccountLockout"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                NumberOfFailedLoginsBeforeAccountLockout = myReader.GetValue(1)
                            End If
                        Case "SendErrorEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                If String.Compare("1", myReader.GetValue(1), True) = 0 Then
                                    SendErrorEmail = True
                                Else
                                    SendErrorEmail = False
                                End If
                            Else
                                SendErrorEmail = False
                            End If

                        Case "ApplicationSourceEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                ApplicationSourceEmail = myReader.GetValue(1)
                            End If
                        Case "UserAccountNoticeEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                strUserAccountNoticeEmails += myReader.GetValue(1) & ";"
                            End If
                        Case "ClientUserAccountNoticeEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                strClientUserAccountNoticeEmails += myReader.GetValue(1) & ";"
                            End If
                        Case "ClaimsUpdateNoticeEmails"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                strClaimsUpdateNoticeEmailss += myReader.GetValue(1) & ";"
                            End If
                        Case "ClaimsADRBCCEmails"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                strClaimsADRBCCEmails += myReader.GetValue(1) & ";"
                            End If
                        Case "ErrorNoticeEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                strErrorNoticeEmails += myReader.GetValue(1) & ";"
                            End If
                        Case "ContactUSEmail"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                strContactUSEmails += myReader.GetValue(1) & ";"
                            End If
                            'Case "CaseLoginNoticeEmail"
                            '    If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                            '        strCaseLoginNoticeEmails += myReader.GetValue(1) & ";"
                            '    End If
                        Case "Homepage"
                            If Not myReader.GetValue(1) Is System.DBNull.Value AndAlso Not String.IsNullOrEmpty(myReader.GetValue(1)) Then
                                Homepage = myReader.GetValue(1)
                            End If
                        Case "PWMinimumLength"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                PWMinimumLength = myReader.GetValue(1)
                            End If
                        Case "PWPattern"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                PWPattern = myReader.GetValue(1)
                            End If
                        Case "PWDescription"
                            If Not myReader.GetValue(1) Is System.DBNull.Value Then
                                PWDescription = myReader.GetValue(1)
                            End If
                    End Select
                End While

                If Not String.IsNullOrEmpty(strUserAccountNoticeEmails) Then
                    UserAccountNoticeEmails = strUserAccountNoticeEmails.Remove(strUserAccountNoticeEmails.Length - 1)
                End If
                If Not String.IsNullOrEmpty(strUserAccountNoticeEmails) Then
                    ClientUserAccountNoticeEmails = strClientUserAccountNoticeEmails.Remove(strClientUserAccountNoticeEmails.Length - 1)
                End If
                If Not String.IsNullOrEmpty(strClaimsUpdateNoticeEmailss) Then
                    ClaimsUpdateNoticeEmails = strClaimsUpdateNoticeEmailss.Remove(strClaimsUpdateNoticeEmailss.Length - 1)
                End If
                If Not String.IsNullOrEmpty(strClaimsADRBCCEmails) Then
                    'ClaimsADRBCCEmails = ClaimsADRBCCEmails.Remove(ClaimsADRBCCEmails.Length - 1)
                End If
                If Not String.IsNullOrEmpty(strErrorNoticeEmails) Then
                    ErrorNoticeEmails = strErrorNoticeEmails.Remove(strErrorNoticeEmails.Length - 1)
                End If
                If Not String.IsNullOrEmpty(strContactUSEmails) Then
                    ContactUSEmails = strContactUSEmails.Remove(strContactUSEmails.Length - 1)
                End If
                'If Not String.IsNullOrEmpty(strCaseLoginNoticeEmails) Then
                '    CaseLoginNoticeEmails = strCaseLoginNoticeEmails.Remove(strCaseLoginNoticeEmails.Length - 1)
                'End If

            Catch ex As Exception
                ' Throw ex
                ' Dim msg As String = ex.ToString
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
            End Try
        End Sub

    End Class

End Namespace
