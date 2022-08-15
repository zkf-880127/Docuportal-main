Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Net.Mail
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web
Imports System.Collections.Generic
Imports System.Diagnostics

Namespace Webapps.Utils

    Public Class CommonUtilsv2
        Public Structure AppSettings
            Public RequirePWChange As Boolean
            Public DaysPermanentPWExipres As Integer
            Public HoursTempPWExpires As Integer
            Public DaysAfterPWExpirationPWCannotBeChanged As Integer
            Public NumberOfPriorPWsNotBeUsed As Integer
            Public SiteURL As String
            Public ApplicationSourceEmail As String
            Public UserAccountNoticeEmails As String
            Public ErrorNoticeEmails As String
            Public CaseLoginNoticeEmails As String
            Public SendErrorEmail As Boolean
            Public Sub Initialize()
                RequirePWChange = True
            End Sub
        End Structure

        Public Shared Function EscapeXML(ByVal strIn As String) As String
            Return System.Security.SecurityElement.Escape(strIn)
        End Function
        Public Shared Sub DCO_LogError(ByVal errorMsg As String, ByVal errorLocation As String, ByVal User As String, ByVal PageName As String)
            Dim datetimestamp As String = GetCurrentDateTimeDisplayString()
            Dim strRecepient As String = Webapps.Utils.ApplicationSettings.ErrorNoticeEmails ' GetErrorNoticeEmails()
            Dim strFrom As String = Webapps.Utils.ApplicationSettings.ApplicationSourceEmail 'GetApplicationSourceEmail()
            Dim strUserID As String = User
            'If strUserID = "" Then strUserID = HttpContext.Current.Request.UserHostAddress()
            Dim strSubject As String = "Web App Error in " & Webapps.Utils.ApplicationSettings.SiteTitle & " - " & GetEnvironmentAndHost() & "  Page: " & PageName
            Dim strBody As String = "<font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
            strBody += "<b>Web App Error in " & Webapps.Utils.ApplicationSettings.SiteTitle & "  " & GetEnvironmentAndHost() & "  Page: " & PageName & "</b>"
            strBody += "</font>"
            strBody += "<table width='100%'><font face='Verdana, Arial, Helvetica, sans-serif' size='2' color='#00658c'>"
            strBody += "<tr><td colspan=2><hr /></td></tr>"
            strBody += "<tr><td width='25%' align='right'><u>Item</u>&nbsp;&nbsp;</td><td width='75%' align='left'><u>Value</u></td></tr>"
            strBody += "<tr><td align='right'>Date / Time:&nbsp;&nbsp;</td><td align='left'>" + datetimestamp + "</td></tr>"
            strBody += "<tr><td align='right'>Page:&nbsp;&nbsp;</td><td align='left'>" + PageName + "</td></tr>"
            strBody += "<tr><td align='right'>Server:&nbsp;&nbsp;</td><td align='left'>" + GetHostSiteID() + "</td></tr>"
            strBody += "<tr><td align='right'>Stated Location:&nbsp;&nbsp;</td><td align='left'>" + errorLocation + "</td></tr>"
            strBody += "<tr><td align='right'>User:&nbsp;&nbsp;</td><td align='left'>" + strUserID + "</td></tr>"
            strBody += "<tr><td align='right'>Error description:&nbsp;&nbsp;</td><td align='left'>" + errorMsg + "</td></tr>"
            strBody += "<tr><td colspan=2><hr /></td></tr>"
            strBody += "</font></table>"
            CommonUtilsv2.SendEMail(strFrom, strRecepient, strBody, strSubject)
        End Sub

        Public Shared Sub CreateErrorLog(ByVal errorMsg As String, ByVal errorLocation As String, ByVal User As String, ByVal PageName As String, ByVal IP As String, Optional ByVal SendEmailOverride As Integer = 0)
            Dim bSendEmail As Boolean = Webapps.Utils.ApplicationSettings.SendErrorEmail
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
            Dim strSQL As String = "Insert Into tbl_Web_ErrorLog (SiteID, IPAddress, UserID, ErrorMessage, ErrorDateTime, ErrorLocation, ErrorPage) " & _
                                   " Values (@SiteID, @IPAddress, @UserID, @ErrorMessage, @ErrorDateTime, @ErrorLocation, @ErrorPage)"
            If User = "" Then

                User = IP
                'User = HttpContext.Current.Request.UserHostAddress()
            End If
            Dim params As SqlParameter() = { _
                New SqlParameter("@SiteID", GetHostSiteID()), _
                New SqlParameter("@IPAddress", IP), _
                New SqlParameter("@UserID", User), _
                New SqlParameter("@ErrorMessage", errorMsg), _
                New SqlParameter("@ErrorDateTime", DateTime.UtcNow), _
                New SqlParameter("@ErrorLocation", errorLocation), _
                New SqlParameter("@ErrorPage", PageName) _
                }

            Try
                RunNonQuery(dbKey, strSQL, CommandType.Text, params)
                If SendEmailOverride = 0 Then
                    If bSendEmail Then
                        If User = "" Then
                            DCO_LogError(errorMsg, errorLocation, IP, PageName)
                        Else
                            DCO_LogError(errorMsg, errorLocation, User, PageName)
                        End If
                    End If
                ElseIf SendEmailOverride = 1 Then
                    If User = "" Then
                        DCO_LogError(errorMsg, errorLocation, IP, PageName)
                    Else
                        DCO_LogError(errorMsg, errorLocation, User, PageName)
                    End If
                ElseIf SendEmailOverride = 2 Then
                    'Do not send e-mail
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        Public Shared Sub CreateErrorLog(ByVal errorLocation As String, ByRef exp As Exception, ByVal User As String, ByVal PageName As String, ByVal IP As String, Optional ByVal SendEmailOverride As Integer = 0)
            Dim bSendEmail As Boolean = Webapps.Utils.ApplicationSettings.SendErrorEmail
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
            Dim strSQL As String = "Insert Into tbl_Web_ErrorLog (SiteID, IPAddress, UserID, ErrorMessage, ErrorDateTime, ErrorLocation, ErrorPage) " & _
                                   " Values (@SiteID, @IPAddress, @UserID, @ErrorMessage, @ErrorDateTime, @ErrorLocation, @ErrorPage)"
            Dim errorMsg As String = String.Empty
            errorMsg = GetErrorDetails(errorLocation, exp)
            If User = "" Then

                User = IP
                'User = HttpContext.Current.Request.UserHostAddress()
            End If
            Dim params As SqlParameter() = { _
                New SqlParameter("@SiteID", GetHostSiteID()), _
                New SqlParameter("@IPAddress", IP), _
                New SqlParameter("@UserID", User), _
                New SqlParameter("@ErrorMessage", errorMsg), _
                New SqlParameter("@ErrorDateTime", DateTime.UtcNow), _
                New SqlParameter("@ErrorLocation", errorLocation), _
                New SqlParameter("@ErrorPage", PageName) _
                }

            Try
                RunNonQuery(dbKey, strSQL, CommandType.Text, params)
                If SendEmailOverride = 0 Then
                    If bSendEmail Then
                        If User = "" Then
                            DCO_LogError(errorMsg, "", IP, PageName)
                        Else
                            DCO_LogError(errorMsg, errorLocation, User, PageName)
                        End If
                    End If
                ElseIf SendEmailOverride = 1 Then
                    If User = "" Then
                        DCO_LogError(errorMsg, errorLocation, IP, PageName)
                    Else
                        DCO_LogError(errorMsg, errorLocation, User, PageName)
                    End If
                ElseIf SendEmailOverride = 2 Then
                    'Do not send e-mail
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Sub



        Public Shared Sub CreateErrorLog(ByRef exp As Exception, ByVal User As String, ByVal PageName As String, ByVal IP As String, Optional ByVal SendEmailOverride As Integer = 0)
            Dim bSendEmail As Boolean = Webapps.Utils.ApplicationSettings.SendErrorEmail
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
            Dim strSQL As String = "Insert Into tbl_Web_ErrorLog (SiteID, IPAddress, UserID, ErrorMessage, ErrorDateTime, ErrorLocation, ErrorPage) " & _
                                   " Values (@SiteID, @IPAddress, @UserID, @ErrorMessage, @ErrorDateTime, @ErrorLocation, @ErrorPage)"
            Dim errorMsg As String = String.Empty
            errorMsg = GetErrorDetails(exp)
            Dim errorLocation As String = String.Empty
            If User = "" Then

                User = IP
                'User = HttpContext.Current.Request.UserHostAddress()
            End If
            Dim params As SqlParameter() = { _
                New SqlParameter("@SiteID", GetHostSiteID()), _
                New SqlParameter("@IPAddress", IP), _
                New SqlParameter("@UserID", User), _
                New SqlParameter("@ErrorMessage", errorMsg), _
                New SqlParameter("@ErrorDateTime", DateTime.UtcNow), _
                New SqlParameter("@ErrorLocation", errorLocation), _
                New SqlParameter("@ErrorPage", PageName) _
                }

            Try
                RunNonQuery(dbKey, strSQL, CommandType.Text, params)
                If SendEmailOverride = 0 Then
                    If bSendEmail Then
                        If User = "" Then
                            DCO_LogError(errorMsg, "", IP, PageName)
                        Else
                            DCO_LogError(errorMsg, errorLocation, User, PageName)
                        End If
                    End If
                ElseIf SendEmailOverride = 1 Then
                    If User = "" Then
                        DCO_LogError(errorMsg, errorLocation, IP, PageName)
                    Else
                        DCO_LogError(errorMsg, errorLocation, User, PageName)
                    End If
                ElseIf SendEmailOverride = 2 Then
                    'Do not send e-mail
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
        Public Shared Function ValidateForm(ByVal ctrlParent As Control, Optional ByVal tbErrMsg As String = "Invalid Input", Optional ByVal ddErrMsg As String = "Invalid Selection", _
                                            Optional ByVal tbError As Integer = 0, Optional ByVal ddError As Integer = 0) As String
            Dim retError As StringBuilder = New StringBuilder
            Dim maxLength As Integer = 0
            Dim errorString = ""
            Dim errorLocation = "Server-side form validation"
            Dim userID As String = ""
            If Not HttpContext.Current.Session("User") = Nothing Then
                userID = HttpContext.Current.Session("User")
            End If
            For Each ctrl As Control In ctrlParent.Controls
                If TypeOf ctrl Is TextBox Then
                    If Not TryCast(ctrl, TextBox).MaxLength = Nothing Then
                        maxLength = CType(ctrl, TextBox).MaxLength
                    Else
                        maxLength = 50
                    End If
                    If Not ValidateFormFields(CType(ctrl, TextBox).Text, DataTypes.String, True, True, True, maxLength) Then
                        tbError += 1
                        If Not String.IsNullOrEmpty(ctrl.ID.ToString) Then
                            If ctrl.ID.ToString.IndexOf("Password") > 0 Then
                                errorString = "Invalid textbox input for " + ctrl.ID + "."
                                CreateErrorLog(errorString, errorLocation, userID, System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress(), 2)
                            Else
                                errorString = "Invalid textbox input for " + ctrl.ID + ": " + CType(ctrl, TextBox).Text
                                CreateErrorLog(errorString, errorLocation, userID, System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress(), 2)
                            End If
                        End If
                    End If
                ElseIf TypeOf ctrl Is DropDownList Then
                    If Not CType(ctrl, DropDownList).SelectedItem Is Nothing Then
                        If Not ValidateFormFields(CType(ctrl, DropDownList).SelectedItem.ToString, DataTypes.String, True, True, True) Then
                            ddError += 1
                            errorString = "Invalid dropdown selecteditem for " + ctrl.ID + ": " + CType(ctrl, DropDownList).SelectedItem.ToString
                            CreateErrorLog(errorString, errorLocation, userID, System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress(), 2)
                        End If
                    End If
                    If Not ValidateFormFields(CType(ctrl, DropDownList).SelectedValue.ToString, DataTypes.String, True, True, True) Then
                        ddError += 1
                        errorString = "Invalid dropdown selectedvalue for " + ctrl.ID + ": " + CType(ctrl, DropDownList).SelectedValue.ToString
                        CreateErrorLog(errorString, errorLocation, userID, System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress(), 2)
                    End If
                Else
                    If ctrl.Controls.Count > 0 Then
                        ValidateForm(ctrl, , , tbError, ddError)
                    End If
                End If
            Next
            If tbError > 0 Then
                retError.Append(tbErrMsg + "<br />")
            End If
            If ddError > 0 Then
                retError.Append(ddErrMsg + "<br />")
            End If
            Return retError.ToString
        End Function

        Public Shared Sub SendEMail(ByVal fromEmail As String, ByVal toRecepients As String, ByVal message As String, ByVal subject As String, Optional ByVal ccRecepients As String = Nothing, Optional ByVal StrTimess As Integer = 1)
            Dim EMAIL_HOST As String = Webapps.Utils.ApplicationSettings.EmailHost ' smtpout.secureserver.net
            Dim EMAIL_PORT As Integer = Webapps.Utils.ApplicationSettings.EmailPort  '80

            Dim bSuccess = False
            Try
                If (fromEmail.Length > 0 And toRecepients.Length > 0 And message.Length > 0) Then
                    Dim Authentication As New Net.NetworkCredential(Webapps.Utils.ApplicationSettings.EmailAccount, Webapps.Utils.ApplicationSettings.EmailAccountPW)
                    ' Dim Authentication As New Net.NetworkCredential("wallyl@jmsassoc.com", "Wspyh@#&08")
                    Dim EMail As New MailMessage
                    Dim SMTPServer As New SmtpClient
                    Dim arrEmailTo As Array = Split(toRecepients, ";")
                    Dim arrEmailCC As Array = Split(ccRecepients, ";")
                    For i As Integer = 0 To arrEmailTo.Length - 1
                        EMail.To.Add(arrEmailTo(i))
                    Next i
                    If Not String.IsNullOrEmpty(ccRecepients) Then
                        For i As Integer = 0 To arrEmailCC.Length - 1
                            EMail.CC.Add(arrEmailCC(i))
                        Next i
                    End If
                    ' EMail.From = New System.Net.Mail.MailAddress(fromEmail)
                    EMail.From = New System.Net.Mail.MailAddress(Webapps.Utils.ApplicationSettings.ApplicationSourceEmail)
                    'EMail.From = New System.Net.Mail.MailAddress("wallyl@jmsassoc.com")

                    EMail.IsBodyHtml = True
                    EMail.Subject = subject
                    EMail.Body = message
                    SMTPServer.Host = EMAIL_HOST '"smtpout.secureserver.net" 
                    SMTPServer.Port = EMAIL_PORT '80
                    SMTPServer.EnableSsl = True
                    SMTPServer.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network
                    'SMTPServer.UseDefaultCredentials = False
                    SMTPServer.Credentials = Authentication
                    While StrTimess < 20
                        Try
                            SMTPServer.Send(EMail)
                            StrTimess = 20
                        Catch ex As Exception
                            StrTimess = StrTimess + 1
                            System.Threading.Thread.Sleep(10000)

                            If (StrTimess = 20) Then

                                Dim errString As String = "Host; " & EMAIL_HOST & " Port: " & EMAIL_PORT & " From: " & Webapps.Utils.ApplicationSettings.ApplicationSourceEmail & ". Error: " & ex.Message
                                Dim errLocation As String = "CommonUtilsv2.SendEMail()"
                                CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress(), 2)
                                Throw New Exception(errString)
                            End If
                        End Try
                    End While
                End If
            Catch ex As Exception
                Dim errString As String = "Host; " & EMAIL_HOST & " Port: " & EMAIL_PORT & " From: " & Webapps.Utils.ApplicationSettings.ApplicationSourceEmail & ". Error: " & ex.Message
                Dim errLocation As String = "CommonUtilsv2.SendEMail()"
                CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress(), 2)
                Throw New Exception(errString)
            End Try

        End Sub

        Public Shared Sub SendEMailBCC(ByVal fromEmail As String, ByVal toRecepients As String, ByVal message As String, ByVal subject As String, Optional ByVal ccBCCRecepients As String = Nothing, Optional ByVal StrTimess As Integer = 1)
            Dim EMAIL_HOST As String = Webapps.Utils.ApplicationSettings.EmailHost  ' System.Configuration.ConfigurationManager.AppSettings("MailRelay")
            Dim EMAIL_PORT As Integer = Webapps.Utils.ApplicationSettings.EmailPort  '587

            Try
                If (fromEmail.Length > 0 And toRecepients.Length > 0 And message.Length > 0) Then
                    Dim Authentication As New Net.NetworkCredential(Webapps.Utils.ApplicationSettings.EmailAccount, Webapps.Utils.ApplicationSettings.EmailAccountPW)
                    Dim EMail As New MailMessage
                    Dim SMTPServer As New SmtpClient
                    Dim arrEmailTo As Array = Split(toRecepients, ";")
                    Dim arrEmailBCC As Array = Split(ccBCCRecepients, ";")
                    For i As Integer = 0 To arrEmailTo.Length - 1
                        EMail.To.Add(arrEmailTo(i))
                    Next i
                    If Not String.IsNullOrEmpty(ccBCCRecepients) Then
                        For i As Integer = 0 To arrEmailBCC.Length - 1
                            EMail.Bcc.Add(arrEmailBCC(i))
                        Next i
                    End If
                    ' EMail.From = New System.Net.Mail.MailAddress(fromEmail)
                    EMail.From = New System.Net.Mail.MailAddress(Webapps.Utils.ApplicationSettings.ApplicationSourceEmail)
                    EMail.IsBodyHtml = True
                    EMail.Subject = subject
                    EMail.Body = message
                    SMTPServer.Host = EMAIL_HOST '"smtpout.secureserver.net" 
                    SMTPServer.Port = EMAIL_PORT '80
                    '  SMTPServer.EnableSsl = False
                    SMTPServer.EnableSsl = True
                    SMTPServer.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network
                    SMTPServer.UseDefaultCredentials = False
                    SMTPServer.Credentials = Authentication
                    While StrTimess < 20
                        Try
                            SMTPServer.Send(EMail)
                            StrTimess = 20
                        Catch ex As Exception
                            StrTimess = StrTimess + 1
                            System.Threading.Thread.Sleep(10000)

                            If (StrTimess = 20) Then

                                Dim errString As String = "Host; " & EMAIL_HOST & " Port: " & EMAIL_PORT & " From: " & Webapps.Utils.ApplicationSettings.ApplicationSourceEmail & ". Error: " & ex.Message
                                Dim errLocation As String = "CommonUtilsv2.SendEMail()"
                                CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress(), 2)
                                Throw New Exception(errString)
                            End If
                        End Try
                    End While
                End If
            Catch ex As Exception
                Dim errString As String = "Host; " & EMAIL_HOST & " Port: " & EMAIL_PORT & " From: " & Webapps.Utils.ApplicationSettings.ApplicationSourceEmail & ". Error: " & ex.Message
                Dim errLocation As String = "CommonUtilsv2.SendEMail()"
                CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress(), 2)
                Throw New Exception(errString)
            End Try
        End Sub


        Public Shared Sub SendEMail_Relay(ByVal fromEmail As String, ByVal toRecepients As String, ByVal message As String, ByVal subject As String, Optional ByVal ccRecepients As String = Nothing)
            Dim EMAIL_HOST As String = Webapps.Utils.ApplicationSettings.MailRelay  ' System.Configuration.ConfigurationManager.AppSettings("MailRelay")

            If (fromEmail.Length > 0 And toRecepients.Length > 0 And message.Length > 0) Then

                'Dim Authentication As New Net.NetworkCredential("login name/email", "password")
                Dim EMail As New MailMessage
                Dim SMTPServer As New SmtpClient
                Dim arrEmailTo As Array = Split(toRecepients, ";")
                Dim arrEmailCC As Array = Split(ccRecepients, ";")
                For i As Integer = 0 To arrEmailTo.Length - 1
                    EMail.To.Add(arrEmailTo(i))
                Next i

                If Not String.IsNullOrEmpty(ccRecepients) Then

                    For i As Integer = 0 To arrEmailCC.Length - 1
                        EMail.CC.Add(arrEmailCC(i))
                    Next i

                End If

                ' EMail.From = New System.Net.Mail.MailAddress(fromEmail)
                EMail.From = New System.Net.Mail.MailAddress(Webapps.Utils.ApplicationSettings.ApplicationSourceEmail)
                EMail.IsBodyHtml = True
                EMail.Subject = subject
                EMail.Body = message
                SMTPServer.Host = EMAIL_HOST
                SMTPServer.Port = 25
                '       SMTPServer.Credentials = Authentication
                SMTPServer.Send(EMail)
            End If

        End Sub

        Public Shared Sub SendEMailBCC_Relay(ByVal fromEmail As String, ByVal toRecepients As String, ByVal message As String, ByVal subject As String, Optional ByVal ccBCCRecepients As String = Nothing)
            Dim EMAIL_HOST As String = Webapps.Utils.ApplicationSettings.MailRelay

            If (fromEmail.Length > 0 And toRecepients.Length > 0 And message.Length > 0) Then

                'Dim Authentication As New Net.NetworkCredential("login name/email", "password")
                Dim EMail As New MailMessage
                Dim SMTPServer As New SmtpClient
                Dim arrEmailTo As Array = Split(toRecepients, ";")
                Dim arrEmailBCC As Array = Split(ccBCCRecepients, ";")
                For i As Integer = 0 To arrEmailTo.Length - 1
                    EMail.To.Add(arrEmailTo(i))
                Next i

                If Not String.IsNullOrEmpty(ccBCCRecepients) Then

                    For i As Integer = 0 To arrEmailBCC.Length - 1
                        If Not String.IsNullOrEmpty(arrEmailBCC(i)) Then
                            EMail.Bcc.Add(arrEmailBCC(i))
                        End If
                    Next i

                End If

                EMail.From = New System.Net.Mail.MailAddress(fromEmail)
                EMail.IsBodyHtml = True
                EMail.Subject = subject
                EMail.Body = message
                SMTPServer.Host = EMAIL_HOST
                SMTPServer.Port = 25
                '       SMTPServer.Credentials = Authentication
                SMTPServer.Send(EMail)
            End If

        End Sub

        Public Shared Sub SendEmailAttach(ByVal fromEmail As String, ByVal toRecepients As String, ByVal message As String, ByVal subject As String, Optional ByVal ccBCCRecepients As String = Nothing, Optional ByVal AttachmentFile As String = Nothing, Optional ByVal MemAttach() As Attachment = Nothing)
            Dim EMAIL_HOST As String = Webapps.Utils.ApplicationSettings.MailRelay

            If (fromEmail.Length > 0 And toRecepients.Length > 0 And message.Length > 0) Then

                Dim EMail As New MailMessage
                Dim SMTPServer As New SmtpClient
                Dim arrEmailTo As Array = Split(toRecepients, ";")
                Dim arrEmailBCC As Array = Split(ccBCCRecepients, ";")

                Dim arrAttach As Array = Split(AttachmentFile, ";")

                For i As Integer = 0 To arrEmailTo.Length - 1
                    EMail.To.Add(arrEmailTo(i))
                Next i

                If Not String.IsNullOrEmpty(ccBCCRecepients) Then

                    For i As Integer = 0 To arrEmailBCC.Length - 1
                        If Not String.IsNullOrEmpty(arrEmailBCC(i)) Then
                            EMail.Bcc.Add(arrEmailBCC(i))
                        End If
                    Next i

                End If
                If Not String.IsNullOrEmpty(AttachmentFile) Then
                    For Each strFileName As String In arrAttach
                        Dim aData As Attachment = New Attachment(strFileName)
                        EMail.Attachments.Add(aData)
                    Next
                End If

                If Not MemAttach Is Nothing Then
                    For Each attFile As Attachment In MemAttach
                        EMail.Attachments.Add(attFile)
                    Next

                End If

                EMail.From = New System.Net.Mail.MailAddress(fromEmail)
                EMail.IsBodyHtml = True
                EMail.Subject = subject
                EMail.Body = message
                SMTPServer.Host = EMAIL_HOST
                SMTPServer.Port = 25
                SMTPServer.Send(EMail)
            End If

        End Sub

        Public Shared Function GetADRPage(ByVal iDocID As Integer) As String
            Dim strADRLandingPage As String = ""
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim myConn As New SqlClient.SqlConnection(dbKey)

            Dim strSQL As String = " SELECT ClaimNumber FROM [tbl_Web_Claim_ADR_Documents] WHERE ID = @ID  "

            Dim myComm As New SqlCommand(strSQL, myConn)
            myComm.Parameters.AddWithValue("@ID", iDocID)
            Try
                myConn.Open()
                Dim myReader As SqlDataReader = myComm.ExecuteReader()
                While myReader.Read
                    If Not IsDBNull(myReader(myReader.GetOrdinal("ClaimNumber"))) Then
                        Dim strClaimNo As String = myReader.GetValue(myReader.GetOrdinal("ClaimNumber"))
                        'strADRLandingPage =  Webapps.Utils.ApplicationSettings.SiteURL & "/ClaimRoom-ManageClaim.aspx?ClaimNumber=" & strClaimNo
                        strADRLandingPage = "ClaimRoom-ManageClaimADR.aspx?ClaimNumber=" & strClaimNo
                    End If
                End While
                myReader.Close()
            Catch ex As Exception
                Dim errString As String = ex.Message
            Finally
                myComm.Dispose()
                myConn.Close()
            End Try

            Return strADRLandingPage
        End Function

        Public Shared Function FormatAPCurrency(ByVal apCurrency As String) As String
            If IsNumeric(apCurrency) = False Then
                If apCurrency Is System.DBNull.Value Then
                    apCurrency = ""
                End If
            Else
                Return FormatCurrency(apCurrency, 2)
            End If
            Return apCurrency
        End Function

        Public Shared Function GetDataReader(ByVal dbCNStr As String, ByVal strSQL As String, ByVal cmdType As CommandType) As IDataReader
            Dim aReader As SqlDataReader
            Dim sqlCnn As New SqlClient.SqlConnection(dbCNStr)
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            sqlCmd.CommandType = cmdType
            aReader = Nothing
            Try
                sqlCnn.Open()
                aReader = sqlCmd.ExecuteReader(CommandBehavior.CloseConnection)
                Return aReader
            Catch ex As Exception
                If (Not aReader Is Nothing AndAlso Not aReader.IsClosed) Then
                    aReader.Close()
                    sqlCmd.Dispose()
                    sqlCnn.Close()
                End If
                Throw ex
            End Try
        End Function

        Public Shared Function GetDataReader(ByVal dbCNStr As String, ByVal strSQL As String, ByVal cmdType As CommandType, ByVal parameters As SqlParameter()) As IDataReader
            Dim aReader As SqlDataReader
            Dim sqlCnn As New SqlClient.SqlConnection(dbCNStr)
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            sqlCmd.CommandType = cmdType

            For Each parameter As SqlParameter In parameters

                sqlCmd.Parameters.Add(parameter)

            Next

            aReader = Nothing
            Try
                sqlCnn.Open()
                aReader = sqlCmd.ExecuteReader(CommandBehavior.CloseConnection)
                Return aReader
            Catch ex As Exception
                If (Not aReader Is Nothing AndAlso Not aReader.IsClosed) Then
                    aReader.Close()
                    sqlCmd.Dispose()
                    sqlCnn.Close()
                End If
                Throw ex
            End Try
        End Function

        Public Shared Function GetDataSet(ByVal dbCNStr As String, ByVal strSQL As String, ByVal cmdType As CommandType) As DataSet
            Dim sqlCnn As New SqlClient.SqlConnection(dbCNStr)
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim sqlDa As New SqlClient.SqlDataAdapter
            Dim sqlDs As New DataSet
            sqlCmd.CommandTimeout = 0
            sqlCmd.CommandType = cmdType

            Try
                sqlDa.SelectCommand = sqlCmd
                sqlDa.Fill(sqlDs)
                Return sqlDs
            Catch ex As Exception
                Throw ex
            Finally
                sqlDs.Dispose()
                sqlDa.Dispose()
                sqlCmd.Dispose()
                sqlCnn.Close()
                sqlCnn.Dispose()
            End Try
        End Function

        Public Shared Function GetDataSet(ByVal dbCNStr As String, ByVal strSQL As String, ByVal cmdType As CommandType, ByVal parameters As SqlParameter()) As DataSet
            Dim sqlCnn As New SqlClient.SqlConnection(dbCNStr)
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim sqlDa As New SqlClient.SqlDataAdapter
            Dim sqlDs As New DataSet
            sqlCmd.CommandTimeout = 0

            sqlCmd.CommandType = cmdType

            For Each parameter As SqlParameter In parameters

                sqlCmd.Parameters.Add(parameter)

            Next

            Try
                sqlDa.SelectCommand = sqlCmd
                sqlDa.Fill(sqlDs)
                Return sqlDs
            Catch ex As Exception
                Throw ex
            Finally
                sqlDs.Dispose()
                sqlDa.Dispose()
                sqlCmd.Dispose()
                sqlCnn.Close()
                sqlCnn.Dispose()
            End Try
        End Function

        Public Shared Function RunScalarQuery(ByVal dbCNStr As String, ByVal strSQL As String, ByVal cmdType As CommandType) As String
            Dim ReturnStr As String
            Dim sqlCnn As New SqlClient.SqlConnection(dbCNStr)
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            sqlCmd.CommandType = cmdType

            Try
                sqlCnn.Open()
                Dim Stra As Object = sqlCmd.ExecuteScalar()
                If Not IsNothing(Stra) Then
                    ReturnStr = Stra.ToString()
                Else
                    ReturnStr = Nothing
                End If
                Return ReturnStr
            Catch ex As Exception
                sqlCmd.Dispose()
                sqlCnn.Close()
                Throw ex
            Finally
                sqlCmd.Dispose()
                sqlCnn.Close()
                sqlCnn.Dispose()
            End Try
        End Function

        Public Shared Function RunScalarQuery(ByVal dbCNStr As String, ByVal strSQL As String, ByVal cmdType As CommandType, ByVal parameters As SqlParameter()) As String
            Dim ReturnStr As String
            Dim sqlCnn As New SqlClient.SqlConnection(dbCNStr)
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            sqlCmd.CommandType = cmdType

            If (parameters IsNot Nothing) Then
                For Each parameter As SqlParameter In parameters

                    sqlCmd.Parameters.Add(parameter)

                Next
            End If

            Try
                sqlCnn.Open()
                Dim Stra As Object = sqlCmd.ExecuteScalar()
                If Not IsNothing(Stra) Then
                    ReturnStr = Stra.ToString()
                Else
                    ReturnStr = Nothing
                End If
                Return ReturnStr
            Catch ex As Exception
                sqlCmd.Dispose()
                sqlCnn.Close()
                Throw ex
            Finally
                sqlCmd.Dispose()
                sqlCnn.Close()
                sqlCnn.Dispose()
            End Try
        End Function

        Public Shared Sub RunNonQuery(ByVal dbCNStr As String, ByVal strSQL As String, ByVal cmdType As CommandType)
            Dim sqlCnn As New SqlClient.SqlConnection(dbCNStr)
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            sqlCmd.CommandType = cmdType

            Try
                sqlCnn.Open()
                sqlCmd.ExecuteNonQuery()
            Catch ex As Exception
                sqlCmd.Dispose()
                sqlCnn.Close()
                Throw ex
            Finally
                sqlCmd.Dispose()
                sqlCnn.Close()
                sqlCnn.Dispose()
            End Try
        End Sub

        Public Shared Sub RunNonQuery(ByVal dbCNStr As String, ByVal strSQL As String, ByVal cmdType As CommandType, ByVal parameters As SqlParameter())
            Dim sqlCnn As New SqlClient.SqlConnection(dbCNStr)
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            sqlCmd.CommandType = cmdType

            For Each parameter As SqlParameter In parameters

                sqlCmd.Parameters.Add(parameter)

            Next

            Try
                sqlCnn.Open()
                sqlCmd.ExecuteNonQuery()
            Catch ex As Exception
                sqlCmd.Dispose()
                sqlCnn.Close()
                Throw ex
            Finally
                sqlCmd.Dispose()
                sqlCnn.Close()
                sqlCnn.Dispose()
            End Try
        End Sub

        Public Shared Sub DefineDataTables(ByRef ds As DataSet)

            Dim result As Boolean

            For index As Integer = 0 To ds.Tables.Count - 1

                System.Diagnostics.Debug.WriteLine(ds.Tables(index).TableName) ' Current Table
                System.Diagnostics.Debug.WriteLine(ds.Tables(index).Columns(0).ColumnName)

                result = ds.Tables(index).Columns(0).ColumnName.EndsWith("_ID_NAME")

                If result Then
                    ds.Tables(index).TableName = ds.Tables(index).Columns(0).ColumnName.Replace("_ID_NAME", String.Empty)
                End If

            Next

        End Sub

        'Should always use this one and then the calling function to insert the first item
        Public Shared Sub PopulateDropDownBox(ByVal dbCNStr As String, ByVal strSQL As String, ByVal dd As DropDownList, ByVal dataValueField As String, ByVal dataTextField As String)
            Dim sqlDs As New DataSet
            Try
                'sqlDs = GetDataSet(dbCNStr, strSQL)
                'dd.ClearSelection()
                'dd.Items.Clear()
                'dd.SelectedIndex = -1
                sqlDs = GetDataSet(dbCNStr, strSQL, CommandType.Text)
                dd.DataSource = sqlDs.Tables(0)
                dd.DataTextField = dataTextField
                dd.DataValueField = dataValueField
                dd.DataBind()
                sqlDs.Tables.Clear()
            Catch ex As Exception
                Throw ex
            Finally
                sqlDs.Dispose()
            End Try
        End Sub

        Public Shared Sub LoadDropDownBox(ByVal dbCNStr As String, ByVal strSQL As String, ByVal dd As DropDownList, ByVal dataValueField As String, ByVal dataTextField As String, Optional ByVal parameters As SqlParameter() = Nothing)
            Dim sqlDs As New DataSet
            Try
                If parameters Is Nothing Then
                    sqlDs = GetDataSet(dbCNStr, strSQL, CommandType.Text)
                Else
                    sqlDs = GetDataSet(dbCNStr, strSQL, CommandType.Text, parameters)
                End If
                dd.DataSource = sqlDs.Tables(0)
                dd.DataTextField = dataTextField
                dd.DataValueField = dataValueField
                dd.DataBind()
                sqlDs.Tables.Clear()
            Catch ex As Exception
                Throw ex
            Finally
                sqlDs.Dispose()
            End Try
        End Sub



        Public Shared Function InsertAuditTrail(ByVal dbCNStr As String, ByVal action As String, ByVal fieldId As String, ByVal newValue As String, ByVal oldValue As String, ByVal keyID As String, ByVal subKey As String, ByVal strDescription As String, ByVal strDateTime As String, ByVal user As String, Optional ByVal roomName As String = "") As Boolean
            Dim bSuccess As Boolean = False
            Dim dbKey As String = dbCNStr
            Dim myConn As New SqlClient.SqlConnection(dbKey)
            Dim strParQuery As String = ""
            If roomName = "Preferences" Then
                strParQuery = "INSERT INTO tbl_web_Preferences_Audit_Trail(Action_type, Key_ID, Sub_Key, Record_Field_ID,  Old_Value, New_Value, Description, UpdatedDate, UpdatedBy)"
                strParQuery += " VALUES (@Action_type, @Key_ID, @Sub_Key, @Record_Field_ID, @Old_Value, @New_Value, @Description, @currentDateTime, @currentUser) "
            ElseIf roomName = "Report" Then
                strParQuery = "INSERT INTO tbl_web_Report_Audit_Trail(Action_type, Key_ID, Sub_Key, Record_Field_ID,  Old_Value, New_Value, Description, UpdatedDate, UpdatedBy)"
                strParQuery += " VALUES (@Action_type, @Key_ID, @Sub_Key, @Record_Field_ID, @Old_Value, @New_Value, @Description, @currentDateTime, @currentUser) "
            Else
                strParQuery = "INSERT INTO tbl_web_Audit_Trail(Action_type, Key_ID, Sub_Key, Record_Field_ID,  Old_Value, New_Value, Description, UpdatedDate, UpdatedBy)"
                strParQuery += " VALUES (@Action_type, @Key_ID, @Sub_Key, @Record_Field_ID, @Old_Value, @New_Value, @Description, @currentDateTime, @currentUser) "
            End If
            Dim myComm As New SqlCommand(strParQuery, myConn)

            Try
                myConn.Open()
                myComm.Parameters.AddWithValue("@Action_type", action)
                myComm.Parameters.AddWithValue("@Key_ID", keyID) ' example: ClaimNumber
                myComm.Parameters.AddWithValue("@Sub_Key", subKey)  '
                myComm.Parameters.AddWithValue("@Record_Field_ID", fieldId)
                myComm.Parameters.AddWithValue("@Old_Value", oldValue)
                myComm.Parameters.AddWithValue("@New_Value", newValue)
                myComm.Parameters.AddWithValue("@Description", strDescription)
                myComm.Parameters.AddWithValue("@currentDateTime", strDateTime)
                myComm.Parameters.AddWithValue("@currentUser", user)
                myComm.ExecuteNonQuery()
                myComm.Dispose()
                bSuccess = True
            Catch ex As Exception
                bSuccess = False
                Throw ex
            Finally
                myComm.Dispose()
                myConn.Close()
            End Try
            Return bSuccess
        End Function

        Public Enum DataTypes
            [Int]
            [Long]
            [Double]
            [String]
            [Date]
            [Boolean]
        End Enum

        Public Enum LogType
            [User] = 1
            [System] = 2          
        End Enum

        Public Shared Function Validate(ByRef obj As Object, ByVal datatype As DataTypes, ByVal checkForScript As Boolean, ByVal checkForSQL As Boolean, ByVal encodeToHtml As Boolean, Optional ByVal maxlength As Integer = 50) As Boolean

            If String.IsNullOrEmpty(obj) Then
                Return False
            End If

            Select Case datatype

                Case DataTypes.Int

                    ' Check to see if it is integer
                    Dim num As Int32
                    If Int32.TryParse(obj, num) Then
                        Return True
                    Else
                        Return False
                    End If

                Case DataTypes.Long

                    ' Check to see if it is integer
                    Dim num As Long
                    If Long.TryParse(obj, num) Then
                        Return True
                    Else
                        Return False
                    End If

                Case DataTypes.Double

                    ' Check to see if it is integer
                    Dim num As Double
                    If Double.TryParse(obj, num) Then
                        Return True
                    Else
                        Return False
                    End If

                Case DataTypes.String

                    ' Check to see if the string is a number
                    'If IsNumeric(obj) Then
                    '    Return False
                    'End If

                    ' Check to see if the string contain numbers
                    'If Regex.IsMatch(obj, "[0-9]+", RegexOptions.IgnoreCase) Then
                    '    Return False
                    'End If

                    If checkForScript Then

                        ' Check to see if it contains script tag
                        If ContainsScriptTag(obj) Then
                            'encode when encodeHTML=ture
                            'HttpUtility.HtmlEncode(obj) ' Always encode a script.
                            Return False
                        End If

                    End If

                    ' Check to see that it does not exceed max length
                    Dim length As Integer = obj.ToString().Length
                    If length > maxlength Then
                        Return False
                    End If

                    If checkForSQL Then

                        If ContainsSQLKeywordsAndSymbols(obj) Then
                            Return False
                        End If

                    End If

                    If encodeToHtml Then

                        ' Encode string to html, so that if it is a script that gets through it will not execute if encoded to html
                        HttpUtility.HtmlEncode(obj)

                    End If

                    Return True

                Case DataTypes.Date

                    ' Check to see if it is Date
                    If Not IsDateValid(obj) Then
                        Return False
                    End If

                    If Not IsDate(obj) Then
                        Return False
                    End If

                    If IsDateWithinBounds(obj) Then
                        Return True
                    Else
                        Return False
                    End If

                Case DataTypes.Boolean

                    ' Check to see if it is integer
                    Dim result As Boolean
                    If Boolean.TryParse(obj, result) Then
                        Return True
                    Else
                        Return False
                    End If

                Case Else

                    Return False

            End Select

        End Function

        Public Shared Function ValidateFormFields(ByRef obj As Object, ByVal datatype As DataTypes, ByVal checkForScript As Boolean, ByVal checkForSQL As Boolean, ByVal encodeToHtml As Boolean, Optional ByVal maxlength As Integer = 50) As Boolean
            Select Case datatype
                Case DataTypes.Int
                    ' Check to see if it is integer
                    Dim num As Int32
                    If Int32.TryParse(obj, num) Then
                        Return True
                    Else
                        Return False
                    End If
                Case DataTypes.Long
                    ' Check to see if it is long
                    Dim num As Long
                    If Long.TryParse(obj, num) Then
                        Return True
                    Else
                        Return False
                    End If
                Case DataTypes.Double
                    ' Check to see if it is double
                    Dim num As Double
                    If Double.TryParse(obj, num) Then
                        Return True
                    Else
                        Return False
                    End If
                Case DataTypes.String
                    If checkForScript Then
                        ' Check to see if it contains script tag
                        If ContainsScriptTag(obj) Then
                            Return False
                        End If

                    End If
                    ' Check to see that it does not exceed max length
                    Dim length As Integer = obj.ToString().Length
                    If length > maxlength Then
                        Return False
                    End If
                    If encodeToHtml Then
                        ' Encode string to html, so that if it is a script that gets through it will not execute if encoded to html
                        HttpUtility.HtmlEncode(obj)
                    End If
                    Return True
                Case DataTypes.Date
                    ' Check to see if it is Date
                    If Not IsDateValid(obj) Then
                        Return False
                    End If
                    If Not IsDate(obj) Then
                        Return False
                    End If
                    If IsDateWithinBounds(obj) Then
                        Return True
                    Else
                        Return False
                    End If
                Case DataTypes.Boolean
                    ' Check to see if it is integer
                    Dim result As Boolean
                    If Boolean.TryParse(obj, result) Then
                        Return True
                    Else
                        Return False
                    End If
                Case Else
                    Return False
            End Select
        End Function

        Public Shared Function ValidateFormFieldsIncludeAll(ByRef obj As Object, ByVal datatype As DataTypes, ByVal checkForScript As Boolean, ByVal checkForSQL As Boolean, ByVal encodeToHtml As Boolean, Optional ByVal maxlength As Integer = 50) As Boolean
            Select Case datatype
                Case DataTypes.Int
                    ' Check to see if it is integer
                    Dim num As Int32
                    If Int32.TryParse(obj, num) Then
                        Return True
                    Else
                        Return False
                    End If
                Case DataTypes.Long
                    ' Check to see if it is long
                    Dim num As Long
                    If Long.TryParse(obj, num) Then
                        Return True
                    Else
                        Return False
                    End If
                Case DataTypes.Double
                    ' Check to see if it is double
                    Dim num As Double
                    If Double.TryParse(obj, num) Then
                        Return True
                    Else
                        Return False
                    End If
                Case DataTypes.String
                    If checkForScript Then
                        ' Check to see if it contains script tag
                        If ContainsScriptTag(obj) Then
                            Return False
                        End If

                    End If
                    ' Check to see that it does not exceed max length
                    Dim length As Integer = obj.ToString().Length
                    If length > maxlength Then
                        Return False
                    End If
                    If checkForSQL Then
                        If ContainsSQLKeywordsAndSymbols(obj) Then
                            Return False
                        End If
                    End If
                    If encodeToHtml Then
                        ' Encode string to html, so that if it is a script that gets through it will not execute if encoded to html
                        HttpUtility.HtmlEncode(obj)
                    End If
                    Return True
                Case DataTypes.Date
                    ' Check to see if it is Date
                    If Not IsDateValid(obj) Then
                        Return False
                    End If
                    If Not IsDate(obj) Then
                        Return False
                    End If
                    If IsDateWithinBounds(obj) Then
                        Return True
                    Else
                        Return False
                    End If
                Case DataTypes.Boolean
                    ' Check to see if it is integer
                    Dim result As Boolean
                    If Boolean.TryParse(obj, result) Then
                        Return True
                    Else
                        Return False
                    End If
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function IsDateValid(ByVal obj As String) As Boolean

            Try
                DateTime.Parse(obj)
                Return True
            Catch e As FormatException
                Return False
            End Try

        End Function

        Private Shared Function IsDateWithinBounds(ByVal obj As String) As Boolean

            Dim dt As DateTime

            Try
                dt = DateTime.Parse(obj)
            Catch e As FormatException
                Return False
            End Try

            Dim dtMin As DateTime = DateTime.Parse("1/1/1753 12:00:00 AM") ' I did this because DateTime.MinValue doesn't always seem to return 1/1/1753 12:00:00 AM

            ' Check to see if it is greater than 1/1/1753 12:00:00 AM
            If dtMin.Date > dt.Date Then
                Return False
            End If

            ' Check to see if it is less than 12/31/9999 11:59:59 PM
            If DateTime.MaxValue.Date < dt.Date Then
                Return False
            End If

            Return True

        End Function

        Private Shared Function ContainsScriptTag(ByVal obj As String) As Boolean

            If Regex.IsMatch(obj, "<((?s).*)>|[<>]") Then
                Return True
            Else
                Return False
            End If

        End Function

        Private Shared Function ContainsSQLKeywordsAndSymbols(ByVal obj As String) As Boolean

            Dim total As Integer = 0
            Dim total2 As Integer = 0

            Dim SQLSymbols As String() = {"--", ";", "/*", "*/", "@@", "@"}
            Dim SQLKeywordsRegex As String = "\bchar\b|\bnchar\b|\bvarchar\b|\bnvarchar\b|\balter\b|\bbegin\b|\bcast\b|\bcreate\b|\bcursor\b|\bdeclare\b|\bdelete\b|\bdrop\b|\bend\b|\bexec\b|\bexecute\b|\bfetch\b|\binsert\b|\bkill\b|\bopen\b|\bselect\b|\bsys\b|\bsysobjects\b|\bsyscolumns\b|\btable\b|\bupdate\b"

            If Regex.IsMatch(obj, SQLKeywordsRegex, RegexOptions.IgnoreCase) Then
                total += 1
            End If

            For Each Item As String In SQLSymbols

                If obj.Contains(Item.Trim().ToLower()) Then
                    total2 += 1
                End If

            Next

            If total > 0 OrElse (total > 0 AndAlso total2 > 0) Then
                Return True
            Else
                Return False
            End If

        End Function

        Public Shared Function GetQueryStrings(ByVal url As String, ByRef dictQueryStrings As Dictionary(Of String, String)) As Integer

            Dim pos As Integer = url.IndexOf("?")
            If pos > 1 Then

                Dim queryString As String = url.Substring(pos)
                queryString = queryString.Replace("?", "")
                Dim keys() As String = queryString.Split("&")
                Dim temp() As String = queryString.Split("=")

                For Each Item As String In keys

                    temp = Item.Split("=")
                    dictQueryStrings.Add(temp(0).ToLower(), temp(1))
                    temp = Nothing

                Next

            End If

            Return dictQueryStrings.Count

        End Function

        Public Shared Function GetFileNameFromURLStripOutQueryString(ByVal url As String) As String

            Return System.IO.Path.GetFileName(Regex.Replace(url, "\?.*", String.Empty))

        End Function

        Public Shared Function InsertObjectionAuditTrail(ByVal dbCNStr As String, ByVal action As String, ByVal fieldId As String, ByVal newValue As String, ByVal oldValue As String, ByVal objID As Integer, ByVal strDescription As String, ByVal strDateTime As String, ByVal user As String) As Boolean
            Dim bSuccess As Boolean = False
            Dim dbKey As String = dbCNStr
            Dim myConn As New SqlClient.SqlConnection(dbKey)

            Dim strParQuery As String = "INSERT INTO tbl_web_Audit_Trail(Action_type, ObjectionID, Record_Field_ID,  Old_Value, New_Value, Description, UpdatedDate, UpdatedBy)"
            strParQuery = strParQuery + " VALUES (@Action_type, @ObjectionID, @Record_Field_ID, @Old_Value, @New_Value, @Description, @currentDateTime, @currentUser) "
            Dim myComm As New SqlCommand(strParQuery, myConn)

            Try
                myConn.Open()
                myComm.Parameters.AddWithValue("@Action_type", action)
                myComm.Parameters.AddWithValue("@ObjectionID", objID)
                myComm.Parameters.AddWithValue("@Record_Field_ID", fieldId)
                myComm.Parameters.AddWithValue("@Old_Value", oldValue)
                myComm.Parameters.AddWithValue("@New_Value", newValue)
                myComm.Parameters.AddWithValue("@Description", strDescription)
                myComm.Parameters.AddWithValue("@currentDateTime", strDateTime)
                myComm.Parameters.AddWithValue("@currentUser", user)
                myComm.ExecuteNonQuery()
                myComm.Dispose()
                bSuccess = True
            Catch ex As Exception
                bSuccess = False
                Throw ex
            Finally
                myComm.Dispose()
                myConn.Close()
            End Try
            Return bSuccess
        End Function

        Public Shared Function GetAllowedMaxExportCount() As Integer
            Dim iMaxCount As Integer = 5000
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            Dim strSQL As String = "SELECT KeyValue From tbl_Apps_AppSettings WHERE KeyName= 'AllowedMaxExportRowCount' "
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim myReader As SqlDataReader = Nothing
            Try
                sqlCnn.Open()
                myReader = sqlCmd.ExecuteReader()
                While myReader.Read()
                    If Not IsDBNull(myReader(0)) Then
                        If Not String.IsNullOrEmpty(myReader.GetString(0)) Then
                            iMaxCount = Integer.Parse(myReader.GetString(0))
                        End If
                    End If
                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in GetAllowedMaxExportCount")
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try
            Return iMaxCount
        End Function

        'Public Shared Function GetTotalAmountRow(ByVal params As Hashtable, ByVal whereclause As String, ByVal tablename As String, ByVal colname As String) As Double
        '    Dim myConn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("dbKey"))
        '    Dim strSQL As String = ""
        '    Dim aReader As SqlDataReader = Nothing
        '    If Not String.IsNullOrEmpty(whereclause) Then
        '        strSQL = "SELECT SUM(Cast(Cast(" + colname + " as Real) as Numeric(18,2))) FROM " & tablename
        '        strSQL += " WHERE " + whereclause
        '    Else
        '        strSQL = "SELECT SUM(Cast(Cast(" + colname + " as Real) as Numeric(18,2))) FROM " & tablename
        '    End If
        '    Dim myComm As New SqlCommand(strSQL, myConn)
        '    If Not String.IsNullOrEmpty(whereclause) Then
        '        Dim param As DictionaryEntry
        '        For Each param In params
        '            myComm.Parameters.AddWithValue(param.Key, param.Value)
        '        Next
        '    End If
        '    Dim total As Double = 0.0
        '    Try
        '        myConn.Open()
        '        aReader = myComm.ExecuteReader(CommandBehavior.CloseConnection)
        '        While aReader.Read
        '            If Not IsDBNull(aReader(0)) Then
        '                total = CType(aReader(0), Double)
        '            End If
        '        End While
        '    Catch ex As Exception
        '        Throw ex
        '    Finally
        '        If Not aReader Is Nothing AndAlso Not aReader.IsClosed Then
        '            aReader.Close()
        '        End If
        '        myComm.Dispose()
        '        myConn.Close()
        '    End Try
        '    Return total
        'End Function

        Public Shared Function ConvertToDisplayTimeZoneDateTime(ByVal fromDT As DateTime) As DateTime
            If Not IsDate(fromDT) Then
                Return Nothing
            End If
            Dim diaplayTime As DateTime = Nothing
            Dim displayTimeZone As TimeZoneInfo
            Dim displayTimeString As String = Nothing
            Try
                displayTimeZone = TimeZoneInfo.FindSystemTimeZoneById(Webapps.Utils.ApplicationSettings.DisplayTimeZone)
            Catch ex As TimeZoneNotFoundException
                displayTimeZone = TimeZoneInfo.Local
            Catch ex As InvalidTimeZoneException
                displayTimeZone = TimeZoneInfo.Local
            End Try
            diaplayTime = TimeZoneInfo.ConvertTimeFromUtc(fromDT, displayTimeZone)
            Return diaplayTime
        End Function

        Public Shared Function GetCurrentDateTimeInDisplayTimeZone() As DateTime
            Return ConvertToDisplayTimeZoneDateTime(DateTime.UtcNow)
        End Function

        Public Shared Function GetCurrentDateTimeDisplayString() As String
            Dim diaplayTime As DateTime = DateTime.UtcNow
            Dim displayTimeZone As TimeZoneInfo
            Dim displayTimeString As String = Nothing
            Try
                displayTimeZone = TimeZoneInfo.FindSystemTimeZoneById(Webapps.Utils.ApplicationSettings.DisplayTimeZone)
                ' displayTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
            Catch ex As TimeZoneNotFoundException
                displayTimeZone = TimeZoneInfo.Local
            Catch ex1 As InvalidTimeZoneException
                displayTimeZone = TimeZoneInfo.Local
            Catch ex2 As Exception
                displayTimeZone = TimeZoneInfo.Local
            End Try
            diaplayTime = TimeZoneInfo.ConvertTimeFromUtc(diaplayTime, displayTimeZone)
            ' displayTimeString = diaplayTime.ToString("MMMM dd, yyyy @ hh:mm:ss")
            Return String.Format("{0} {1}", diaplayTime, IIf(displayTimeZone.IsDaylightSavingTime(diaplayTime), displayTimeZone.DaylightName, displayTimeZone.StandardName))
        End Function

        Public Shared Function GetAllowedUploadFileSize() As Long
            Dim iMaxCount As Long = 5120000  '5MB  12,582,912  12MB
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            Dim strSQL As String = "Select KeyValue  From tbl_Apps_AppSettings Where KeyName = 'AllowedUploadFileSize' "
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim myReader As SqlDataReader = Nothing
            Try
                sqlCnn.Open()
                myReader = sqlCmd.ExecuteReader()
                While myReader.Read()
                    If Not IsDBNull(myReader(0)) Then
                        If Not String.IsNullOrEmpty(myReader.GetString(0)) Then
                            iMaxCount = Long.Parse(myReader.GetString(0))
                        End If
                    End If
                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in GetAllowedUploadFileSize")
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try
            Return iMaxCount
        End Function

        Public Shared Function GetAllowedUploadFiletypes() As String
            Dim strFileTypes As String = ".doc;.docx;.xls;.ppt;.pptx;.mdb;.mdbx;.pdf;.tiff;.tif;.csv;.txt;.png"
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            Dim strSQL As String = "SELECT KeyValue  From tbl_Apps_AppSettings Where KeyName = 'AllowedUploadFileTypes' "
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim myReader As SqlDataReader = Nothing
            Try
                sqlCnn.Open()
                myReader = sqlCmd.ExecuteReader()
                While myReader.Read()
                    If Not IsDBNull(myReader(0)) Then
                        If Not String.IsNullOrEmpty(myReader.GetString(0)) Then
                            strFileTypes = myReader.GetString(0)
                        End If
                    End If
                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in GetAllowedUploadFiletypes")
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try
            Return strFileTypes
        End Function

        Public Shared Function GetUserIDByLoginID(ByVal strLogin As String) As String
            Dim strUserID As String = Nothing
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            Dim strSQL As String = "SELECT [ID] FROM tbl_usr_Details WHERE UserID =@LoginID "
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim myReader As SqlDataReader = Nothing
            Try
                sqlCnn.Open()
                sqlCmd.Parameters.AddWithValue("@LoginID", strLogin)
                myReader = sqlCmd.ExecuteReader()
                While myReader.Read()
                    If Not IsDBNull(myReader(0)) Then
                        If Not String.IsNullOrEmpty(myReader.GetInt32(0)) Then
                            strUserID = myReader.GetInt32(0).ToString()
                        End If
                    End If
                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in GetUserIDByLoginID for LoginID: " & strLogin & ". The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try
            Return strUserID
        End Function
        Public Shared Function GetContractRecordsFilterByUser(ByVal strLoginID As String) As String
            If CustomRoles.IsInRole("R_Contracts_Edit_All") OrElse CustomRoles.IsInRole("R_Contracts_View_All") Then
                Return Nothing
            End If
            Dim strRet As String = ""
            Dim strAssignedFilter As String = ""

            If CustomRoles.IsInRole("R_Contracts_Edit_Assigned") OrElse CustomRoles.IsInRole("R_Contracts_View_Assigned") Then
                strAssignedFilter = " ContractOwnerC=" & CommonUtilsv2.GetUserIDByLoginID(strLoginID)
                strRet = strAssignedFilter
            End If

            'get allowed category
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)

            Dim strSQL As String = " SELECT distinct lookupDesc AS Category FROM tbl_WEB_Lookup WHERE ( LookupType ='Role_ContractViewCategoryMapping') "
            strSQL = strSQL + " AND LookupCode  in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "
            strSQL = strSQL + " UNION "
            strSQL = strSQL + " SELECT distinct lookupDesc AS Category FROM tbl_WEB_Lookup WHERE ( LookupType ='Role_ContractEditCategoryMapping') "
            strSQL = strSQL + " AND LookupCode in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "

            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim sbRet As StringBuilder = New StringBuilder()
            Dim myReader As SqlDataReader = Nothing

            Try
                sqlCnn.Open()
                sqlCmd.Parameters.AddWithValue("@USER_ID", strLoginID)
                myReader = sqlCmd.ExecuteReader()
                Dim iCount As Integer = 0
                While myReader.Read()
                    If Not IsDBNull(myReader(0)) Then
                        If Not String.IsNullOrEmpty(myReader.GetString(0)) Then
                            If iCount = 0 Then
                                sbRet.Append(" Category IN ( ")
                                sbRet.Append(" '" & myReader.GetString(0) & "'")
                            Else
                                sbRet.Append(" ,'" & myReader.GetString(0) & "'")
                            End If

                            iCount = iCount + 1
                        End If
                    End If
                End While
                If iCount > 0 Then
                    sbRet.Append(")")
                End If
            Catch ex As Exception
                Throw New Exception("Error occurred in GetContractRecordsFilterByUser for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            If String.IsNullOrEmpty(strAssignedFilter) Then
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    ' No access to any category
                    strRet = " Category ='XXXXXXXXXX' "
                Else
                    strRet = sbRet.ToString
                End If
            Else
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    strRet = strAssignedFilter
                Else
                    strRet = "(" & sbRet.ToString & " OR " & strAssignedFilter & ")"
                End If
            End If
            Return strRet

        End Function

        Public Shared Function GetContractEditRecordsFilterByUser(ByVal strLoginID As String) As String
            If CustomRoles.IsInRole("R_Contracts_Edit_All") Then
                Return Nothing
            End If
            Dim strRet As String = ""
            Dim strAssignedFilter As String = ""

            If CustomRoles.IsInRole("R_Contracts_Edit_Assigned") Then
                strAssignedFilter = " ContractOwnerC=" & CommonUtilsv2.GetUserIDByLoginID(strLoginID)
                strRet = strAssignedFilter
            End If
            'get allowed category
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            Dim strSQL As String = " SELECT distinct lookupDesc AS Category FROM tbl_WEB_Lookup WHERE ( LookupType ='Role_ContractEditCategoryMapping') "
            strSQL = strSQL + " AND LookupCode  in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "

            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim sbRet As StringBuilder = New StringBuilder()
            Dim myReader As SqlDataReader = Nothing

            Try
                sqlCnn.Open()
                sqlCmd.Parameters.AddWithValue("@USER_ID", strLoginID)
                myReader = sqlCmd.ExecuteReader()
                Dim iCount As Integer = 0
                While myReader.Read()
                    If Not IsDBNull(myReader(0)) Then
                        If Not String.IsNullOrEmpty(myReader.GetString(0)) Then
                            If iCount = 0 Then
                                sbRet.Append(" Category IN ( ")
                                sbRet.Append(" '" & myReader.GetString(0) & "'")
                            Else
                                sbRet.Append(" ,'" & myReader.GetString(0) & "'")
                            End If

                            iCount = iCount + 1
                        End If
                    End If
                End While
                If iCount > 0 Then
                    sbRet.Append(")")
                End If
            Catch ex As Exception
                Throw New Exception("Error occurred in GetContractEditRecordsFilterByUser for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            If String.IsNullOrEmpty(strAssignedFilter) Then
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    ' No access to any category
                    strRet = " Category ='XXXXXXXXXX' "
                Else
                    strRet = sbRet.ToString
                End If
            Else
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    strRet = strAssignedFilter
                Else
                    strRet = "(" & sbRet.ToString & " OR " & strAssignedFilter & ")"
                End If
            End If
            Return strRet
        End Function

        Public Shared Function CanEditContract(ByVal strLoginID As String, ByVal iContractID As String) As Boolean
            Dim bRet As Boolean = False
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim myConn As New SqlClient.SqlConnection(dbKey)
            Dim strSQL As String = " SELECT c.ContractID FROM tbl_Web_Contracts c INNER JOIN tbl_Web_Contract_Assessment ca ON c.ContractID =ca.ContractID  WHERE c.ContractID =@ID "

            Dim strFilter As String = CommonUtilsv2.GetContractEditRecordsFilterByUser(strLoginID)
            If Not String.IsNullOrEmpty(strFilter) Then
                strSQL = strSQL & " AND (" & strFilter & ") "
            End If

            Dim myComm As New SqlCommand(strSQL, myConn)
            myComm.Parameters.AddWithValue("@ID", iContractID)
            Dim myReader As SqlDataReader = Nothing
            Try
                myConn.Open()
                myReader = myComm.ExecuteReader()
                If myReader.HasRows Then
                    bRet = True
                End If
            Catch ex As Exception
                Throw New Exception("Error occurred in CanEditContract for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myComm.Dispose()
                myConn.Close()
            End Try

            Return bRet
        End Function


        Public Shared Function GetSupplierRecordsFilterByUser(ByVal strLoginID As String) As String
            If CustomRoles.IsInRole("R_Suppliers_Edit_All") OrElse CustomRoles.IsInRole("R_Suppliers_View_All") Then
                Return Nothing
            End If
            Dim strRet As String = ""
            Dim strAssignedFilter As String = ""

            If CustomRoles.IsInRole("R_Suppliers_Edit_Assigned") OrElse CustomRoles.IsInRole("R_Suppliers_View_Assigned") Then
                strAssignedFilter = " sa.SupplierOwner=" & CommonUtilsv2.GetUserIDByLoginID(strLoginID)
                strRet = strAssignedFilter
            End If
            'get allowed category
            'Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            'Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            'Dim strSQL As String = "SELECT lookupDesc FROM tbl_WEB_Lookup WHERE LookupType ='Role_ContractEditCategoryMapping' "
            'strSQL = strSQL + " AND LookupCode not in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "
            'strSQL = strSQL + " AND LookupDesc not in (SELECT lookupDesc FROM tbl_WEB_Lookup where  LookupType ='Role_ContractViewCategoryMapping' "
            'strSQL = strSQL + " AND LookupCode in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID)) "

            'Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            'Dim myReader As SqlDataReader = Nothing
            Dim sbRet As StringBuilder = New StringBuilder()

            If String.IsNullOrEmpty(strAssignedFilter) Then
           
                strRet = ""
            Else
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    strRet = strAssignedFilter
                Else
                    strRet = "(" & sbRet.ToString & " OR " & strAssignedFilter & ")"
                End If
            End If
            Return strRet
        End Function

        Public Shared Function GetMSARecordsFilterByUser(ByVal strLoginID As String) As String
            If CustomRoles.IsInRole("R_MSAs_Edit_All") OrElse CustomRoles.IsInRole("R_MSAs_View_All") Then
                Return Nothing
            End If
            Dim strRet As String = ""
            Dim strAssignedFilter As String = ""

            If CustomRoles.IsInRole("R_MSAs_Edit_Assigned") OrElse CustomRoles.IsInRole("R_MSAs_View_Assigned") Then
                strAssignedFilter = " Reviewer=" & CommonUtilsv2.GetUserIDByLoginID(strLoginID)
                strRet = strAssignedFilter
            End If

            If String.IsNullOrEmpty(strAssignedFilter) Then
                strRet = ""
            End If
            Return strRet
        End Function

        Public Shared Function CanEditMSA(ByVal strLoginID As String, ByVal strRecordkey As String) As Boolean
            Dim bRet As Boolean = False
            If CustomRoles.IsInRole("R_MSAs_Edit_All") Then
                Return True
            End If
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            If CustomRoles.IsInRole("R_MSAs_Edit_Assigned") Then
                Dim strSQLassigned As String = "Select u.UserID from tbl_Web_MSAs ma Inner Join tbl_usr_details u On ma.Reviewer=u.ID Where MSAID=@Key "
                Dim strAnalyst As String = Nothing
                Dim params As SqlParameter() = {New SqlParameter("@Key", strRecordkey)}
                Dim myreader1 As SqlDataReader = Nothing
                Try
                    myreader1 = CommonUtilsv2.GetDataReader(dbKey, strSQLassigned, CommandType.Text, params)
                    While myreader1.Read
                        strAnalyst = myreader1(0)
                        If String.Compare(strLoginID.ToLower, strAnalyst.ToLower) = 0 Then
                            Return True
                        End If
                    End While
                Catch ex As Exception
                    Throw ex
                Finally
                    If Not myreader1 Is Nothing AndAlso Not myreader1.IsClosed Then
                        myreader1.Close()
                    End If
                End Try
            End If
            Return bRet
        End Function

        Public Shared Function GetSupplierEditRecordsFilterByUser(ByVal strLoginID As String) As String
            If CustomRoles.IsInRole("R_Suppliers_Edit_All") Then
                Return Nothing
            End If
            Dim strRet As String = ""
            Dim strAssignedFilter As String = ""

            If CustomRoles.IsInRole("R_Suppliers_Edit_Assigned") Then
                strAssignedFilter = " sa.SupplierOwner=" & CommonUtilsv2.GetUserIDByLoginID(strLoginID)
                strRet = strAssignedFilter
            End If
            'get allowed category
            'Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            'Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            'Dim strSQL As String = "SELECT lookupDesc FROM tbl_WEB_Lookup WHERE LookupType ='Role_ContractEditCategoryMapping' "
            'strSQL = strSQL + " AND LookupCode not in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "

            'Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            'Dim myReader As SqlDataReader = Nothing
            Dim sbRet As StringBuilder = New StringBuilder()
            'Try
            '    sqlCnn.Open()
            '    sqlCmd.Parameters.AddWithValue("@USER_ID", strLoginID)
            '    myReader = sqlCmd.ExecuteReader()
            '    Dim iCount As Integer = 0
            '    While myReader.Read()
            '        If Not IsDBNull(myReader(0)) Then
            '            If Not String.IsNullOrEmpty(myReader.GetString(0)) Then
            '                If iCount = 0 Then
            '                    sbRet.Append(" Category NOT IN ('" & myReader.GetString(0) & "'")
            '                    iCount = iCount + 1
            '                Else
            '                    sbRet.Append(",'" & myReader.GetString(0) & "'")
            '                End If
            '            End If
            '        End If
            '    End While
            '    If iCount > 0 Then
            '        sbRet.Append(")")
            '    End If
            'Catch ex As Exception
            '    Throw New Exception("Error occurred in GetContractEditRecordsFilterByUser for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            'Finally
            '    If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
            '        myReader.Close()
            '    End If
            '    sqlCmd.Dispose()
            '    sqlCmd = Nothing
            '    sqlCnn.Close()
            'End Try

            If String.IsNullOrEmpty(strAssignedFilter) Then
                'If String.IsNullOrEmpty(sbRet.ToString) Then
                '    ' No access to any category
                '    strRet = " Category ='XXXXXXXXXX' "
                'Else
                '    strRet = sbRet.ToString
                'End If
                strRet = ""
            Else
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    strRet = strAssignedFilter
                Else
                    strRet = "(" & sbRet.ToString & " OR " & strAssignedFilter & ")"
                End If

            End If
            Return strRet
        End Function

        Public Shared Function GetClaimRecordFilterByUser(ByVal strLoginID As String) As String
            ' assume all have view access to claims discussed on 7/12/2019
            'Now want to implement access by category
            ' Return Nothing

            If CustomRoles.IsInRole("R_Claims_Edit_All") OrElse CustomRoles.IsInRole("R_Claims_View_All") Then
                Return Nothing
            End If

            Dim strRet As String = ""
            Dim strAssignedClaimFilter As String = ""
            If CustomRoles.IsInRole("R_Claims_Edit_Assigned") OrElse CustomRoles.IsInRole("R_Claims_View_Assigned") Then
                strAssignedClaimFilter = " ReviewerID=" & CommonUtilsv2.GetUserIDByLoginID(strLoginID)
            End If
            'get allowed category
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            'Dim strSQL As String = " SELECT distinct lookupDesc FROM tbl_WEB_Lookup WHERE (LookupType ='Role_ClaimViewCategoryMapping' OR LookupType ='Role_ClaimEditCategoryMapping') "
            'strSQL = strSQL + " EXCEPT "
            'strSQL = strSQL + " SELECT distinct lookupDesc FROM tbl_WEB_Lookup WHERE (LookupType ='Role_ClaimViewCategoryMapping' OR LookupType ='Role_ClaimEditCategoryMapping') "
            'strSQL = strSQL + " AND LookupCode  in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "

            Dim strSQL As String = " SELECT distinct lookupDesc FROM tbl_WEB_Lookup WHERE (LookupType ='Role_ClaimViewCategoryMapping' OR LookupType ='Role_ClaimEditCategoryMapping') "
            strSQL = strSQL + " AND LookupCode  in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "

            Dim strMajor As String = ""
            'Dim strMinor As String = ""
            'Dim strMinorSub As String = ""
            Dim myReader As SqlDataReader = Nothing
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim sbRet As StringBuilder = New StringBuilder()
            Try
                sqlCnn.Open()
                sqlCmd.Parameters.AddWithValue("@USER_ID", strLoginID)
                myReader = sqlCmd.ExecuteReader()
                Dim iCount As Integer = 0
                Dim bFirst As Boolean = True
                While myReader.Read()
                    strMajor = ""
                    If Not IsDBNull(myReader(0)) Then
                        strMajor = myReader.GetString(0)
                        If bFirst Then
                            bFirst = False
                            sbRet.Append(" ProposedMajor IN ( ")
                            sbRet.Append(" '" & strMajor & "'")
                        Else
                            sbRet.Append(" ,'" & strMajor & "'")
                        End If

                    End If

                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in GetClaimRecordFilterByUser for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            If String.IsNullOrEmpty(strAssignedClaimFilter) Then
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    ' No access to any category
                    strRet = " ProposedMajor ='XXXXXXXXXX' "
                Else
                    strRet = sbRet.ToString + " ) "
                End If
            Else
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    strRet = strAssignedClaimFilter
                Else
                    strRet = "(" & sbRet.ToString & ") " & " OR " & strAssignedClaimFilter & ")"
                End If
            End If

            Return strRet
        End Function

        Public Shared Function GetClaimEditRecordFilterByUser_By_Category(ByVal strLoginID As String) As String
            If CustomRoles.IsInRole("R_Claims_Edit_All") Then
                Return Nothing
            End If

            Dim strRet As String = ""
            Dim strAssignedClaimFilter As String = ""
            If CustomRoles.IsInRole("R_Claims_Edit_Assigned") Then
                strAssignedClaimFilter = " ReviewerID=" & CommonUtilsv2.GetUserIDByLoginID(strLoginID)
            End If

            ''get allowed Edit category
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)

            Dim strSQL As String = " SELECT distinct lookupDesc AS Major FROM tbl_WEB_Lookup WHERE ( LookupType ='Role_ClaimEditCategoryMapping') "
            strSQL = strSQL + " AND LookupCode in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim sbRet As StringBuilder = New StringBuilder()
            Dim strCategory As String = ""
            Dim myReader As SqlDataReader = Nothing
            Try
                sqlCnn.Open()
                sqlCmd.Parameters.AddWithValue("@USER_ID", strLoginID)
                myReader = sqlCmd.ExecuteReader()
                Dim iCount As Integer = 0
                Dim bFirst As Boolean = True
                While myReader.Read()
                    If Not IsDBNull(myReader(0)) Then
                        strCategory = myReader.GetString(0)
                        If bFirst Then
                            bFirst = False
                            sbRet.Append(" ProposedMajor IN ( ")
                            sbRet.Append(" '" & strCategory & "'")
                        Else
                            sbRet.Append(" ,'" & strCategory & "'")
                        End If

                    End If

                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in CanEditClaim for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            If String.IsNullOrEmpty(strAssignedClaimFilter) Then
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    ' No access to any category
                    strRet = " ProposedMajor='XXXXXXXXXX' "
                Else
                    strRet = sbRet.ToString + " ) "
                End If
            Else
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    strRet = strAssignedClaimFilter
                Else
                    strRet = "(" & sbRet.ToString & ") " & " OR " & strAssignedClaimFilter & ")"
                End If

            End If

            Return strRet
        End Function

        Public Shared Function CanEditClaim_ByCategory(ByVal strLoginID As String, ByVal strClaimNumber As String) As Boolean
            If CustomRoles.IsInRole("R_Claims_Edit_All") Then
                Return True
            End If
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            If CustomRoles.IsInRole("R_Claims_Edit_Assigned") Then
                Dim strSQLassigned As String = "Select u.UserID from tbl_Web_Claim_Assessment ca Inner Join tbl_usr_details u On ca.Analyst=u.ID Where Claimnumber=@ClaimNumber "
                Dim strAnalyst As String = Nothing
                Dim params As SqlParameter() = {New SqlParameter("@ClaimNumber", strClaimNumber)}
                Dim myreader1 As SqlDataReader = Nothing
                Try
                    myreader1 = CommonUtilsv2.GetDataReader(dbKey, strSQLassigned, CommandType.Text, params)
                    While myreader1.Read
                        strAnalyst = myreader1(0)
                        If String.Compare(strLoginID.ToLower, strAnalyst.ToLower) = 0 Then
                            Return True
                        End If
                    End While
                Catch ex As Exception
                    Throw ex
                Finally
                    If Not myreader1 Is Nothing AndAlso Not myreader1.IsClosed Then
                        myreader1.Close()
                    End If
                End Try
            End If

            Dim strRet As String = ""
            Dim strSQLClaimNumber = "Select CA.Claimnumber from tbl_Web_Claim_Assessment CA INNER JOIN tbl_Web_FiledClaimScheduleRegister CR ON CR.ClaimNumber =CA.ClaimNumber AND CR.Claimtype='C' WHERE CA.Claimnumber=@ClaimNumber "
            'get allowed category
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            'Dim strSQL As String = " SELECT distinct lookupDesc, Comments, lookupDesc +'/'+ Comments AS Category FROM tbl_WEB_Lookup WHERE ( LookupType ='Role_ClaimEditCategoryMapping') "
            'strSQL = strSQL + " EXCEPT "
            Dim strSQL As String = " SELECT distinct lookupDesc, Comments, lookupDesc +'/'+ Comments AS Category  FROM tbl_WEB_Lookup WHERE ( LookupType ='Role_ClaimEditCategoryMapping') "
            strSQL = strSQL + " AND LookupCode  in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "
            Dim myReaderTwo As SqlDataReader = Nothing
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim sbRet As StringBuilder = New StringBuilder()
            Dim strCategory As String = ""
            Dim myReader As SqlDataReader = Nothing
            Try
                sqlCnn.Open()
                sqlCmd.Parameters.AddWithValue("@USER_ID", strLoginID)
                myReader = sqlCmd.ExecuteReader()
                Dim iCount As Integer = 0
                Dim bFirst As Boolean = True
                While myReader.Read()
                    If Not IsDBNull(myReader(2)) Then
                        strCategory = myReader.GetString(2)
                        If bFirst Then
                            bFirst = False
                            sbRet.Append(" CA.ProposedCategory IN ( ")
                            sbRet.Append(" '" & strCategory & "'")
                        Else
                            sbRet.Append(" ,'" & strCategory & "'")
                        End If

                    End If

                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in CanEditClaim for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            If String.IsNullOrEmpty(sbRet.ToString) Then
                ' No access to any category
                ' strRet = " CA.Major ='XXXXXXXXXX' "
                strRet = " CA.ProposedCategory ='XXXXXXXXXX' "
            Else
                sbRet.Append(" )")
                strSQLClaimNumber = strSQLClaimNumber + " AND " + sbRet.ToString
            End If

            sqlCmd = New SqlCommand(strSQLClaimNumber, sqlCnn)
            Dim bRet As Boolean = False
            Try
                sqlCnn.Open()

                sqlCmd.Parameters.AddWithValue("@ClaimNumber", strClaimNumber)
                myReaderTwo = sqlCmd.ExecuteReader()
                While myReaderTwo.Read()
                    bRet = True
                    Exit While
                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in CanEditClaim for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReaderTwo Is Nothing AndAlso Not myReaderTwo.IsClosed Then
                    myReaderTwo.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            Return bRet
        End Function

        'Change to use major only
        Public Shared Function GetClaimEditRecordFilterByUser(ByVal strLoginID As String) As String
            If CustomRoles.IsInRole("R_Claims_Edit_All") Then
                Return Nothing
            End If

            Dim strRet As String = ""
            Dim strAssignedClaimFilter As String = ""
            If CustomRoles.IsInRole("R_Claims_Edit_Assigned") Then
                strAssignedClaimFilter = " ReviewerID=" & CommonUtilsv2.GetUserIDByLoginID(strLoginID)
            End If

            ''get allowed Edit category
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)

            ' Dim strSQL As String = " SELECT distinct lookupDesc AS Major, Comments AS Minor, lookupDesc +'/'+ Comments AS Category  FROM tbl_WEB_Lookup WHERE ( LookupType ='Role_ClaimEditCategoryMapping') "
            Dim strSQL As String = " SELECT distinct lookupDesc AS Major  FROM tbl_WEB_Lookup WHERE ( LookupType ='Role_ClaimEditCategoryMapping') "
            strSQL = strSQL + " AND LookupCode in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim sbRet As StringBuilder = New StringBuilder()
            Dim strCategory As String = ""
            Dim myReader As SqlDataReader = Nothing
            Try
                sqlCnn.Open()
                sqlCmd.Parameters.AddWithValue("@USER_ID", strLoginID)
                myReader = sqlCmd.ExecuteReader()
                Dim iCount As Integer = 0
                Dim bFirst As Boolean = True
                While myReader.Read()
                    If Not IsDBNull(myReader(0)) Then
                        strCategory = myReader.GetString(0)
                        If bFirst Then
                            bFirst = False
                            sbRet.Append(" ProposedMajor IN ( ")
                            sbRet.Append(" '" & strCategory & "'")
                        Else
                            sbRet.Append(" ,'" & strCategory & "'")
                        End If

                    End If

                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in GetClaimEditRecordFilterByUser for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            If String.IsNullOrEmpty(strAssignedClaimFilter) Then
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    ' No access to any category
                    strRet = " ProposedMajor ='XXXXXXXXXX' "
                Else
                    strRet = sbRet.ToString + " ) "
                End If
            Else
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    strRet = strAssignedClaimFilter
                Else
                    strRet = "(" & sbRet.ToString & ") " & " OR " & strAssignedClaimFilter & ")"
                End If

            End If

            Return strRet
        End Function

        'using the assessement table 
        Public Shared Function CanEditClaim(ByVal strLoginID As String, ByVal strClaimNumber As String) As Boolean

            Dim bRet As Boolean = False
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim myConn As New SqlClient.SqlConnection(dbKey)
            Dim strSQL = " SELECT * FROM [vw_w_claimsRegister] WHERE Claimnumber=@ClaimNumber "
            Dim strFilter As String = CommonUtilsv2.GetClaimEditRecordFilterByUser(strLoginID)
            If Not String.IsNullOrEmpty(strFilter) Then
                strSQL = strSQL & " AND (" & strFilter & ") "
            End If

            Dim myComm As New SqlCommand(strSQL, myConn)
            myComm.Parameters.AddWithValue("@Claimnumber", strClaimNumber)
            Dim myReader As SqlDataReader = Nothing
            Try

                myConn.Open()
                myReader = myComm.ExecuteReader()
                If myReader.HasRows Then
                    bRet = True
                End If
            Catch ex As Exception
                Throw New Exception("Error occurred in CanEditClaim for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myComm.Dispose()
                myConn.Close()
            End Try

            Return bRet

        End Function

        Public Shared Function CanEditClaim_Using_Register_Table(ByVal strLoginID As String, ByVal strClaimNumber As String) As Boolean
            If CustomRoles.IsInRole("R_Claims_Edit_All") Then
                Return True
            End If
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            If CustomRoles.IsInRole("R_Claims_Edit_Assigned") Then
                Dim strSQLassigned As String = "Select u.UserID from tbl_Web_Claim_Assessment ca Inner Join tbl_usr_details u On ca.Analyst=u.ID Where Claimnumber=@ClaimNumber "
                Dim strAnalyst As String = Nothing
                Dim params As SqlParameter() = {New SqlParameter("@ClaimNumber", strClaimNumber)}
                Dim myreader1 As SqlDataReader = Nothing
                Try
                    myreader1 = CommonUtilsv2.GetDataReader(dbKey, strSQLassigned, CommandType.Text, params)
                    While myreader1.Read
                        strAnalyst = myreader1(0)
                        If String.Compare(strLoginID.ToLower, strAnalyst.ToLower) = 0 Then
                            Return True
                        End If
                    End While
                Catch ex As Exception
                    Throw ex
                Finally
                    If Not myreader1 Is Nothing AndAlso Not myreader1.IsClosed Then
                        myreader1.Close()
                    End If
                End Try
            End If

            Dim strRet As String = ""
            Dim strSQLClaimNumber = "Select Claimnumber from tbl_Web_FiledClaimScheduleRegister Where Claimtype='C' AND Claimnumber=@ClaimNumber "
            'get allowed category
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            Dim strSQL As String = " SELECT distinct lookupDesc, Comments FROM tbl_WEB_Lookup WHERE ( LookupType ='Role_ClaimEditCategoryMapping') "
            strSQL = strSQL + " EXCEPT "
            strSQL = strSQL + " SELECT distinct lookupDesc, Comments FROM tbl_WEB_Lookup WHERE ( LookupType ='Role_ClaimEditCategoryMapping') "
            strSQL = strSQL + " AND LookupCode  in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "
            Dim myReaderTwo As SqlDataReader = Nothing
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim sbRet As StringBuilder = New StringBuilder()
            Dim strMajor As String = ""
            Dim strMinor As String = ""
            Dim strMinorSub As String = ""
            Dim myReader As SqlDataReader = Nothing
            Try
                sqlCnn.Open()
                sqlCmd.Parameters.AddWithValue("@USER_ID", strLoginID)
                myReader = sqlCmd.ExecuteReader()
                Dim iCount As Integer = 0
                Dim bFirst As Boolean = True
                While myReader.Read()
                    strMajor = ""
                    strMinor = ""
                    strMinorSub = ""
                    If bFirst Then
                        bFirst = False
                    Else
                        sbRet.Append(" AND ")
                    End If
                    If Not IsDBNull(myReader(0)) Then
                        strMajor = myReader.GetString(0)
                    End If
                    If Not IsDBNull(myReader(1)) Then
                        strMinor = myReader.GetString(1)
                    End If
                    'If Not IsDBNull(myReader(2)) Then
                    '    strMinorSub = myReader.GetString(2)
                    'End If

                    sbRet.Append(" NOT ( (CA.Major ='" & strMajor & "')")
                    If Not String.IsNullOrEmpty(strMinor) Then
                        sbRet.Append(" AND (CA.Minor ='" & strMinor & "')")
                    End If
                    'If Not String.IsNullOrEmpty(strMinorSub) Then
                    '    sbRet.Append(" AND (isnull(SubMinor,'') ='" & strMinorSub & "')")
                    'End If
                    sbRet.Append(" )")
                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in CanEditClaim for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            If String.IsNullOrEmpty(sbRet.ToString) Then
                ' No access to any category
                strRet = " CA.Major ='XXXXXXXXXX' "
            Else
                strSQLClaimNumber = strSQLClaimNumber + " AND " + sbRet.ToString
            End If

            sqlCmd = New SqlCommand(strSQLClaimNumber, sqlCnn)
            Dim bRet As Boolean = False
            Try
                sqlCnn.Open()

                sqlCmd.Parameters.AddWithValue("@ClaimNumber", strClaimNumber)
                myReaderTwo = sqlCmd.ExecuteReader()
                While myReaderTwo.Read()
                    bRet = True
                    Exit While
                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in CanEditClaim for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReaderTwo Is Nothing AndAlso Not myReaderTwo.IsClosed Then
                    myReaderTwo.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            Return bRet
        End Function

        Public Shared Function CanEditSupplier(ByVal strLoginID As String, ByVal strRecordkey As String) As Boolean
            Dim bRet As Boolean = False
            If CustomRoles.IsInRole("R_Suppliers_Edit_All") Then
                Return True
            End If
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            If CustomRoles.IsInRole("R_Suppliers_Edit_Assigned") Then
                Dim strSQLassigned As String = "Select u.UserID from tbl_Web_Supplier_Assessment ca Inner Join tbl_usr_details u On ca.SupplierOwner=u.ID Where Match=@Key "
                Dim strAnalyst As String = Nothing
                Dim params As SqlParameter() = {New SqlParameter("@Key", strRecordkey)}
                Dim myreader1 As SqlDataReader = Nothing
                Try
                    myreader1 = CommonUtilsv2.GetDataReader(dbKey, strSQLassigned, CommandType.Text, params)
                    While myreader1.Read
                        strAnalyst = myreader1(0)
                        If String.Compare(strLoginID.ToLower, strAnalyst.ToLower) = 0 Then
                            Return True
                        End If
                    End While
                Catch ex As Exception
                    Throw ex
                Finally
                    If Not myreader1 Is Nothing AndAlso Not myreader1.IsClosed Then
                        myreader1.Close()
                    End If
                End Try
            End If
            Return bRet
        End Function


        Public Shared Function GetContract2RecordsFilterByUser(ByVal strLoginID As String) As String
            If CustomRoles.IsInRole("R_Contracts_Edit_All_2") OrElse CustomRoles.IsInRole("R_Contracts_2_View_All") Then
                Return Nothing
            End If
            Dim strRet As String = ""
            Dim strAssignedFilter As String = ""

            If CustomRoles.IsInRole("R_Contracts_Edit_Assigned_2") OrElse CustomRoles.IsInRole("R_Contracts_2_View_Assigned") Then
                strAssignedFilter = " ca.ContractOwnerC=" & CommonUtilsv2.GetUserIDByLoginID(strLoginID)
                strRet = strAssignedFilter
            End If
            'get allowed category
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            Dim strSQL As String = "SELECT lookupDesc FROM tbl_WEB_Lookup WHERE LookupType ='Role_ContractEditCategoryMapping' "
            strSQL = strSQL + " AND LookupCode not in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "
            strSQL = strSQL + " AND LookupDesc not in (SELECT lookupDesc FROM tbl_WEB_Lookup where  LookupType ='Role_ContractViewCategoryMapping' "
            strSQL = strSQL + " AND LookupCode in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID)) "

            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim myReader As SqlDataReader = Nothing
            Dim sbRet As StringBuilder = New StringBuilder()
            Try
                sqlCnn.Open()
                sqlCmd.Parameters.AddWithValue("@USER_ID", strLoginID)
                myReader = sqlCmd.ExecuteReader()
                Dim iCount As Integer = 0
                While myReader.Read()
                    If Not IsDBNull(myReader(0)) Then
                        If Not String.IsNullOrEmpty(myReader.GetString(0)) Then
                            If iCount = 0 Then
                                sbRet.Append(" Category NOT IN ('" & myReader.GetString(0) & "'")
                                iCount = iCount + 1
                            Else
                                sbRet.Append(",'" & myReader.GetString(0) & "'")
                            End If
                        End If
                    End If
                End While
                If iCount > 0 Then
                    sbRet.Append(")")
                End If
            Catch ex As Exception
                Throw New Exception("Error occurred in GetContractRecordsFilterByUser for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            If String.IsNullOrEmpty(strAssignedFilter) Then
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    ' No access to any category
                    strRet = " Category ='XXXXXXXXXX' "
                Else
                    strRet = sbRet.ToString
                End If
            Else
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    strRet = strAssignedFilter
                Else
                    strRet = "(" & sbRet.ToString & " OR " & strAssignedFilter & ")"
                End If
            End If
            Return strRet


        End Function

        Public Shared Function GetContract2EditRecordsFilterByUser(ByVal strLoginID As String) As String
            If CustomRoles.IsInRole("R_Contracts_Edit_All_2") Then
                Return Nothing
            End If
            Dim strRet As String = ""
            Dim strAssignedFilter As String = ""

            If CustomRoles.IsInRole("R_Contracts_Edit_Assigned_2") Then
                strAssignedFilter = " ca.ContractOwnerC=" & CommonUtilsv2.GetUserIDByLoginID(strLoginID)
                strRet = strAssignedFilter
            End If
            'get allowed category
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            Dim strSQL As String = "SELECT lookupDesc FROM tbl_WEB_Lookup WHERE LookupType ='Role_ContractEditCategoryMapping' "
            strSQL = strSQL + " AND LookupCode not in (SELECT Role_ID from tbl_ROLES_UserRoles Where USER_ID=@USER_ID) "

            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim myReader As SqlDataReader = Nothing
            Dim sbRet As StringBuilder = New StringBuilder()
            Try
                sqlCnn.Open()
                sqlCmd.Parameters.AddWithValue("@USER_ID", strLoginID)
                myReader = sqlCmd.ExecuteReader()
                Dim iCount As Integer = 0
                While myReader.Read()
                    If Not IsDBNull(myReader(0)) Then
                        If Not String.IsNullOrEmpty(myReader.GetString(0)) Then
                            If iCount = 0 Then
                                sbRet.Append(" Category NOT IN ('" & myReader.GetString(0) & "'")
                                iCount = iCount + 1
                            Else
                                sbRet.Append(",'" & myReader.GetString(0) & "'")
                            End If
                        End If
                    End If
                End While
                If iCount > 0 Then
                    sbRet.Append(")")
                End If
            Catch ex As Exception
                Throw New Exception("Error occurred in GetContractEditRecordsFilterByUser for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            If String.IsNullOrEmpty(strAssignedFilter) Then
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    ' No access to any category
                    strRet = " Category ='XXXXXXXXXX' "
                Else
                    strRet = sbRet.ToString
                End If
            Else
                If String.IsNullOrEmpty(sbRet.ToString) Then
                    strRet = strAssignedFilter
                Else
                    strRet = "(" & sbRet.ToString & " OR " & strAssignedFilter & ")"
                End If
            End If
            Return strRet
        End Function

        Public Shared Function GetHostIPAddress() As String
            GetHostIPAddress = String.Empty
            Dim strHostName As String = System.Net.Dns.GetHostName()
            Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)

            For Each ipheal As System.Net.IPAddress In iphe.AddressList
                If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                    GetHostIPAddress = ipheal.ToString()
                End If
            Next
            Return GetHostIPAddress
        End Function


        Public Shared Function GetHostSiteID() As String
            GetHostSiteID = String.Empty
            Dim strHostIP As String = CommonUtilsv2.GetHostIPAddress()
            If Not String.IsNullOrEmpty(strHostIP) Then
                Dim iLastIndex As Integer = strHostIP.LastIndexOf(".")
                If iLastIndex > 0 Then
                    GetHostSiteID = strHostIP.Substring(iLastIndex + 1, strHostIP.Length - iLastIndex - 1)
                End If
            End If
            Return GetHostSiteID
        End Function

        Public Shared Function GetEnvironmentAndHost() As String
            Dim strHostIP As String = GetHostSiteID()
            Dim strRet As String = ""
            If Not String.IsNullOrEmpty(Webapps.Utils.ApplicationSettings.Environment) Then
                strRet += Webapps.Utils.ApplicationSettings.Environment
            End If
            If Not String.IsNullOrEmpty(strHostIP) Then
                strRet += "(" & strHostIP & ")"
            End If
            Return strRet
        End Function



        Public Shared Function GetUserEmailByLoginID(ByVal strLogin As String) As String
            Dim dbKey1 As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim strSQL As String = "SELECT email FROM tbl_usr_Details WHERE UserID =@LoginID "
            Dim myReader As SqlDataReader = Nothing
            Dim strUserEmails As String = ""
            Dim params As SqlParameter() = { _
                New SqlParameter("@LoginID", strLogin) _
            }

            Try
                myReader = CommonUtilsv2.GetDataReader(dbKey1, strSQL, CommandType.Text, params)
                While myReader.Read
                    strUserEmails += myReader(0)
                    strUserEmails += ";"
                End While
                strUserEmails = strUserEmails.Remove(strUserEmails.Length - 1)
            Catch ex As Exception
                ' Throw ex
                strUserEmails = ""
                Throw New Exception("Error occurred in GetUserEmailByLoginID().  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
            End Try
            Return strUserEmails
        End Function

        Public Shared Function GetUserEmailByUserID(ByVal anID As Integer) As String
            Dim dbKey1 As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim strSQL As String = "SELECT email FROM tbl_usr_Details WHERE ID =@ID "
            Dim myReader As SqlDataReader = Nothing
            Dim strUserEmails As String = ""
            Dim params As SqlParameter() = {
                New SqlParameter("@ID", anID)
            }

            Try
                myReader = CommonUtilsv2.GetDataReader(dbKey1, strSQL, CommandType.Text, params)
                While myReader.Read
                    strUserEmails += myReader(0)
                    strUserEmails += ";"
                End While
                strUserEmails = strUserEmails.Remove(strUserEmails.Length - 1)
            Catch ex As Exception
                ' Throw ex
                strUserEmails = ""
                Throw New Exception("Error occurred in GetUserEmailByUserD().  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
            End Try
            Return strUserEmails
        End Function

        Public Shared Function GetFullChangeLog(ByVal strFieldName As String, ByVal strOldValue As String, ByVal strNewValue As String, ByRef bDataChanged As Boolean, Optional ByVal stFieldDataType As String = "String") As String
            Dim strRet As String = String.Empty
            If String.IsNullOrEmpty(strFieldName) Then
                Return String.Empty
            End If
            If (String.IsNullOrEmpty(strOldValue) AndAlso String.IsNullOrEmpty(strNewValue)) Then
                Return String.Empty
            End If

            If String.IsNullOrEmpty(strOldValue) Then
                bDataChanged = True
                strOldValue = "Not Set"
            ElseIf String.IsNullOrEmpty(strNewValue) Then
                bDataChanged = True
                strNewValue = "Not Set"
            End If


            Select Case stFieldDataType
                Case "String"
                    If String.Compare(strOldValue, strNewValue, True) = 0 Then
                        strRet = "<tr><td align='left' style='color:#FFFFFF;'>" + strFieldName + "</td><td align='left' style='color:#FFFFFF;'>" + "Data not changed.  The value is: " + strNewValue.ToString + "</td></tr>"
                    Else
                        bDataChanged = True
                        strRet = "<tr><td align='left' style='color:#FFFFFF;'>" + strFieldName + "</td><td align='left' style='color:#FFFFFF;'>" + "Changed from " + strOldValue.ToString + " to " + strNewValue.ToString + "</td></tr>"
                    End If
                Case "Numeric"
                    If IsNumeric(strOldValue) Then
                        If IsNumeric(strNewValue) Then
                            If Decimal.Parse(strOldValue) = Decimal.Parse(strNewValue) Then
                            Else
                                bDataChanged = True
                                strRet = "<tr><td align='left' style='color:#FFFFFF;'>" + strFieldName + "</td><td align='left' style='color:#FFFFFF;'>" + "Changed from " + strOldValue.ToString + " to " + strNewValue.ToString + "</td></tr>"
                            End If
                        Else ' strNewValue not numeric (should be "Not Set")
                            bDataChanged = True
                            strRet = "<tr><td align='left' style='color:#FFFFFF;'>" + strFieldName + "</td><td align='left' style='color:#FFFFFF;'>" + "Changed from " + strOldValue.ToString + " to " + strNewValue.ToString + "</td></tr>"
                        End If
                    Else ' strOldValue not numeric (should be "Not Set")
                        If IsNumeric(strNewValue) Then
                            bDataChanged = True
                            strRet = "<tr><td align='left' style='color:#FFFFFF;'>" + strFieldName + "</td><td align='left' style='color:#FFFFFF;'>" + "Changed from " + strOldValue.ToString + " to " + strNewValue.ToString + "</td></tr>"
                        End If
                    End If
            End Select
            Return strRet
        End Function

        Public Shared Function CreateAuditLog(ByVal strIDValue As String, ByVal strLoginID As String) As Boolean
            Dim bRet As Boolean = False
            If strIDValue < 1 Then
                'no source record
                Return False
            End If
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
            Dim strNow As String = DateTime.Now.ToString

            'Dim strSQL As String = "INSERT INTO  tbl_Web_DocumentAttributeslog(ID,DCN,QueueID,Category,DocType,Status,Priority,Description,Comments,CommentsExport,DisplaySequence,Index1,Index2,Index3,Index4,Index5,Index6,Index7,Index8,Index9,Index10,"
            'strSQL = strSQL & "QueueStartDate,UploadedDate,UploadedBy,UpdatedDate,UpdatedBy,QueueAge,DocAge,Owner,DeleteNotes,RecordStatus,ClientName,LogCreatedDate,LogCreatedBy) "
            'strSQL = strSQL & "  Select ID,DCN,QueueID,Category,DocType,Status,Priority,Description,Comments,CommentsExport,DisplaySequence,Index1,Index2,Index3,Index4,Index5,Index6,Index7,Index8,Index9,Index10,"
            'strSQL = strSQL & "QueueStartDate,UploadedDate,UploadedBy,UpdatedDate,UpdatedBy,QueueAge,DocAge,Owner,DeleteNotes,RecordStatus,ClientName,@strNow,@strLoginID FROM tbl_Web_DocumentAttributes WHERE DCN=@DCN;"
            Dim strSQL As String = " INSERT INTO  tbl_Web_DocumentAttributeslog select *,@strNow,@strLoginID from tbl_Web_DocumentAttributes where DCN=@DCN; "

            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            sqlCmd.Parameters.AddWithValue("@DCN", strIDValue)
            sqlCmd.Parameters.AddWithValue("@strNow", strNow)
            sqlCmd.Parameters.AddWithValue("@strLoginID", strLoginID)

            Try
                sqlCnn.Open()
                sqlCmd.ExecuteNonQuery()
                bRet = True
            Catch ex As Exception
                Throw New Exception("Error occurred in CreateAuditLog() for Source table: tbl_Web_DocumentAttributes AND record ID : " & strIDValue & ".  The error detail is: " & ex.Message)
                bRet = False
            Finally
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try
            Return bRet
        End Function

        'Public Shared Function CreateAuditLog(ByVal strSourceTableName As String, ByVal strIDColumnName As String, ByVal strIDValue As String, ByVal strLogTableName As String, ByVal strLoginID As String) As Boolean
        '    Dim bRet As Boolean = False
        '    If strIDValue < 1 Then
        '        'no source record
        '        Return False
        '    End If
        '    Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        '    Dim sqlCnn As New SqlClient.SqlConnection(dbKey)
        '    Dim strNow As String = DateTime.Now.ToString
        '    Dim strSQL As String = " INSERT INTO " & strLogTableName & " SELECT *, '" & strNow & "', '" & strLoginID & "' FROM " & strSourceTableName & " WHERE " & strIDColumnName & "=" & strIDValue

        '    Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
        '    Try
        '        sqlCnn.Open()
        '        sqlCmd.ExecuteNonQuery()
        '        bRet = True
        '    Catch ex As Exception
        '        Throw New Exception("Error occurred in CreateAuditLog() for Source table: " & strSourceTableName & " AND record ID : " & strIDValue & ".  The error detail is: " & ex.Message)
        '        bRet = False
        '    Finally
        '        sqlCmd.Dispose()
        '        sqlCmd = Nothing
        '        sqlCnn.Close()
        '    End Try
        '    Return bRet
        'End Function

        Public Shared Function CreateDCNLog(ByVal DCN As Integer, ByVal LogType As Integer, ByVal strLogComment As String, ByVal strUserID As String, Optional ByVal StrUserIP As String = "") As Boolean
            Dim bRet As Boolean = False
            If DCN < 1 Then
                Return False
            End If
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)

            Dim strSQL As String = "prc_CreateDCNLog"

            Dim params As SqlParameter() = {
              New SqlParameter("@DCN", DCN),
              New SqlParameter("@LogType", LogType),
              New SqlParameter("@Comments", strLogComment),
              New SqlParameter("@UserID", strUserID),
              New SqlParameter("@UserIP", StrUserIP)
              }

            Try
                RunNonQuery(dbKey, strSQL, CommandType.StoredProcedure, params)
                bRet = True
            Catch ex As Exception
                Throw New Exception("Error occurred in CreateDCNLog() for DCN : " & DCN.ToString & ".  The error detail is: " & ex.Message)
                bRet = False
            Finally

            End Try
            Return bRet
        End Function


        Public Shared Function GetErrorDetails(ByRef exp As Exception) As String
            Dim sRet As String = String.Empty
            If exp Is Nothing Then exp = New Exception("Exception not available")
            Dim sTrace As StackTrace = New StackTrace(exp, True)
            Dim aFrame As StackFrame = sTrace.GetFrame(0)
            Dim SourceFile As String = aFrame.GetFileName()
            Dim lineNo As Int32 = aFrame.GetFileLineNumber()
            Dim CallingMethod As String = String.Empty
            If Not sTrace.GetFrame(1) Is Nothing Then
                CallingMethod = sTrace.GetFrame(1).GetMethod().Name
            End If
            If String.IsNullOrEmpty(SourceFile) Then SourceFile = String.Empty
            If String.IsNullOrEmpty(CallingMethod) Then CallingMethod = String.Empty
            If lineNo = Nothing Then lineNo = 0

            sRet = " Calling Method: " & CallingMethod & vbTab & "Source File: " & SourceFile & vbTab & " Line: " & lineNo.ToString & vbTab & " Error Detail: " & exp.Message
            Return sRet
        End Function

        Public Shared Function GetErrorDetails(ByVal strErrLoc As String, ByRef exp As Exception) As String
            Dim sRet As String = String.Empty
            If String.IsNullOrEmpty(strErrLoc) Then strErrLoc = String.Empty
            If exp Is Nothing Then exp = New Exception("Exception not available")
            sRet = "Error location: " & strErrLoc & vbCrLf & " Error Detail: " & exp.Message & vbCrLf & "StackTrace: " & exp.StackTrace
            Return sRet
        End Function

        Public Shared Function GetCurrentUserEamil(ByVal UserId As String) As String
            Dim myComm As New SqlCommand()
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim strSQL As String = "Select Email From tbl_usr_Details Where UserID=@UserID"
            Dim myReader As SqlDataReader = Nothing
            Dim params As SqlParameter() = {
            New SqlParameter("@UserID", UserId)
            }
            Dim StrEamil As String = ""
            Try
                myReader = CommonUtilsv2.GetDataReader(dbKey, strSQL, CommandType.Text, params)
                myReader.Read()
                StrEamil = myReader(0)
            Catch ex As Exception
                StrEamil = ""
                Throw New Exception("Error occurred in GetCurrentUserEamil().  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing Then
                    myReader.Close()
                End If
            End Try
            Return StrEamil
        End Function

        Public Shared Function GetDCNRecordFilterByUser(ByVal strUserID As String) As String

            If CustomRoles.IsInRole("R_WG_Edit_All") OrElse CustomRoles.IsInRole("R_WG_View_All") Then
                Return Nothing
            End If

            Dim strRet As String = ""

            'get allowed WGs
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim sqlCnn As New SqlClient.SqlConnection(dbKey)

            Dim strSQL As String = " SELECT distinct WGID FROM v_w_WGUserList WHERE User_ID =@User_ID "
        
            Dim strWGID As String = ""

            Dim myReader As SqlDataReader = Nothing
            Dim sqlCmd As New SqlCommand(strSQL, sqlCnn)
            Dim sbRet As StringBuilder = New StringBuilder()
            Try
                sqlCnn.Open()
                sqlCmd.Parameters.AddWithValue("@User_ID", strUserID)
                myReader = sqlCmd.ExecuteReader()
                Dim iCount As Integer = 0
                Dim bFirst As Boolean = True
                While myReader.Read()
                    strWGID = ""
                    If Not IsDBNull(myReader(0)) Then
                        strWGID = myReader.GetInt32(0).ToString()
                        If bFirst Then
                            bFirst = False
                            sbRet.Append(" GroupID IN ( ")
                            sbRet.Append(strWGID)
                        Else
                            sbRet.Append(" , " & strWGID)
                        End If

                    End If

                End While
            Catch ex As Exception
                Throw New Exception("Error occurred in GetDCNRecordFilterByUser for UserID: " & strUserID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                sqlCmd.Dispose()
                sqlCmd = Nothing
                sqlCnn.Close()
            End Try

            If String.IsNullOrEmpty(sbRet.ToString) Then
                strRet = " GroupID IN ( -9999 ) "
            Else
                strRet = sbRet.ToString & ") "
            End If

            Return strRet
        End Function

        Public Shared Function HasAccesstoDCN(ByVal strLoginID As String, ByVal iDCN As String) As Boolean
            Dim bRet As Boolean = False
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
            Dim myConn As New SqlClient.SqlConnection(dbKey)
            Dim strSQL As String = " SELECT * FROM v_w_QueueRegister  WHERE DCN =@DCN "

            Dim strFilter As String = CommonUtilsv2.GetDCNRecordFilterByUser(strLoginID)
            If Not String.IsNullOrEmpty(strFilter) Then
                strSQL = strSQL & " AND (" & strFilter & ") "
            End If

            Dim myComm As New SqlCommand(strSQL, myConn)
            myComm.Parameters.AddWithValue("@DCN", iDCN)
            Dim myReader As SqlDataReader = Nothing
            Try
                myConn.Open()
                myReader = myComm.ExecuteReader()
                If myReader.HasRows Then
                    bRet = True
                End If
            Catch ex As Exception
                Throw New Exception("Error occurred in HasAccesstoDCN for LoginID: " & strLoginID & ".  The error detail is: " & ex.Message)
            Finally
                If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myComm.Dispose()
                myConn.Close()
            End Try

            Return bRet
        End Function

    End Class

End Namespace