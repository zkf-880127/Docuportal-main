Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports Webapps.Utils

Public Class CustomRoles

    Public Shared Function IsInRole(ByRef ds As DataSet, ByVal roleid As Integer) As Boolean

        Dim expression As String = String.Format("Roles_ID = {0}", roleid)

        Dim foundRows() As DataRow = ds.Tables(2).Select(expression)

        If foundRows.Length > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Shared Function IsInRole(ByRef ds As DataSet, ByRef roleList As List(Of Integer)) As Boolean

        For Each Row As DataRow In ds.Tables(2).Rows

            If roleList.Contains(Row("Roles_ID")) Then

                Return True

            End If
        Next
        Return False

    End Function

    Public Shared Function RolesForPageLoad() As Boolean

        ' TODO : CREATE A METHOD FOR THE CASTING
        Dim ds As DataSet = Nothing 'DirectCast(HttpContext.Current.Session("Roles"), DataSet)

        If ds Is Nothing Then
            GetData(ds)
        End If

        Dim count As Integer = 0

        If String.Compare(CommonUtilsv2.GetFileNameFromURLStripOutQueryString(System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString)), "Page.aspx", True) = 0 Then

            If Not ds Is Nothing Then

                If HttpContext.Current.Request.QueryString.AllKeys.Length > 0 Then

                    ' CHECK TO SEE IF PAGE IS WITHIN ROLES
                    Dim databasePage As String = String.Empty

                    For Each Row As DataRow In ds.Tables(0).Rows

                        If String.Compare(System.IO.Path.GetFileName(Row("pagename")), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), True) = 0 Then
                            count += 1
                            databasePage = Row("pagename")
                        End If

                    Next

                    ' IF PAGE EXISTS IN ROLES CHECK TO SEE IF QUERY STRING KEYS ARE VALID IS WITHIN ROLES
                    If count > 0 Then

                        Dim querystringkeycount As Integer = 0
                        Dim querystringcount As Integer = 0

                        If HttpContext.Current.Request.QueryString.AllKeys.Length > 0 Then

                            Dim dictQueryStringsFromDatabase As New System.Collections.Generic.Dictionary(Of String, String)

                            querystringcount = CommonUtilsv2.GetQueryStrings(databasePage, dictQueryStringsFromDatabase)

                            For Each Key As String In HttpContext.Current.Request.QueryString.AllKeys

                                'System.Diagnostics.Debug.WriteLine(String.Format("QueryStringKey = '{0}' and QueryStringValue = '{1}' and FileName = '{2}' ", Key, Request.QueryString(Key), System.IO.Path.GetFileName(Page.AppRelativeVirtualPath)))

                                If dictQueryStringsFromDatabase.ContainsKey(Key.ToLower()) Then
                                    querystringkeycount += 1
                                Else
                                    Exit For
                                End If

                            Next

                        End If

                        ' IF QUERY STRINGS KEY(S) ARE NOT THE SAME KEYS AND NUMBER OF KEY REMOVE ACCESS TO PAGE
                        If querystringcount = querystringkeycount Then
                            count = querystringcount
                        Else
                            count = 0
                        End If

                    End If

                Else

                    Dim expression As String = String.Format(" pagename = '{0}' ", System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString)) ' Page.AppRelativeVirtualPath
                    Dim foundRows() As DataRow
                    foundRows = ds.Tables(0).Select(expression)
                    count = foundRows.Length

                End If

            End If

        Else

            If Not ds Is Nothing Then

                If HttpContext.Current.Request.QueryString.AllKeys.Length > 0 Then

                    ' CHECK TO SEE IF PAGE IS WITHIN ROLES
                    Dim currentPage As String = CommonUtilsv2.GetFileNameFromURLStripOutQueryString(System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString))
                    Dim databasePage As String = String.Empty

                    For Each Row As DataRow In ds.Tables(0).Rows

                        If String.Compare(CommonUtilsv2.GetFileNameFromURLStripOutQueryString(System.IO.Path.GetFileName(Row("pagename"))), currentPage, True) = 0 Then
                            count += 1
                            databasePage = Row("pagename")
                        End If

                    Next

                    ' IF PAGE EXISTS IN ROLES CHECK TO SEE IF QUERY STRING KEYS ARE VALID IS WITHIN ROLES
                    If count > 0 Then

                        Dim querystringkeycount As Integer = 0
                        Dim querystringcount As Integer = 0

                        If HttpContext.Current.Request.QueryString.AllKeys.Length > 0 Then

                            Dim dictQueryStringsFromDatabase As New System.Collections.Generic.Dictionary(Of String, String)

                            querystringcount = CommonUtilsv2.GetQueryStrings(databasePage, dictQueryStringsFromDatabase)

                            For Each Key As String In HttpContext.Current.Request.QueryString.AllKeys

                                'System.Diagnostics.Debug.WriteLine(String.Format("QueryStringKey = '{0}' and QueryStringValue = '{1}' and FileName = '{2}' ", Key, Request.QueryString(Key), System.IO.Path.GetFileName(Page.AppRelativeVirtualPath)))

                                If dictQueryStringsFromDatabase.ContainsKey(Key.ToLower()) Then
                                    querystringkeycount += 1
                                Else
                                    Exit For
                                End If

                            Next

                        End If

                        ' IF QUERY STRINGS KEY(S) ARE NOT THE SAME KEYS AND NUMBER OF KEY REMOVE ACCESS TO PAGE
                        If querystringcount = querystringkeycount Then
                            count = querystringcount
                        Else
                            count = 0
                        End If

                    End If

                Else

                    Dim expression As String = String.Format(" pagename = '{0}' ", System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString)) ' Page.AppRelativeVirtualPath
                    Dim foundRows() As DataRow
                    foundRows = ds.Tables(0).Select(expression)
                    count = foundRows.Length

                End If

            End If

        End If

        If count > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Shared Function HasAccess(ByVal name As String) As Boolean

        Dim ds As System.Data.DataSet = DirectCast(HttpContext.Current.Session("Roles"), System.Data.DataSet)
        If ds Is Nothing Then
            GetData(ds)
        End If
        Dim count As Integer = 0

        If Not ds Is Nothing Then

            'Dim expression As String = String.Format(" Name = '{0}' ", name)
            Dim expression As String = String.Format(" misc_id = '{0}' ", name)
            Dim foundRows() As System.Data.DataRow
            foundRows = ds.Tables(1).Select(expression)
            count += foundRows.Length

        End If

        If count > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Shared Sub TransferIfNotInRole(Optional ByVal isModalOrIframe As Boolean = False)
        Dim strPage As String = CommonUtilsv2.GetFileNameFromURLStripOutQueryString(System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString))

        If String.IsNullOrEmpty(strPage) Then
            HttpContext.Current.Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
            Exit Sub
        End If

        If isModalOrIframe Then

            ' DO THE FOLLOWING IF IS MODAL OR IFRAME
            If String.Compare(CommonUtilsv2.GetFileNameFromURLStripOutQueryString(System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString)), "Page.aspx", True) = 0 Then

                ' IF PAGE.ASPX:
                If String.IsNullOrEmpty(HttpContext.Current.Session("User")) Then
                    HttpContext.Current.Response.Redirect(System.Configuration.ConfigurationManager.AppSettings("Login"), False)
                    ' Do Nothing
                Else
                    If String.Compare(System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Webapps.Utils.ApplicationSettings.Homepage, True) = 0 Then
                        ' Do nothing
                    Else
                        HttpContext.Current.Response.Write("<script>window.open('" & Webapps.Utils.ApplicationSettings.Homepage & "','_parent');<" & "/script>")
                    End If

                End If

            Else

                ' IF ANOTHER PAGE THAT IS NOT PAGE.ASPX:
                If String.IsNullOrEmpty(HttpContext.Current.Session("User")) Then
                    HttpContext.Current.Session("RedirectURLAfterLogin") = System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString)
                    HttpContext.Current.Response.Write("<script>window.open('" & System.Configuration.ConfigurationManager.AppSettings("Login") & "','_parent');<" & "/script>")
                Else
                    HttpContext.Current.Session("RedirectURLAfterLogin") = Nothing
                    HttpContext.Current.Response.Write("<script>window.open('" & Webapps.Utils.ApplicationSettings.Homepage & "','_parent');<" & "/script>")
                End If

            End If

        Else
            ' DO THE FOLLOWING IF NOT MODAL OR IFRAME
            If String.Compare(CommonUtilsv2.GetFileNameFromURLStripOutQueryString(System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString)), "Page.aspx", True) = 0 Then

                ' IF PAGE.ASPX:
                If String.IsNullOrEmpty(HttpContext.Current.Session("User")) Then
                    HttpContext.Current.Response.Redirect(System.Configuration.ConfigurationManager.AppSettings("Login"), False)
                    ' Do Nothing
                Else
                    If String.Compare(System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Webapps.Utils.ApplicationSettings.Homepage, True) = 0 Then
                        ' Do nothing
                    Else
                        HttpContext.Current.Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
                    End If

                End If

            Else

                ' IF ANOTHER PAGE THAT IS NOT PAGE.ASPX:
                If String.IsNullOrEmpty(HttpContext.Current.Session("User")) Then
                    HttpContext.Current.Session("RedirectURLAfterLogin") = System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString)
                    HttpContext.Current.Response.Write("<script>window.open('" & System.Configuration.ConfigurationManager.AppSettings("Login") & "','_parent');<" & "/script>")
                Else
                    HttpContext.Current.Session("RedirectURLAfterLogin") = Nothing
                    HttpContext.Current.Response.Write("<script>window.open('" & Webapps.Utils.ApplicationSettings.Homepage & "','_parent');<" & "/script>")
                End If

            End If
        End If

    End Sub

    Public Shared Sub GetData(ByRef ds As System.Data.DataSet)

        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")

        Dim sp As String = "prc_GetRolesInformation"
        Dim params As System.Data.SqlClient.SqlParameter() = {
            New System.Data.SqlClient.SqlParameter("@UserName", HttpContext.Current.Session("User")),
            New System.Data.SqlClient.SqlParameter("@LanguageTag", GetLanguageSubTag())
            }

        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            ds = Webapps.Utils.CommonUtilsv2.GetDataSet(dbKey, sp, System.Data.CommandType.StoredProcedure, params)

        Catch ex As Exception

            Throw ex

        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try

    End Sub

    Public Shared Function GetLanguageSubTag() As String

        Dim mainLanguage As String = ""

        Dim request As HttpRequest = HttpContext.Current.Request

        If Not request.UserLanguages Is Nothing Then
            mainLanguage = request.UserLanguages(0) ' Full language 
        End If

        'Dim mainLanguagePrefix As String = request.UserLanguages(0).Substring(0, 2) ' Sub Tag

        Return mainLanguage

    End Function

    Public Shared Function HasAccesstoPage(ByVal strPageName As String) As Boolean

        Dim ds As DataSet = DirectCast(HttpContext.Current.Session("Roles"), DataSet)

        If ds Is Nothing Then
            GetData(ds)
        End If
        Dim databasePage As String = ""
        Dim count As Integer = 0

        If Not ds Is Nothing Then
            ' TO DO:  Need to handle page with querrystring
            'If HttpContext.Current.Request.QueryString.AllKeys.Length > 0 Then

            'Dim expression As String = String.Format(" pagename = '{0}' ", strPageName) ' Page.AppRelativeVirtualPath
            'Dim foundRows() As DataRow
            'foundRows = ds.Tables(0).Select(expression)
            'count = foundRows.Length

            For Each Row As DataRow In ds.Tables(0).Rows

                If String.Compare(CommonUtilsv2.GetFileNameFromURLStripOutQueryString(System.IO.Path.GetFileName(Row("pagename"))), strPageName, True) = 0 Then
                    count += 1
                    databasePage = Row("pagename")
                End If

            Next

        End If

        If count > 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Shared Function HasAccessToMainID(ByVal MainID As String) As Boolean

        Dim ds As System.Data.DataSet = DirectCast(HttpContext.Current.Session("Roles"), System.Data.DataSet)
        Dim count As Integer = 0

        If Not ds Is Nothing Then

            Dim expression As String = String.Format(" main_id = '{0}' ", MainID)
            Dim foundRows() As System.Data.DataRow
            foundRows = ds.Tables(0).Select(expression)
            count += foundRows.Length

        End If

        If count > 0 Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Shared Function IsInRole(ByVal roleID As String) As Boolean

        Dim ds As System.Data.DataSet = DirectCast(HttpContext.Current.Session("Roles"), System.Data.DataSet)
        Dim count As Integer = 0

        If ds Is Nothing Then
            GetData(ds)
        End If
        If Not ds Is Nothing Then

            Dim expression As String = String.Format(" role_id = '{0}' ", roleID)
            Dim foundRows() As System.Data.DataRow
            foundRows = ds.Tables(2).Select(expression)
            count += foundRows.Length

        End If

        If count > 0 Then
            Return True
        Else
            Return False
        End If

    End Function
End Class
