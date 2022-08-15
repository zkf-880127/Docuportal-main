Imports System.Collections.Specialized
Imports System.Web
Imports System.Data.SqlClient
Imports Webapps.Utils

Namespace Configuration

    Public Class ExtendedSiteMapProvider
        Inherits XmlSiteMapProvider

        Public Overrides Sub Initialize(ByVal name As String, ByVal attributes As NameValueCollection)
            MyBase.Initialize(name, attributes)
            Dim resolveHandler As New SiteMapResolveEventHandler(AddressOf SmartSiteMapProvider_SiteMapResolve)
            AddHandler Me.SiteMapResolve, resolveHandler
        End Sub

        Function SmartSiteMapProvider_SiteMapResolve(ByVal sender As Object, ByVal e As SiteMapResolveEventArgs) As SiteMapNode
            If (SiteMap.CurrentNode Is Nothing) Then Return Nothing

            Dim this As New XmlSiteMapProvider
            Dim temp As SiteMapNode
            temp = SiteMap.CurrentNode.Clone(True)
            Dim u As Uri = New Uri(e.Context.Request.Url.ToString())
            Dim tempNode As SiteMapNode = temp

            While Not tempNode Is Nothing
                Dim qs As String = GetQueryString(tempNode, e.Context)
                If Not qs Is Nothing Then
                    If Not tempNode Is Nothing Then
                        tempNode.Url += qs
                        tempNode.Title = GetMenuName(tempNode.Url)
                    End If
                End If
                tempNode = tempNode.ParentNode
            End While

            Return temp
        End Function

        Private Function GetMenuName(ByVal querystring As String) As String
            Dim ret As String = ""
            querystring = querystring.Substring(querystring.LastIndexOf("/") + 1)
            querystring = querystring.Substring(0, querystring.LastIndexOf("=") + 1)
            Dim sql As String = "Select Menuname from tbl_ROLES_Main where pagename=@pagename"
            Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
            Dim myreader As sqldatareader = Nothing
            Dim params As SqlParameter() = {New SqlParameter("@pagename", querystring)}
            Try
                myreader = CommonUtilsv2.GetDataReader(dbKey, sql, Data.CommandType.Text, params)
                While myreader.Read
                    ret = myreader(0).ToString
                End While
            Catch ex As Exception
                Throw ex
            End Try

            Return ret
        End Function

        Private Function GetQueryString(ByVal node As SiteMapNode, ByVal context As HttpContext) As String
            If node("queryStringToInclude") Is Nothing Then Return Nothing

            Dim values As NameValueCollection = New NameValueCollection
            Dim vars() As String = node("queryStringToInclude").Split(",".ToCharArray())
            Dim s As String

            For Each s In vars
                Dim var As String = s.Trim()
                If context.Request.QueryString(var) Is Nothing Then Continue For
                values.Add(var, context.Request.QueryString(var))
            Next

            If values.Count = 0 Then Return Nothing

            Return NameValueCollectionToString(values)
        End Function

        Private Function NameValueCollectionToString(ByVal col As NameValueCollection) As String
            Dim parts(col.Count - 1) As String
            Dim keys() As String = col.AllKeys

            For i As Integer = 0 To keys.Length - 1
                parts(i) = keys(i) & "=" & col(keys(i))
            Next

            Dim url As String = "?" & String.Join("&", parts)

            Return url
        End Function

    End Class

End Namespace