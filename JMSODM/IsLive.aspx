<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Webapps.Utils" %>
<script language="VB" runat="server">
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim strSQL As String = " select count(*) from tbl_Apps_AppSettings "
        Try
            Dim CountStr As Integer = StrHelp.GetInt(CommonUtilsv2.RunScalarQuery(dbKey, strSQL, CommandType.Text, Nothing))
            If (CountStr <= 0) Then
                bodymain.InnerText = "Apps_AppSettings table no date!"
            Else
                bodymain.InnerText = "Site is live."
            End If

        Catch ex As Exception
            bodymain.InnerText = ex.ToString()
        End Try
    End Sub
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>   
    <title></title>
</head>
<body runat="server" id="bodymain"> 
</body>
</html>
