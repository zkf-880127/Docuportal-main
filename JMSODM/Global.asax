<%@ Application Language="VB" %>
<%@ Import Namespace="Webapps.Utils" %>
<script runat="server">

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        Application("CulmulativeSessionCount") = 0
        
        Dim lFSizeLimit As Long = CommonUtilsv2.GetAllowedUploadFileSize()
        Dim sFSizeLimit As Single = lFSizeLimit / (1024 * 1024.0)
        Dim strDocSizeLimit As String = FormatNumber(sFSizeLimit, 2, TriState.True, TriState.False, TriState.True) & "MB"
        Application("AllowedUploadFileSize") = lFSizeLimit
        Application("AllowedUploadFileSizeInMB") = strDocSizeLimit
        Application("AllowedMaxExportCount") = CommonUtilsv2.GetAllowedMaxExportCount()
        Application("AllowedUploadFiletypes") = CommonUtilsv2.GetAllowedUploadFiletypes()
        ' Code that runs on application startup
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application shutdown
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when an unhandled error occurs
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when a new session is started
        Application.Lock()  ' Its used to open only one time  for session,suppose the user do referesh page  it will not count the counter.
        Application("CulmulativeSessionCount") = Integer.Parse(Application("CulmulativeSessionCount").ToString) + 1
        If Integer.Parse(Application("CulmulativeSessionCount").ToString) Mod 2 = 0 Then
            ApplicationSettings.ReLoadSettings()
        End If
        Application.UnLock()
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when a session ends. 
        ' Note: The Session_End event is raised only when the sessionstate mode
        ' is set to InProc in the Web.config file. If session mode is set to StateServer 
        ' or SQLServer, the event is not raised.
        '  Dim sessionId As String = Session.SessionID
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' HttpContext.Current.Response.AddHeader("x-frame-options", "DENY")
    End Sub

</script>