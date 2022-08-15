<%@ Page Language="VB" MasterPageFile="PageMaster.master" AutoEventWireup="false" EnableEventValidation="false" ViewStateEncryptionMode="Never" AspCompat="TRUE" %>
<%@ MasterType VirtualPath="PageMaster.master" %>
<%@ Register TagPrefix="obout" Namespace="Obout.Grid" Assembly="obout_Grid_NET" %>
<%@ Register Assembly="Obout.Ajax.UI" Namespace="Obout.Ajax.UI.HTMLEditor" TagPrefix="obout" %>
<%@ Register TagPrefix="obout" Namespace="Obout.Interface" Assembly="obout_Interface" %>
<%@ Register TagPrefix="obout" Namespace="OboutInc.Calendar2" Assembly="obout_Calendar2_NET" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.Page" %>
<%@ Import Namespace="Webapps.Utils" %>
<%@ Register Assembly="PdfViewer" Namespace="PdfViewer" TagPrefix="cc1" %>
<%@ Import Namespace="System.Web.Services" %>
<%@ Import Namespace="Newtonsoft.Json" %>
<%@ Import Namespace="Newtonsoft.Json.Linq" %>

<script language="VB" runat="server">
    Dim bRequirePWChange As Boolean = False
    Dim iNumberOfPriorPWNotbeUsed As Integer = ApplicationSettings.NumberOfPriorPWsNotBeUsed 'CommonUtilsv2.GetNumberOfPriorPWsNotBeUsed()

    Dim strTempMsg As String = ""

    '---- Error logging
    Dim pageName As String = "DocumnetCopy.aspx"
    Dim strRecepient As String = Webapps.Utils.ApplicationSettings.ErrorNoticeEmails
    Dim strFrom As String = Webapps.Utils.ApplicationSettings.ApplicationSourceEmail
    Dim environment As String = Webapps.Utils.ApplicationSettings.Environment
    Dim errLocation As String
    Dim errString As String
    Public strID As String = ""

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

        If CustomRoles.RolesForPageLoad() Then
            strID = Request.QueryString.Get("id")

            If Not IsPostBack Then

                If Not CommonUtilsv2.Validate(strID, CommonUtilsv2.DataTypes.Int, True, True, True) Then
                    Response.Redirect(Webapps.Utils.ApplicationSettings.Homepage, False)
                    Exit Sub
                End If

            End If
        Else
            CustomRoles.TransferIfNotInRole(True)
            Exit Sub
        End If
    End Sub

    <WebMethod()>
    Public Shared Function GetIndexsKey(ByVal ID As Integer) As String
        Dim StrUser As String = HttpContext.Current.Session("User")
        If StrUser Is Nothing Then
            Return "{""msg"":""You don't have access to this Function "",""status"":-1}"
        End If

        ' status 0:error,1:no data,2:success
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim ReturnStr As String = String.Empty
        Dim params As SqlParameter() = Nothing
        Dim ds As DataSet
        Dim strSQL As String = String.Empty

        strSQL = "  select IndexName,DisplayName,DateType from tbl_Web_SearchIndexNames as webindex where webindex.GroupID=(select GroupID from tbl_Web_Queues as webque where webque.ID=(select QueueID from tbl_Web_DocumentAttributes where DCN=@DCN) ) Order by SortOrder  "
        params = {New SqlParameter("@DCN", ID)}

        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)
            If Not ds Is Nothing Then
                Dim strjson As String = String.Empty
                For Each row As System.Data.DataRow In ds.Tables(0).Rows
                    strjson += "{""Index"":""" + row("IndexName") + """,""Name"":""" + row("DisplayName") + """,""Type"":""" + row("DateType") + """},"
                Next
                ReturnStr = "{""msg"":""this success!"",""status"":2,""data"":" + "[" & strjson.ToString().TrimEnd(",") & "]" + "}"
            Else

                ReturnStr = "{""msg"":""this no data!"",""status"":1}"

            End If

        Catch ex As Exception

            ReturnStr = "{""msg"":""" + ex.ToString() + """,""status"":0}"

        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try
        Return ReturnStr
    End Function

    <WebMethod()>
    Public Shared Function GetIndexsValue(ByVal ID As Integer) As String
        Dim StrUser As String = HttpContext.Current.Session("User")
        If StrUser Is Nothing Then
            Return "{""msg"":""You don't have access to this Function "",""status"":-1}"
        End If

        ' status 0:error,1:no data,2:success
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim ReturnStr As String = String.Empty
        Dim params As SqlParameter() = Nothing
        Dim ds As DataSet
        Dim strSQL As String = String.Empty

        strSQL = "select IndexName,DisplayName,DateType,REPLACE(REPLACE(TT.IndexValue,char(13),''),char(10),'')  as IndexValue from tbl_Web_SearchIndexNames as webindex left join ( select IndexValue,[Index] from tbl_Web_DocumentAttributes as C    UNPIVOT(IndexValue for [Index] in (Index1,Index2,Index3,Index4,Index5,Index6,Index7,Index8,Index9,Index10)) AS T  where DCN=@DCN ) as TT on TT.[Index]=webindex.IndexName where webindex.GroupID=(select GroupID from tbl_Web_Queues as webque where webque.ID=(select QueueID from tbl_Web_DocumentAttributes where DCN=@DCN) ) Order by SortOrder;  "
        strSQL += " Select ID,QueueName from tbl_Web_Queues where GroupID=(Select GroupID from tbl_Web_Queues  where ID=(Select QueueID from tbl_Web_DocumentAttributes where DCN=@DCN )) and RecordStatus=1;"
        strSQL += " select ID,DocTypeName from  tbl_Web_DocTypes where GroupID=(select GroupID from tbl_Web_Queues  where ID=(select QueueID from tbl_Web_DocumentAttributes where DCN=@DCN )) and RecordStatus=1;"
        strSQL += " select * from Tbl_Web_States;"
        params = {New SqlParameter("@DCN", ID)}

        Dim myReader As System.Data.SqlClient.SqlDataReader = Nothing
        Try

            ds = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)
            If Not ds Is Nothing Then
                Dim strjson As String = String.Empty
                Dim Indes7state As String = String.Empty 'Get the value of state to query county data  
                For Each row As System.Data.DataRow In ds.Tables(0).Rows
                    If (row("DateType").ToString().Trim().ToUpper() = "SELECT") Then
                        Dim Strchildjson As String = String.Empty
                        ''state
                        If (row("DisplayName").ToString().ToUpper().Contains("STATE")) Then
                            Indes7state = row("IndexValue").ToString().Trim()
                            For Each row1 As System.Data.DataRow In ds.Tables(3).Rows
                                Strchildjson += "{""Id"":""" + row1("state_name").Trim() + """,""Name"":""" + row1("state_name").Trim() + """},"
                            Next
                        End If
                        'County
                        If (row("DisplayName").ToString().ToUpper().Contains("COUNTY")) Then
                            If Not IsNothing(row("IndexValue").ToString()) Then
                                Dim dsCounty As DataSet = Nothing
                                strSQL = "select * from Tbl_Web_Counties where State='" + Indes7state + "'"
                                dsCounty = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text)
                                If Not IsNothing(dsCounty) Then
                                    For Each row2 As System.Data.DataRow In dsCounty.Tables(0).Rows
                                        Strchildjson += "{""Id"":""" + row2("State").Trim() + """,""Name"":""" + row2("County").Trim() + """},"
                                    Next
                                End If
                            End If
                        End If

                        strjson += "{""Index"":""" + row("IndexName") + """,""Name"":""" + row("DisplayName") + """,""Type"":""" + row("DateType") + """,""Value"":""" + row("IndexValue") + """,""data"":" + "[" & Strchildjson.ToString().TrimEnd(",") & "]" + "},"
                    Else
                        Dim StrValue As String = row("IndexValue").ToString().Replace("\", "\\").Replace("'", "&quit&").Replace("""", "!quot!").Replace(",", "%quot%").Replace(":", "-quot-").Trim()
                        strjson += "{""Index"":""" + row("IndexName").Trim() + """,""Name"":""" + row("DisplayName").Trim() + """,""Type"":""" + row("DateType").Trim() + """,""Value"":""" + StrValue + """},"
                    End If

                Next

                Dim strQueuejson As String = String.Empty
                For Each row As System.Data.DataRow In ds.Tables(1).Rows
                    strQueuejson += "{""Index"":""" + row("ID").ToString().Trim() + """,""Name"":""" + row("QueueName").Trim() + """},"
                Next

                Dim strdoctypejson As String = String.Empty
                For Each row As System.Data.DataRow In ds.Tables(2).Rows
                    strdoctypejson += "{""Index"":""" + row("ID").ToString().Trim() + """,""Name"":""" + row("DocTypeName").Trim() + """},"
                Next

                ReturnStr = "{""msg"":""this success!"",""status"":2,""data1"":" + "[" & strjson.ToString().TrimEnd(",") & "]" + ",""data2"":" + "[" & strQueuejson.ToString().TrimEnd(",") & "]" + ",""data3"":" + "[" & strdoctypejson.ToString().TrimEnd(",") & "]" + "}"

            Else

                ReturnStr = "{""msg"":""this no data!"",""status"":1}"

            End If

        Catch ex As Exception

            ReturnStr = "{""msg"":""" + ex.ToString() + """,""status"":0}"

        Finally
            If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
                myReader.Close()
            End If
        End Try
        Return ReturnStr
    End Function

    <WebMethod()>
    Public Shared Function UpdateDocumentList(ByVal StrID As Integer, ByVal StrData As Object) As String
        Dim StrUser As String = HttpContext.Current.Session("User")
        If StrUser Is Nothing Then
            Return "{""msg"":""You don't have access to this Function "",""status"":-1}"
        End If

        ' status 0:error,1:no data,2:success
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim ReturnStr As String = String.Empty
        'Dim params As SqlParameter() = Nothing
        'Dim ds As DataSet

        Try
            'Dim StrInsertAttachment As StringBuilder = New StringBuilder()
            'Dim strSQLAttachment As String = "select ID from tbl_Web_Attachments where DCN=@DCN"
            'params = {New SqlParameter("@DCN", StrID)}
            'ds = CommonUtilsv2.GetDataSet(dbKey, strSQLAttachment, CommandType.Text, params)
            'If Not ds Is Nothing Then
            '    For Each row As System.Data.DataRow In ds.Tables(0).Rows
            '        StrInsertAttachment.Append(String.Format(" INSERT INTO tbl_Web_Attachments ([FileName],[Description],DocumentSizeInMB) select [FileName],[Description],DocumentSizeInMB from  tbl_Web_Attachments  where ID={0} SET  @intAttac=@@IDENTITY;", row("ID")))
            '        StrInsertAttachment.Append(String.Format(" UPDATE tbl_Web_Attachments SET DCN=@intDeclare,[FileName]=cast(@intDeclare as varchar)+'-'+cast(@intAttac as varchar)+substring([FileName],charindex('.',[FileName]),len([FileName])+1-charindex('.',[FileName])),UploadedBy='{1}',UpdatedBy='{1}',ImageID={0}  where ID=@intAttac;", row("ID"), HttpContext.Current.Session("User")))
            '    Next
            'End If
            Dim params As SqlParameter() = Nothing
            Dim dsIndex As DataSet
            Dim StrInsertAttachment As StringBuilder = New StringBuilder()
            Dim strSQLAttachment As String = "select IndexName,ColumnName from tbl_Web_SearchIndexNames where  GroupID=(select GroupID from tbl_Web_Queues where ID=(select QueueID from tbl_Web_DocumentAttributes where DCN=@DCN))"
            params = {New SqlParameter("@DCN", StrID)}
            dsIndex = CommonUtilsv2.GetDataSet(dbKey, strSQLAttachment, CommandType.Text, params)

            Dim Jar As JArray = New JArray()
            Try
                Jar = JArray.Parse(StrData.Replace("\", "\\"))
            Catch ex As Exception
                StrData = Nothing
            End Try
            If StrData Is Nothing Then
                Return "{""msg"":""Please fill in the value"",""status"":0}"
            End If


            Dim strSQL As StringBuilder = New StringBuilder()
            'strSQL.Append("declare @intDeclare int;declare @intDeclareAttac int;declare @intAttac int;")
            strSQL.Append("declare @intDeclare int;declare @intDeclareAttac int;declare @intAllDeclare varchar(1000);set @intAllDeclare='';")
            For Each Obj As JObject In Jar
                Dim StrCreateLog As String = "Document copied from DCN " & StrID.ToString & ":  <br>"
                Dim Strqueue As Integer = 0
                Dim StrDocType As Integer = 0
                Dim StrSQLIndees As String = String.Empty
                Dim JObProperty As JToken = Obj
                For Each item As JProperty In JObProperty
                    If Not String.IsNullOrEmpty(item.Value) Then
                        If (item.Name.ToUpper() = "QUEUEID") Then
                            Strqueue = item.Value
                        ElseIf (item.Name.ToUpper() = "DOCTYPE") Then
                            StrDocType = item.Value
                        Else
                            Dim StrValue As String = item.Value.ToString().Replace("&quit&", "''").Replace("!quot!", """").Replace("%quot%", ",").Replace("-quot-", ":").Replace("#quot#", "\\").Replace("&nbsp;&nbsp;", "\t").Trim()
                            StrSQLIndees += "," + item.Name + "=N'" + StrValue + "'"
                            Dim IndexName As String = dsIndex.Tables(0).Select(" IndexName='" & item.Name.ToString() & "'")(0)("ColumnName").ToString()
                            StrCreateLog = StrCreateLog & IndexName & "=>" & " Create " & StrValue & ","
                        End If
                    End If
                Next

                strSQL.Append(String.Format(" INSERT INTO tbl_Web_Documents([FileName],ImageBlob,UploadedBy,UpdatedBy,ImageDCN) SELECT [FileName],ImageBlob,'{1}','{1}',{0} from tbl_Web_Documents where DCN={0}  SET  @intDeclare=@@IDENTITY; ", StrID, HttpContext.Current.Session("User")))
                strSQL.Append(" UPDATE tbl_Web_Documents set[FileName]=cast(@intDeclare as varchar)+substring([FileName],charindex('.',[FileName]),len([FileName])+1-charindex('.',[FileName]))  where DCN=@intDeclare;")
                strSQL.Append(String.Format(" INSERT INTO tbl_Web_DocumentAttributes (DCN,QueueID,DocType,Category,[Status],[Priority],[Description],DisplaySequence,QueueStartDate,Comments,CommentsExport,[Owner]) select @intDeclare,{1},{2},Category,[Status],[Priority],[Description],DisplaySequence,getDate(),Comments,CommentsExport,'{3}' from tbl_Web_DocumentAttributes where DCN={0} SET  @intDeclareAttac=@@IDENTITY;", StrID, Strqueue, StrDocType, HttpContext.Current.Session("User")))
                strSQL.Append(String.Format(" UPDATE tbl_Web_DocumentAttributes SET UploadedBy='{2}',UpdatedDate=getDate(),UpdatedBy='{2}' {3} where  ID=@intDeclareAttac; ", Strqueue, StrDocType, HttpContext.Current.Session("User"), StrSQLIndees))

                strSQL.Append(String.Format(" INSERT INTO tbl_Web_DCNLog([DCN],[LogType],[Comments],[UserID],[UserIP],[CreatedBy],[UpdatedDate],[UpdatedBy])VALUES(@intDeclare,{0},'{1}','{2}','{3}','{2}',getDate(),'{2}');", 2, StrCreateLog.TrimEnd(",") & "<br>", StrUser, HttpContext.Current.Request.UserHostAddress()))

                strSQL.Append(" Set @intAllDeclare=@intAllDeclare + CONVERT(varchar(50),@intDeclare) + ','; ")

                'strSQL.Append(StrInsertAttachment.ToString())

            Next
            'dbKey, strSQLAttachment, CommandType.Text, params'
            strSQL.Append(" select @intAllDeclare; ")
            Dim ReturnData As String = CommonUtilsv2.RunScalarQuery(dbKey, strSQL.ToString(), CommandType.Text)
            If (Not String.IsNullOrEmpty(ReturnData)) Then
                ReturnStr = "{""msg"":""Copying documents is successful!"",""status"":1,""data"":""" & ReturnData & """}"
            Else
                ReturnStr = "{""msg"":""Copying documents is successful,But did not return the modified value !"",""status"":1,""data"":""""}"
            End If

        Catch ex As Exception
            ReturnStr = "{""msg"":""" + ex.ToString() + """,""status"":0}"
        Finally
            'If Not myReader Is Nothing AndAlso Not myReader.IsClosed Then
            '    myReader.Close()
            'End If
        End Try
        Return ReturnStr
    End Function

</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
<script language="javascript" type="text/javascript" src="includes/js/My97DatePicker/WdatePicker.js"></script>
     <script type="text/javascript">
         var Strid =<%=strID %>;
         var leti=1;
         var strNO=1;
         var noupdate=false;//Not editable:ture,editable:false
         function backfunction() {
             if (window.history.length > 1) {
                 window.history.go(-1);
             } else {
                 window.opener.location.reload();
             }
         }

         $(function () {
             GetHeadHtml();//get head table
             GetBodyHtml();//get body table

             $("#DocumentTable").on("change", '.selectstate', function () {
                 var $this=$(this);
                 var strid = $this.find("option:selected").val();
                 if (strid) {
                     $.ajax({
                         type: "Post",
                         contentType: "application/json;charset=UTF-8",
                         url: "/Document-Detail.aspx/GetCountiesSelect",
                         data: "{ID:'" + strid + "'}",
                         success: function (result) {
                             var strjson = JSON.parse(result.d);
                             if(strjson.status ==-1){
                                 window.location.href="/login.aspx";
                             }
                             var strchildoption = "<option value=''>select a value</option>";
                             if (strjson.status == 2) {
                                 $(strjson.data).each(function (i, dom) {
                                     strchildoption += "<option value='" + dom.Name + "'>" + dom.Name + "</option>";
                                 });                           
                             }
                             $this.parent().parent().find(".selectcounty").html(strchildoption);
                         }
                     });
                 }else {
                     var strchildoption = "<option value=''>select a value</option>";
                     $this.parent().parent().find(".selectcounty").html(strchildoption);
                 }
             });
         });
         
         function GetHeadHtml() {
             $.ajax({
                 type: "Post",
                 contentType: "application/json;charset=UTF-8",
                 url: "/DocumnetCopy.aspx/GetIndexsKey",
                 data: "{ID:" + Strid + "}",
                 success: function (result) {
                     var strjson = JSON.parse(result.d);
                     if(strjson.status ==-1){
                         window.location.href="/login.aspx";
                     }
                     if (strjson.status == 2) {
                         var strtr = "<tr>"
                         strtr += "<td style='padding: 0px 5px;'>NO</td><td  style='padding: 0px 20px;'>operating</td><td>Queue</td><td>doctype</td>"
                         $(strjson.data).each(function (i, dom) {
                             strtr += "<td rel=" + dom.Index + ">" + dom.Name + "</td>";
                         });
                         strtr += "</tr>"
                         
                         $("#DocumentTable thead").html(strtr);
                     }
                 }
             });
         }

         function GetBodyHtml()
         {
             $.ajax({
                 type: "Post",
                 contentType: "application/json;charset=UTF-8",
                 url: "/DocumnetCopy.aspx/GetIndexsValue",
                 data: "{ID:" + Strid + "}",
                 success: function (result) {
                     var strjson = JSON.parse(result.d.replace(/(\t)+/g,'\\t'));
                     if(strjson.status ==-1){
                         window.location.href="/login.aspx";
                     }
                     if (strjson.status == 2) {
                         var strtr = "<tr>"
                         strtr += "<td>" + strNO + "</td>"//strNO
                         strtr += "<td><a href='javascript:void(0);' onclick='GetBodyHtml()' style='color:red;' title='add'><img src='/Images/Add01.png' style='width:12px;'></a><a href='javascript:void(0);' onclick='deletetd(this)' style='color:red;margin-left:10px;' title='delete'><img src='/Images/delete16.png' style='width:12px;'></a><a href='javascript:void(0);' onclick='copytd(this)' style='color:red;margin-left:10px;' title='copy'><img src='/Images/copy01.png' style='width:12px;'></a></td>"
                         var strqueue="<select>";
                         $(strjson.data2).each(function (i, dom) {
                             strqueue += "<option value=" + dom.Index + ">" + dom.Name + "</option>";
                         });
                         strqueue += "</select>";
                         strtr += "<td rel='QueueID'>" + strqueue + "</td>"
                         var strdoctype = "<select>";
                         $(strjson.data3).each(function (i, dom) {
                             strdoctype += "<option value=" + dom.Index + ">" + dom.Name + "</option>";

                         });
                         strdoctype += "</select>";
                         strtr += "<td  rel='DocType'>" + strdoctype + "</td>"
                         $(strjson.data1).each(function (i, dom) {
                             strtr += "<td rel=" + dom.Index + ">";
                             if (dom.Type.toUpperCase().trim() == "SELECT") {
                                 var strchildoption = "<option value=''>select a value</option>";
                                 var stronclicclass = "";
                                 //state
                                 if (dom.Name.toUpperCase().trim().indexOf("STATE") >= 0) {
                                     stronclicclass="class='selectstate'"
                                     $(dom.data).each(function (i1, dom1) {
                                         if (dom.Value.trim() == dom1.Id) {
                                             strchildoption += "<option value='" + dom1.Id + "' selected='selected'>" + dom1.Name + "</option>";
                                         } else {
                                             strchildoption += "<option value='" + dom1.Id + "'>" + dom1.Name + "</option>";
                                         }
                                     });
                                 }
                                 //county
                                 if (dom.Name.toUpperCase().trim().indexOf("COUNTY") >= 0) {
                                     stronclicclass = "class='selectcounty'"
                                     $(dom.data).each(function (i2, dom2) {
                                         if (dom.Value.trim() == dom2.Name) {
                                             strchildoption += "<option value='" + dom2.Name + "'  selected='selected'>" + dom2.Name + "</option>";
                                         } else {
                                             strchildoption += "<option value='" + dom2.Name + "'>" + dom2.Name + "</option>";
                                         }
                                     });
                                 }

                                 strtr += "<select id='tb" + dom.Index + "' style='width:120px;'" + stronclicclass + "  rel='" + dom.Index + "'>" + strchildoption + "<select>";

                             } else if (dom.Type.toUpperCase().trim() == "DATE") {
                                 strtr += "<input  name='" + dom.Index + leti + "' type='text' readonly='readonly' id='" + dom.Index + "_" + leti + "' maxlength='30'  class='Wdate' onclick=\"WdatePicker({lang:\'en\',dateFmt:\'MM/dd/yyyy\'})\"  value='" + dom.Value + "' style='width:120px;'>"
                             } else {
                                 strtr += "<input name='" + dom.Index + leti + "' type='text'  id='" + dom.Index + "_" + leti + "' maxlength='30'  value='" + ReturnTrunchar(dom.Value) + "' style='width:120px;'>"
                             }
                             strtr += "</td>";
                         });
                         strtr += "</tr>"
                         
                         $("#DocumentTable tbody").append(strtr);
                         leti++;
                         let ii=1;
                         $("#DocumentTable tbody tr").each(function(){
                             $(this).children("td:eq(0)").text(ii);
                             ii++
                         });
                         strNO=ii;
                     }
                 }
             });
         }

         function deletetd(_obj) {
             if($("#DocumentTable tbody tr").length>1){
                 $(_obj).parent().parent().remove();
                 let ii=1;
                 $("#DocumentTable tbody tr").each(function(){
                     $(this).children("td:eq(0)").text(ii);
                     ii++
                 });
                 strNO=ii;
             }else
             {
                 alert("There is only one left and cannot be deleted!");
             }
         }
         function copytd(_obj) {
             $("#DocumentTable tbody").append($(_obj).parent().parent().clone());
             $(_obj).parent().parent().find("select").each(function(j,dom){
                 $("#DocumentTable tbody").find("tr:last").find("select:eq("+j+")").val($(dom).val());
             });

             let ii=1;
             $("#DocumentTable tbody tr").each(function(){
                 $(this).children("td:eq(0)").text(ii);
                 ii++
             });
             strNO=ii;
             $(_obj).parent().parent().next().find("input").each(function(){                
                 $(this).attr("name",$(this).parent().attr("rel")+leti);
                 $(this).attr("id",$(this).parent().attr("rel")+"_"+leti);
             })
             leti++;
         }

         function submitdocument()
         { 
             noupdate=false;
             $(".a-submit").attr("style","display:none;");
             $(".a-submiting").attr("style","background: #e0dbdb;display:'';");
             $(".lableprompt").remove();//'Delete star first 
             var insertjson={};
             var list = new Array();
             $("#DocumentTable tbody tr").each(function(){
                 var $thistr=$(this)
                 var insertchildrenjson={};
                 $thistr.find("select").each(function(){
                     if($(this).val()){
                         var strkey=$(this).parent().attr("rel")
                         insertchildrenjson[strkey]=$(this).val();
                     }else
                     {
                         $(this).parent().append("<lable class='lableprompt'>*<span>Selection cannot be empty!</span></lable>");//
                         noupdate=true;
                     }
                 });
                 $thistr.find("input").each(function(){
                     var strkey=$(this).parent().attr("rel")
                     if($(this).val()){                        
                         if($(this).hasClass("Wdate"))
                         {
                             if(checkTime($(this).val()))
                             {
                                 insertchildrenjson[strkey]=$(this).val();
                             }else
                             {
                                 $(this).parent().append("<lable class='lableprompt'>*<span>Must be in date format!</span></lable>");//
                                 noupdate=true;
                             }
                         }else
                         {
                             if(!checkTime($(this).val()))
                             {
                                 insertchildrenjson[strkey] = ReturnToPunctuation(qudiaoAll($(this).val()));
                             }else
                             {
                                 $(this).parent().append("<lable class='lableprompt'>*<span>Cannot be in date format !</span></lable>");//
                                 noupdate=true;
                             }
                         }
                     }else
                     {
                         insertchildrenjson[strkey]="";
                     }
                 });
                 
                 list.push(insertchildrenjson);
             });
 
             if(!noupdate && list.length>0)
             { //ajax
                 $.ajax({
                     type: "Post",
                     contentType: "application/json;charset=UTF-8",
                     url: "/DocumnetCopy.aspx/UpdateDocumentList",
                     data: "{StrID:" + Strid + ",'StrData':'"+ JSON.stringify(list) +"'}",
                     success: function (result) {
                         var strjson = JSON.parse(result.d);
                         if(strjson.status ==-1){
                             window.location.href="/login.aspx";
                         }
                         if (strjson.status == 1) {
                             $("#promptcontent").html(strjson.msg+" Please do not submit again !");
                             $("#returndata").html("The batch is successful. Their dcn is:<lable style='color:#009688'>&nbsp;&nbsp;"+ strjson.data  +"</lable>&nbsp;&nbsp;<lable style='color:red;'>Please save it and refresh or leave the page will not exist </lable>");
                             $(".a-submit").attr("style","display:'';");
                             $(".a-submiting").attr("style","background: #e0dbdb;display:none;");
                         }else
                         {
                             $("#promptcontent").html(strjson.msg);
                             $(".a-submit").attr("style","display:'';");
                             $(".a-submiting").attr("style","background: #e0dbdb;display:none;");
                         }
                     }
                 });                  
             }else
             {
                 $("#promptcontent").html("The parameter is empty or there is a star error!");
                 $(".a-submit").attr("style","display:'';");
                 $(".a-submiting").attr("style","background: #e0dbdb;display:none;");
             }
         }

         function checkTime(strdate){
             var date_regex = /^(0[1-9]|1[0-2])\/(0[1-9]|1\d|2\d|3[01])\/(19|20)\d{2}$/ ;
             return date_regex.test(strdate);
         }
         function ReturnTrunchar(_Strobj)
         {
             if(_Strobj.indexOf("&quit&")!=-1)
             {
                 var reg = new RegExp("&quit&","g");
                 _Strobj=_Strobj.replace(reg, '&apos;');
             }
             if(_Strobj.indexOf("!quot!")!=-1)
             {
                 var reg = new RegExp("!quot!","g");
                 _Strobj=_Strobj.replace(reg, '\"');
             }
             if(_Strobj.indexOf("%quot%")!=-1)
             {  
                 var reg = new RegExp("%quot%","g");
                 _Strobj=_Strobj.replace(reg, ',');
             }
             if(_Strobj.indexOf("-quot-")!=-1)
             {
                 var reg = new RegExp("-quot-","g");
                 _Strobj=_Strobj.replace(reg, ':');
             }

             if(_Strobj.indexOf("#quot#")!=-1)
             {
                 var reg = new RegExp("#quot#","g");
                 _Strobj=_Strobj.replace(reg, '\\');
             }
             return _Strobj;
         }

         function ReturnToPunctuation(_Strobj){
             //valstr.replace("'", "&quot&").replace("\"", "!quot!").replace(",", "%quot%").replace(":", "-quot-").replace("\\", "#quot#")
             if(_Strobj.indexOf("'") !=-1)
             {
                 var reg = new RegExp("'","g");
                 _Strobj=_Strobj.replace(reg, '&quot&');
             }
             if (_Strobj.indexOf(/\"/g) != -1) {
                 //var reg = new RegExp("\"","g");
                 _Strobj = _Strobj.replace(/\"/g, '!quot!');
             }
             if(_Strobj.indexOf("\"")!=-1)
             {
                 var reg = new RegExp("\"","g");
                 _Strobj=_Strobj.replace(reg, '!quot!');
             }
             if(_Strobj.indexOf(",")!=-1)
             {  
                 var reg = new RegExp(",","g");
                 _Strobj=_Strobj.replace(reg, '%quot%');
             }
             if(_Strobj.indexOf(":")!=-1)
             {
                 var reg = new RegExp(":","g");
                 _Strobj=_Strobj.replace(reg, '-quot-');
             }

             //if(_Strobj.indexOf("\\")!=-1)
             //{
             //    var reg = new RegExp("\\","g");
             //    _Strobj=_Strobj.replace(reg, '#quot#');
             //}
             return _Strobj;
         }

         function qudiaoAll(_Strobj) {
             _Strobj = _Strobj.replace(/\\r/g, "").replace(/\\n/g, "").replace(/\\t/g, "").replace(/\\b/g, "").replace(/\\f/g, "");
             _Strobj = _Strobj.replace(/\r/g, "").replace(/\n/g, "").replace(/\t/g, "").replace(/\b/g, "").replace(/\f/g, "");

             return _Strobj;
         }


    </script>
    <div align="center" style="height: 100%">
            <div class="body_title">
                <a href="javascript:void(0);" onclick="backfunction()"><span>back</span></a>
                Document Copy
            </div>
            <div class="document_DCN">
             <span style="color:black;">COPY DCN&nbsp;&nbsp;</span>Move Document DCN&nbsp;&nbsp;(<span><%=strID %></span>)
            </div>
            <div class="document_DCN">
                <a href="javascript:void(0);" onclick="submitdocument()"  class="a-submit">submit</a>
                <a href="#" class="a-submiting" style="background: #e0dbdb;display:none;">submiting</a>
            </div>
             <div style="clear:both;"></div>
            <div class="div-doccopyote" >
              <span  class="doccopynote" >* Note: If the submission is successful, don't click submit again. Otherwise, it will be submitted twice.</span><br />
              <span id="promptcontent" style="color:blue;"></span>
              <p id="returndata" style="margin-top: 20px;"></p>
            </div>
            <div style="width:98%;overflow-x:auto;">
            <table id="DocumentTable" width="100%"  border="1" cellpadding="0" cellspacing="0" class="wrapper Documentcopy" >
              <thead></thead>
              <tbody></tbody>
           </table>  
           </div>     
    </div>
</asp:Content>
