<%@ Page Language="VB" MasterPageFile="PageMaster.master" AutoEventWireup="false" EnableEventValidation="false" viewStateEncryptionMode="Auto" ASPCOMPAT="TRUE" %>
<%@ MasterType VirtualPath="PageMaster.master" %>
<%@ Register TagPrefix="obout" Namespace="Obout.Grid" Assembly="obout_Grid_NET" %>
<%@ Register TagPrefix="oint" Namespace="Obout.Interface" Assembly="obout_Interface" %>
<%@ Register TagPrefix="obout" Namespace="Obout.ComboBox" Assembly="obout_ComboBox" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Web.UI.Page" %>
<%@ Import Namespace="Webapps.Utils" %>

<script language="VB" runat="server">
    '************* Error logging Section ***********************
    Dim pageName As String
    Dim errLocation As String
    Dim errString As String
    Dim dsLookupTypeSet As DataSet = Nothing


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
        Session("PreviousPage") = "SiteManager.aspx"
        Dim strLogout As String = "False"
        strLogout = Request.QueryString.Get("logout")
        If strLogout = "True" Then
            Session("Level") = "0"
        End If

        If CustomRoles.RolesForPageLoad() Then
            PopulateEditableLookupTypes()
            CreateGrid()

        Else
            CustomRoles.TransferIfNotInRole()
            Exit Sub
        End If
        '  PopulateEditableLookupTypes()
    End Sub

    Dim strCommandBase As String = "SELECT WL1.ID, WL1.LookupType, WL1.LookupCode, WL1.LookupDesc, WL1.Comments FROM tbl_web_Lookup WL1 " & _
                                      "WHERE WL1.LookupType in (Select WL2.LookupCode from tbl_WEB_Lookup WL2 where WL2.LookupType = 'LookupEditable') " & _
                                      " AND WL1.RecordStatus=1 " & strCommandOrderBy
    Dim strCommandOrderBy As String = " ORDER BY LookupType ASC, LookupCode ASC"
    Dim ComboBox1 As ComboBox

    Private Sub CreateGrid()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim sqlDR As SqlDataReader
        Try
            sqlDR = CommonUtilsv2.GetDataReader(dbKey, strCommandBase, CommandType.Text)
            gridLookupTypes.DataSource = sqlDR
            gridLookupTypes.DataBind()
        Catch ex As Exception
            errString = ex.Message
            errLocation = "CreateGrid()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try
    End Sub

    Private Sub LoadDocTypes()
        'Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        'Dim strSQL As String = " SELECT LookupCode, LookupValue as DocType From tbl_web_Obj_Lookup WHERE LookupType='Obj_Document_Type' Order by LookupCode "
        'Try
        '    dsDocType = CommonUtis.GetDataSet(dbKey, strSQL, CommandType.Text)
        '    Dim ddDocTypes As DropDownList = CType(gridObjDocs.Templates(0).Container.FindControl("ddDocTypes"), DropDownList)
        '    ddDocTypes.DataSource = dsDocType
        '    ddDocTypes.DataBind()
        'Catch ex As Exception
        '    errString = ex.Message
        '    errLocation = "LoadDocTypes()"
        '    LogError(errString, errLocation)
        'Finally
        'End Try
    End Sub

    Private Sub PopulateEditableLookupTypes()
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim strSQL As String = "SELECT LookupType, LookupCode FROM tbl_WEB_Lookup WHERE LookupType Like @LookupType ORDER BY LookupCode "
        'Dim strSQL As String = "SELECT distinct LookupCode FROM tbl_WEB_Lookup WHERE LookupType ='WebPageName' And RecordStatus=1 ORDER BY LookupCode"

        Dim params As SqlParameter() = {New SqlParameter("@LookupType", "LookupEditable")}
        Try
            dsLookupTypeSet = CommonUtilsv2.GetDataSet(dbKey, strSQL, CommandType.Text, params)
            Dim ddLookupTypeValues As DropDownList = CType(gridLookupTypes.Templates(0).Container.FindControl("ddLookupType"), DropDownList)
            ddLookupTypeValues.DataSource = dsLookupTypeSet
            ddLookupTypeValues.DataBind()
            ddLookupTypeValues.Items.Insert(0, New ListItem("Please Select"))
        Catch ex As Exception
            errString = ex.Message
            errLocation = "PopulateEditableLookupTypes()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        Finally
        End Try
    End Sub

    Sub lookupInsert(ByVal sender As Object, ByVal e As GridRecordEventArgs)
        If Not String.IsNullOrEmpty(e.Record("LookupType")) AndAlso Not String.Compare("Please Select", e.Record("LookupType"), True) = 0 Then
        Else
            Exit Sub
        End If
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strInsert As String = "prc_InsertLookup"

        Dim strLUType As String = e.Record("LookupType")
        Dim strLUCode As String = e.Record("LookupCode")
        Dim strLUDesc As String = e.Record("LookupDesc")
        Dim strComments As String = e.Record("Comments")
        Dim strLastUpdatedBy As String = Session("User")

        Dim params As SqlParameter() = { _
            New SqlParameter("@LookupType", strLUType), _
            New SqlParameter("@LookupCode", strLUCode), _
            New SqlParameter("@LookupDesc", strLUDesc), _
            New SqlParameter("@Comments", strComments), _
            New SqlParameter("@LastUpdatedBy", strLastUpdatedBy) _
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strInsert, CommandType.StoredProcedure, params)
            CreateGrid()
        Catch ex As Exception
            errString = ex.Message
            errLocation = "lookupInsert()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try
    End Sub

    Sub lookupUpdate(ByVal sender As Object, ByVal e As GridRecordEventArgs)
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strUpdate As String = "prc_UpdateLookup"

        Dim strID As String = e.Record("ID")
        Dim strLUType As String = e.Record("LookupType")
        Dim strLUCode As String = e.Record("LookupCode")
        Dim strLUDesc As String = e.Record("LookupDesc")
        Dim strComments As String = e.Record("Comments")
        Dim strLastUpdatedBy As String = Session("User")

        Dim params As SqlParameter() = { _
            New SqlParameter("@ID", strID), _
            New SqlParameter("@LookupType", strLUType), _
            New SqlParameter("@LookupCode", strLUCode), _
            New SqlParameter("@LookupDesc", strLUDesc), _
            New SqlParameter("@Comments", strComments), _
            New SqlParameter("@LastUpdatedBy", strLastUpdatedBy) _
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strUpdate, CommandType.StoredProcedure, params)
        Catch ex As Exception
            errString = ex.Message
            errLocation = "lookupUpdate()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try
    End Sub

    Sub lookupDelete(ByVal sender As Object, ByVal e As GridRecordEventArgs)
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")
        Dim strDelete As String = "Update tbl_Web_Lookup Set RecordStatus=0 Where LookupType = @LookupType AND LookupCode = @LookupCode"

        Dim strLUType As String = e.Record("LookupType")
        Dim strLUCode As String = e.Record("LookupCode")

        Dim params As SqlParameter() = { _
            New SqlParameter("@LookupType", strLUType), _
            New SqlParameter("@LookupCode", strLUCode) _
            }
        Try
            CommonUtilsv2.RunNonQuery(dbKey, strDelete, CommandType.Text, params)
        Catch ex As Exception
            errString = ex.Message
            errLocation = "lookupDelete()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), Request.UserHostAddress())
        End Try
    End Sub

    Private Sub RebindGrid(ByVal sender As Object, ByVal e As EventArgs)
        CreateGrid()
    End Sub

</script>

<asp:Content ID="home1" runat="server" ContentPlaceHolderID="Body">
    <script type="text/javascript">
        function onBeforeClientDelete(record) {
            if (confirm("Are you sure you want to delete the record with the LookupCode " + record.LookupCode + "?") == false) {
                return false;
            }
            return true;
        }

        function onClientUpdate(record) {
            var vLookupType = record.LookupType;
          //  alert(vLookupType);
            bValid = Boolean(true);
            if (vLookupType === "" || vLookupType=="Please Select" ) {
                bValid = false;
            }
            if (bValid == false ) {
                alert("Please enter a valid lookup type and re-submit.");
                return false;
            }
            else {
                return true;
            }
        }
    </script>
    <div align="center">
        <table id="Table1" width="100%" runat="server" border="0" cellpadding="0" cellspacing="0" class="wrapper">
          <%--<tr>
             <td colspan="2" align="left" class="body_title" >
                Manage Lookup Values
             </td>           
          </tr>--%>
          <tr><td>&nbsp;</td></tr>
                      <tr>
                <td align="left" colspan="2"  >
                    <obout:Grid id="gridLookupTypes" runat="server" CallbackMode="true" Serialize="true" AutoGenerateColumns="false"
                        PageSizeOptions="1,5,10,15,20" PageSize="20"
                        Width="100%" FolderStyle="styles/grand_graydark" EnableRecordHover="true" AllowSorting="false"
                        AllowRecordSelection="false" AllowMultiRecordSelection="false" KeepSelectedRecords="false"
                        AllowAddingRecords="true" AllowFiltering="true" ShowLoadingMessage="true"
                        AllowGrouping="true" ShowGroupsInfo="false" ShowColumnsFooter="false"
                        GroupBy="LookupType" OnRebind="RebindGrid"
                        ShowGroupFooter="false" OnUpdateCommand="lookupUpdate" OnInsertCommand="lookupInsert" OnDeleteCommand="lookupDelete">
                        <ClientSideEvents OnBeforeClientDelete="onBeforeClientDelete" OnBeforeClientUpdate="onClientUpdate" OnBeforeClientInsert="onClientUpdate"/>
                        <Columns>
                            <obout:Column ID="Column5" DataField="ID" Align="left" ReadOnly="true" Width="0%" Wrap="true" Visible="false" AllowFilter="false" />
                            <obout:Column ID="Column1" HeaderText="Lookup Type" DataField="LookupType" Align="left" ReadOnly="false" Width="15%" Wrap="true" Visible="true" AllowFilter="false">
                                <TemplateSettings EditTemplateId="tplEditLookupType"/>
                            </obout:Column>
                            <obout:Column ID="Column2" HeaderText="Lookup Code" DataField="LookupCode" Align="left" ReadOnly="false" Width="15%" Visible="true" AllowFilter="true">
                                <TemplateSettings EditTemplateId="tplLUCode"/>
                            </obout:Column>
                            <obout:Column ID="Column3" HeaderText="Lookup Description" DataField="LookupDesc" Align="left" ReadOnly="false" Width="30%" Visible="true" AllowFilter="true">
                                <TemplateSettings EditTemplateId="tplLUDesc"/>
                            </obout:Column>
                            <obout:Column ID="Column4" DataField="Comments" Align="left" ReadOnly="false" Width="25%" Visible="true" AllowFilter="true" Wrap="true">
                                <TemplateSettings EditTemplateId="tplComments"/>
                            </obout:Column>
                            <obout:Column ID="Edit" HeaderAlign="center" HeaderText="Edit" Width="15%" AllowEdit="true" AllowDelete="true" runat="server" />
                        </Columns>
                        <Templates>
                                <obout:GridTemplate runat="server" ID="tplEditLookupType" ControlID="ddLookupType" ControlPropertyName="value">
                                        <Template>
                                            <asp:DropDownList runat="server" ID="ddLookupType" DataValueField="LookupCode" DataTextField="LookupCode" height="20px" CssClass="ob_gEC">
                                            </asp:DropDownList>
                                        </Template>
                                 </obout:GridTemplate>                        
                                <obout:GridTemplate runat="server" ID="tplLUType" ControlID="ddLUType" ControlPropertyName="value">
                                    <oint:OboutDropDownList runat="server" ID="ddLUType" Width="100%">
                                        <asp:ListItem>Debtor</asp:ListItem>
                                        <asp:ListItem>MajorMinor</asp:ListItem>
                                        <asp:ListItem>eRoomCategory</asp:ListItem>
                                    </oint:OboutDropDownList>
                                </obout:GridTemplate>

                                <obout:GridTemplate runat="server" ID="tplLUCode" ControlID="txtLUCode" ControlPropertyName="value">
                                    <oint:OboutTextBox runat="server" ID="txtLUCode" Width="100%" />
                                </obout:GridTemplate>
                                <obout:GridTemplate runat="server" ID="tplLUDesc" ControlID="txtLUDesc" ControlPropertyName="value">
                                    <oint:OboutTextBox runat="server" ID="txtLUDesc" Width="100%" />
                                </obout:GridTemplate>
                                <obout:GridTemplate runat="server" ID="tplComments" ControlID="txtComments" ControlPropertyName="value">
                                    <oint:OboutTextBox runat="server" ID="txtComments" Width="100%" />
                                </obout:GridTemplate>
                         </Templates>
                    </obout:Grid>
                </td>
            </tr>
        </table>
    </div>
</asp:Content>