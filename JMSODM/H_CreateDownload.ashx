<%@ WebHandler Language="VB" Class="H_CreateDownload" %>

Imports System
Imports System.IO
Imports System.Web
Imports System.Web.SessionState
Imports System.Data
Imports System.Data.SqlClient
Imports Webapps.Utils
Imports ExportToExcel


Public Class H_CreateDownload : Implements IHttpHandler, IRequiresSessionState
    Dim aColumnModel() As ColumnModel = Nothing
    Dim errString As String = ""
    Dim errLocation As String = ""
    Dim pageName As String = "H_CreateDownload.ashx"

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim bytearr As Byte() = Nothing
        Dim enc As Encoding = Encoding.UTF8
        Dim exportFilename As String = context.Request.QueryString("filename")
        Dim exportReader As SqlDataReader = Nothing
        Dim exportArray As Array = Nothing
        Dim htFilterParams As Hashtable = New Hashtable()


        Select Case exportFilename
            Case "CommentsExport.xlsx"
                Try
                    exportArray = getArrayData(HttpContext.Current.Session("ClaimCommentExport_SelectColumns"), CommandType.Text, "ClaimCommentExport_SearchParams", "ClaimCommentExport_WhereClause")
                Catch ex As Exception
                    MsgBox("Export '" + exportFilename + "' has failed.", , "Export Error")
                    Throw ex
                Finally
                    If Not exportReader Is Nothing AndAlso exportReader.IsClosed Then
                        exportReader.Close()
                    End If
                End Try
                Dim newExportFilename As String = String.Format("CommentsExport-{0}", DateTime.Now.ToString("yyyyMMddHHmmss")) & ".xlsx"
                exportConfigResponse(context, newExportFilename, aColumnModel, exportArray, "Comments")

            Case "ClaimAssessExport.xlsx"
                Try
                    exportArray = getArrayData(HttpContext.Current.Session("ClaimRoomExport_SelectColumns"), CommandType.Text, "ClaimRoomExport_SearchParams", "ClaimRoomExport_WhereClause")
                Catch ex As Exception
                    errString = ex.Message
                    errLocation = "Export " & " - " & exportFilename
                    CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), pageName, context.Request.UserHostAddress())
                Finally
                    If Not exportReader Is Nothing AndAlso exportReader.IsClosed Then
                        exportReader.Close()
                    End If
                End Try
                Dim newExportFilename As String = String.Format("ClaimAssessExport-{0}", DateTime.Now.ToString("yyyyMMddHHmmss")) & ".xlsx"
                exportConfigResponse(context, newExportFilename, aColumnModel, exportArray, "Claim Assessment")

            Case "ClaimScheduleExport.xlsx"
                Try
                    exportArray = getArrayData(HttpContext.Current.Session("ScheduleInvoiceExport_SelectColumns"), CommandType.Text)
                Catch ex As Exception
                    errString = ex.Message
                    errLocation = "Export " & " - " & exportFilename
                    CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), pageName, context.Request.UserHostAddress())
                Finally
                    If Not exportReader Is Nothing AndAlso exportReader.IsClosed Then
                        exportReader.Close()
                    End If
                End Try
                Dim newExportFilename As String = String.Format("ClaimScheduleExport-{0}", DateTime.Now.ToString("yyyyMMddHHmmss")) & ".xlsx"
                exportConfigResponse(context, newExportFilename, aColumnModel, exportArray, "Schedule Invoices")

            Case "ClaimsRoom.xlsx"
                Try
                    exportArray = getArrayData(HttpContext.Current.Session("ClaimsRoom_SelectColumns"), CommandType.Text, "ClaimsRoom_SearchParams", "ClaimsRoom_WhereClause")
                Catch ex As Exception
                    errString = ex.Message
                    errLocation = "Export " & " - " & exportFilename
                    CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), pageName, context.Request.UserHostAddress())
                Finally
                    If Not exportReader Is Nothing AndAlso exportReader.IsClosed Then
                        exportReader.Close()
                    End If
                End Try
                Dim newExportFilename As String = String.Format("ClaimsRoom-{0}", DateTime.Now.ToString("yyyyMMddHHmmss")) & ".xlsx"
                exportConfigResponse(context, newExportFilename, aColumnModel, exportArray, "Claims Room")
            Case "ADRLog.xlsx"
                Try
                    exportArray = getArrayData(HttpContext.Current.Session("ADRCommentExport_SelectColumns"), CommandType.Text, "ADRCommentExport_SearchParams", "ADRCommentExport_WhereClause")
                Catch ex As Exception
                    errString = ex.Message
                    errLocation = "Export " & " - " & exportFilename
                    CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), pageName, context.Request.UserHostAddress())
                Finally
                    If Not exportReader Is Nothing AndAlso exportReader.IsClosed Then
                        exportReader.Close()
                    End If
                End Try
                Dim newExportFilename As String = String.Format("ClaimsRoom-{0}", DateTime.Now.ToString("yyyyMMddHHmmss")) & ".xlsx"
                exportConfigResponse(context, newExportFilename, aColumnModel, exportArray, "ADR Log")
            Case "ContractsExport.xlsx"
                Try
                    exportArray = getArrayData(HttpContext.Current.Session("ContractsExport_SelectColumns"), CommandType.Text, "ContractsExport_SearchParams", "ContractsExport_WhereClause")
                Catch ex As Exception
                    errString = ex.Message
                    errLocation = "Export " & " - " & exportFilename
                    CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), pageName, context.Request.UserHostAddress())
                Finally
                    If Not exportReader Is Nothing AndAlso exportReader.IsClosed Then
                        exportReader.Close()
                    End If
                End Try
                Dim newExportFilename As String = String.Format("ContractsExport-{0}", DateTime.Now.ToString("yyyyMMddHHmmss")) & ".xlsx"
                exportConfigResponse(context, newExportFilename, aColumnModel, exportArray, "Claims Room")

            Case "ContractsExport_2.xlsx"
                Try
                    exportArray = getArrayData(HttpContext.Current.Session("ContractsExport_SelectColumns"), CommandType.Text, "ContractsExport_SearchParams", "ContractsExport_WhereClause")
                Catch ex As Exception
                    errString = ex.Message
                    errLocation = "Export " & " - " & exportFilename
                    CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), pageName, context.Request.UserHostAddress())
                Finally
                    If Not exportReader Is Nothing AndAlso exportReader.IsClosed Then
                        exportReader.Close()
                    End If
                End Try
                Dim newExportFilename As String = String.Format("ContractsExport-{0}", DateTime.Now.ToString("yyyyMMddHHmmss")) & ".xlsx"
                exportConfigResponse(context, newExportFilename, aColumnModel, exportArray, "Claims Room")
            Case "ADROffers.xlsx"
                Dim strSQL As String = " SELECT * FROM [vw_w_ClaimADROffers] "
                strSQL += context.Session("machineListHardwareCurrentViewWhereClause")

                Try
                    exportReader = getTableData(strSQL, CommandType.Text)

                    aColumnModel = GetColumnModel(exportReader)
                    exportArray = getArrayData(exportReader)
                    exportConfigResponse(context, exportFilename, aColumnModel, exportArray, "ADR Offer Rounds")
                Catch ex As Exception
                    MsgBox("Export '" + exportFilename + "' has failed.", , "Export Error")
                    Throw ex
                Finally
                    If Not exportReader Is Nothing AndAlso exportReader.IsClosed Then
                        exportReader.Close()
                    End If
                End Try
        End Select
    End Sub

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

    Private Function GetColumnModel(ByVal aReader As SqlDataReader) As Array
        Dim colCount As Integer = 0
        Dim isNumeric As Boolean = False
        Dim strHeader As String = ""
        Dim countLim As Integer = aReader.FieldCount
        ReDim aColumnModel(countLim - 1)

        While (colCount < countLim)
            aColumnModel(colCount) = New ColumnModel

            If (String.Compare(aReader.GetName(colCount), strHeader, True) <> 0) Then
                strHeader = aReader.GetName(colCount)
                aColumnModel(colCount).Header = strHeader
            End If

            Select Case aReader.GetDataTypeName(colCount)
                Case "char", "varchar", "text", "nchar", "nvarchar", "ntext"
                    aColumnModel(colCount).Type = DataType.String
                    isNumeric = False
                Case "bigint"
                    aColumnModel(colCount).Type = DataType.Long
                    isNumeric = True
                Case "int", "smallint"
                    aColumnModel(colCount).Type = DataType.Integer
                    isNumeric = True
                Case "decimal", "numeric", "money", "smallmoney", "float", "real"
                    aColumnModel(colCount).Type = DataType.Double
                    isNumeric = True
                Case "datetime", "smalldatetime", "date", "time"
                    aColumnModel(colCount).Type = DataType.Date
                    isNumeric = False
                Case "bit"
                    aColumnModel(colCount).Type = DataType.Boolean
                    isNumeric = False
            End Select

            If (isNumeric = True) Then
                aColumnModel(colCount).Alignment = HorizontalAlignment.Right
            Else
                aColumnModel(colCount).Alignment = HorizontalAlignment.Left
            End If

            colCount += 1
        End While
        Return aColumnModel
    End Function

    Private Sub exportConfigResponse(ByVal context As HttpContext, ByVal strFileName As String, ByVal aClnModel As Array, ByVal exportArray As Array, ByVal strName As String)
        Dim memstream As MemoryStream = New MemoryStream
        Try
            ExportToExcel.ExportToExcel.FillSpreadsheetDocument(memstream, aClnModel, exportArray, strName)
            '---- Configure response
            ' context.Response.AddHeader("Content-disposition", "attachment; filename=" & strFileName)
            context.Response.AddHeader("Content-disposition", "attachment; filename=" & ControlChars.Quote & strFileName & ControlChars.Quote)
            context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            context.Response.ContentEncoding = Encoding.UTF8
            memstream.WriteTo(context.Response.OutputStream)
            context.Response.Flush()
        Catch ex As Exception
            errString = ex.Message
            errLocation = "exportConfigResponse " & " - " & strFileName
            CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), pageName, context.Request.UserHostAddress())
        Finally

        End Try
    End Sub

    Private Function getTableData(ByVal strSQL As String, ByVal cmdType As CommandType, Optional ByVal params As SqlParameter() = Nothing) As SqlDataReader

        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim myreader As SqlDataReader = Nothing
        Try
            If Not params Is Nothing Then
                myreader = CommonUtilsv2.GetDataReader(dbKey, strSQL, cmdType, params)
            Else
                myreader = CommonUtilsv2.GetDataReader(dbKey, strSQL, cmdType)
            End If
            Return myreader
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#Region "Export methods"
    Public Function getTableData(ByVal strSQL As String, Optional ByVal params As SqlParameter() = Nothing) As Array
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey_Advanced")

        Dim data As Array = Nothing
        Dim alData As New ArrayList
        Dim iRowID As Integer = 0
        Dim exportReader As SqlDataReader = Nothing
        Dim errString As String = Nothing
        Dim errLocation As String = Nothing
        Try
            If Not params Is Nothing Then
                exportReader = CommonUtilsv2.GetDataReader(dbKey, strSQL, CommandType.Text, params)
            Else
                exportReader = CommonUtilsv2.GetDataReader(dbKey, strSQL, CommandType.Text)
            End If

            While exportReader.Read
                Dim saData(exportReader.FieldCount - 1) As String
                For i = 0 To exportReader.FieldCount - 1
                    saData(i) = If(IsDBNull(exportReader(i)), "", exportReader(i).ToString.Trim())
                Next
                alData.Add(saData)
            End While
        Catch ex As Exception
            errString = ex.Message
            errLocation = "getTableData()"
            CommonUtilsv2.CreateErrorLog(errString, errLocation, HttpContext.Current.Session("User"), System.IO.Path.GetFileName(HttpContext.Current.Request.RawUrl.ToString), HttpContext.Current.Request.UserHostAddress())
        Finally
        End Try
        data = DirectCast(alData.ToArray(GetType(String())), String()())
        exportReader.Close()
        Return data
    End Function

    Private Function getReader(ByVal strSQL As String, ByVal cmdType As CommandType, Optional ByVal params As SqlParameter() = Nothing) As SqlDataReader

        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim exportReader As SqlDataReader = Nothing
        Try
            If Not params Is Nothing Then
                exportReader = CommonUtilsv2.GetDataReader(dbKey, strSQL, cmdType, params)
            Else
                exportReader = CommonUtilsv2.GetDataReader(dbKey, strSQL, cmdType)
            End If
            Return exportReader
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function getArrayData(ByVal aReader As SqlDataReader) As Array
        If aReader Is Nothing OrElse Not aReader.HasRows() Then
            Dim emptyData As Array
            Dim alEmptyData As New ArrayList
            emptyData = DirectCast(alEmptyData.ToArray(GetType(String())), String()())
            Return emptyData
        End If

        Dim data As Array
        Dim alData As New ArrayList
        While aReader.Read
            Dim saData(aReader.FieldCount - 1) As String
            For i = 0 To aReader.FieldCount - 1
                saData(i) = If(IsDBNull(aReader(i)), "", aReader(i).ToString.Trim())
            Next
            alData.Add(saData)
        End While
        data = DirectCast(alData.ToArray(GetType(String())), String()())
        Return data
    End Function

    Private Function getArrayData(ByVal strSQL As String, ByVal cmdType As CommandType, Optional ByVal strSessionIDForParamHT As String = Nothing, Optional ByVal strSessionIDForWhereClause As String = Nothing) As Array
        aColumnModel = Nothing
        Dim dbKey As String = System.Configuration.ConfigurationManager.AppSettings("dbKey")
        Dim aReader As SqlDataReader = Nothing
        Dim strSQLSelect As String = strSQL
        Dim STR_WhereClause As String = Nothing
        Dim HT_SearchParams As Hashtable = Nothing
        If Not strSessionIDForParamHT Is Nothing Then
            HT_SearchParams = DirectCast(HttpContext.Current.Session(strSessionIDForParamHT), Hashtable)
        End If

        If Not strSessionIDForWhereClause Is Nothing Then
            STR_WhereClause = HttpContext.Current.Session(strSessionIDForWhereClause)
        End If
        If Not String.IsNullOrEmpty(STR_WhereClause) Then
            strSQLSelect = strSQLSelect & STR_WhereClause
        End If

        Dim myConn As New SqlClient.SqlConnection(dbKey)
        Dim myCmd As New SqlCommand(strSQLSelect, myConn)
        If Not String.IsNullOrEmpty(STR_WhereClause) Then
            Dim param As DictionaryEntry
            For Each param In HT_SearchParams
                myCmd.Parameters.AddWithValue(param.Key, param.Value)
            Next
        End If
        myCmd.CommandType = cmdType
        Dim data As Array = Nothing
        Try
            myConn.Open()
            aReader = myCmd.ExecuteReader()
            aColumnModel = GetColumnModel(aReader)

            If aReader Is Nothing OrElse Not aReader.HasRows() Then
                Dim emptyData As Array
                Dim alEmptyData As New ArrayList
                emptyData = DirectCast(alEmptyData.ToArray(GetType(String())), String()())
                Return emptyData
            End If
            Dim alData As New ArrayList
            While aReader.Read
                Dim saData(aReader.FieldCount - 1) As String
                For i = 0 To aReader.FieldCount - 1
                    saData(i) = If(IsDBNull(aReader(i)), "", aReader(i).ToString.Trim())
                Next
                alData.Add(saData)
            End While
            data = DirectCast(alData.ToArray(GetType(String())), String()())

        Catch ex As Exception
            CommonUtilsv2.CreateErrorLog(ex.Message, "getArrayData()", HttpContext.Current.Session("User"), "H_CreateDownload.ashx", HttpContext.Current.Request.UserHostAddress())
        Finally
            If Not aReader Is Nothing AndAlso aReader.IsClosed() Then
                aReader.Close()
            End If
        End Try
        Return data

    End Function
#End Region

End Class

