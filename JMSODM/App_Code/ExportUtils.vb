Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.IO
Imports System.Globalization
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExportToExcel

    Public Enum DataType
        [Integer]
        [Long]
        [Double]
        [String]
        [Date]
        [Boolean]
    End Enum

    Public Enum HorizontalAlignment
        [Left]
        [Center]
        [Right]
    End Enum

    Public Enum OutputCellFormat As UInteger
        [Text]
        [Integer]
        [Date]
        [TextHeader]
        [TextHeaderRotated]
        [TextCenter]
        [TextRight]
    End Enum

    Public Class ColumnModel
        Public Property Type() As DataType
            Get
                Return m_Type
            End Get
            Set(value As DataType)
                m_Type = value
            End Set
        End Property
        Private m_Type As DataType
        Public Property Alignment() As HorizontalAlignment
            Get
                Return m_Alignment
            End Get
            Set(value As HorizontalAlignment)
                m_Alignment = value
            End Set
        End Property
        Private m_Alignment As HorizontalAlignment
        Public Property Header() As String
            Get
                Return m_Header
            End Get
            Set(value As String)
                m_Header = value
            End Set
        End Property
        Private m_Header As String
        Public Property IsRotatedHeader() As Boolean
            Get
                Return m_IsRotatedHeader
            End Get
            Set(value As Boolean)
                m_IsRotatedHeader = value
            End Set
        End Property
        Private m_IsRotatedHeader As Boolean
    End Class

    Public Class ExportToExcel

        Private Shared Function ConvertIntToColumnHeader(iCol As UInteger) As StringBuilder
            Dim sb = New StringBuilder()
            Dim A As Integer = 0
            Dim Z As Integer = 25

            While iCol > 0
                If iCol <= Asc("Z"c) - Asc("A"c) Then
                    ' iCol=0 -> 'A', 25 -> 'Z'
                    Exit While
                End If
                sb.Append(ConvertIntToColumnHeader(iCol \ (Asc("Z"c) - Asc("A"c) + 1) - 1))
                iCol = iCol Mod (Asc("Z"c) - Asc("A"c) + 1)
            End While
            sb.Append(ChrW(Asc("A"c) + iCol))

            Return sb
        End Function

        Private Shared Function GetCellReference(ByVal iRow As UInteger, ByVal iCol As UInteger) As String
            Return ConvertIntToColumnHeader(iCol).Append(iRow).ToString()
        End Function

        Private Shared Function CreateColumnHeaderRow(iRow As UInteger, columnModels As IList(Of ColumnModel)) As Row
            Dim r = New Row() With { _
                 .RowIndex = iRow _
            }

            For iCol As Integer = 0 To columnModels.Count - 1
                Dim styleIndex = If(columnModels(iCol).IsRotatedHeader, OutputCellFormat.TextHeaderRotated + 1, OutputCellFormat.TextHeader + 1)
                ' create Cell with InlineString as a child, which has Text as a child
                r.Append(New OpenXmlElement() {New Cell(New InlineString(New Text() With { _
                     .Text = columnModels(iCol).Header _
                })) With { _
                     .DataType = CellValues.InlineString, _
                     .StyleIndex = styleIndex, _
                     .CellReference = GetCellReference(iRow, CUInt(iCol)) _
                }})
            Next

            Return r
        End Function

        Private Shared Function GetStyleIndexFromColumnModel(columnModel As ColumnModel) As UInt32Value
            Select Case columnModel.Type
                Case DataType.[Integer]
                    Return CUInt(OutputCellFormat.[Integer]) + 1
                Case DataType.[Date]
                    Return CUInt(OutputCellFormat.[Date]) + 1
            End Select

            Select Case columnModel.Alignment
                Case HorizontalAlignment.Center
                    Return CUInt(OutputCellFormat.TextCenter) + 1
                Case HorizontalAlignment.Right
                    Return CUInt(OutputCellFormat.TextRight) + 1
                Case Else
                    Return CUInt(OutputCellFormat.Text) + 1
            End Select
        End Function

        Private Shared Function ConvertDateToString([date] As String) As String
            Dim dt As DateTime
            Dim text As String = [date]
            ' default results of conversion
            If DateTime.TryParse([date], dt) Then
                text = dt.ToOADate().ToString(CultureInfo.InvariantCulture)
            End If
            Return text
        End Function

        Private Shared Function CreateRow(iRow As UInt32, data As IList(Of String), columnModels As IList(Of ColumnModel), sharedStrings As IDictionary(Of String, Integer)) As Row
            Dim r = New Row() With { _
                 .RowIndex = iRow _
            }
            For iCol As Integer = 0 To data.Count - 1
                Dim styleIndex = CUInt(OutputCellFormat.Text) + 1
                If columnModels IsNot Nothing AndAlso iCol < columnModels.Count Then
                    styleIndex = GetStyleIndexFromColumnModel(columnModels(iCol))
                    Select Case columnModels(iCol).Type
                        Case DataType.[Integer]
                            ' create Cell with CellValue as a child, which has Text as a child
                            r.Append(New OpenXmlElement() {New Cell(New CellValue() With { _
                                 .Text = data(iCol) _
                            }) With { _
                                 .StyleIndex = styleIndex, _
                                 .CellReference = GetCellReference(iRow, CUInt(iCol)) _
                            }})
                            Continue For
                        Case DataType.Double
                            ' create Cell with CellValue as a child, which has Text as a child
                            r.Append(New OpenXmlElement() {New Cell(New CellValue() With {
                                 .Text = data(iCol)
                            }) With {
                                 .StyleIndex = styleIndex,
                                 .CellReference = GetCellReference(iRow, CUInt(iCol))
                            }})
                            Continue For

                        Case DataType.[Date]
                            ' create Cell with CellValue as a child, which has Text as a child
                            r.Append(New OpenXmlElement() {New Cell(New CellValue() With { _
                                 .Text = ConvertDateToString(data(iCol)) _
                            }) With { _
                                 .StyleIndex = styleIndex, _
                                 .CellReference = GetCellReference(iRow, CUInt(iCol)) _
                            }})
                            Continue For
                    End Select
                End If

                ' default format is text
                If Not sharedStrings.ContainsKey(data(iCol)) Then
                    ' create Cell with InlineString as a child, which has Text as a child
                    r.Append(New OpenXmlElement() {New Cell(New InlineString(New Text() With { _
                         .Text = data(iCol) _
                    })) With { _
                         .DataType = CellValues.InlineString, _
                         .StyleIndex = styleIndex, _
                         .CellReference = GetCellReference(iRow, CUInt(iCol)) _
                    }})
                End If

                ' create Cell with CellValue as a child, which has Text as a child
                r.Append(New OpenXmlElement() {New Cell(New CellValue() With { _
                     .Text = sharedStrings(data(iCol)).ToString(CultureInfo.InvariantCulture) _
                }) With { _
                     .DataType = CellValues.SharedString, _
                     .StyleIndex = styleIndex, _
                     .CellReference = GetCellReference(iRow, CUInt(iCol)) _
                }})
            Next

            Return r
        End Function

        Private Shared Sub FillSpreadsheetDocument(spreadsheetDocument As SpreadsheetDocument, columnModels As IList(Of ColumnModel), data As String()(), sheetName As String)
            If columnModels Is Nothing Then
                Throw New ArgumentNullException("columnModels")
            End If
            If data Is Nothing Then
                Throw New ArgumentNullException("data")
            End If

            ' add empty workbook and worksheet to the SpreadsheetDocument
            Dim workbookPart = spreadsheetDocument.AddWorkbookPart()
            Dim worksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
            Dim workbookStylesPart = workbookPart.AddNewPart(Of WorkbookStylesPart)()

            ' create styles for the header and columns
            ' Index 0 - The default font.
            ' Index 1 - The bold font.
            ' Index 0 - required, reserved by Excel - no pattern
            ' Index 1 - required, reserved by Excel - fill of gray 125
            ' Index 2 - no pattern text on gray background
            ' Index 0 - The default border.
            ' Index 1 - Applies a Left, Right, Top, Bottom border to a cell
            ' Index 0 - The default cell style.  If a cell does not have a style iCol applied it will use this style combination instead
            ' Index 1 - Alignment Left, Text
            ' "@" - text format - see http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat.aspx
            ' Index 2 - Interger Number
            ' "0" - integer format - see http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat.aspx
            ' Index 3 - Interger Date
            ' "14" - date format mm-dd-yy - see http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat.aspx
            ' Index 4 - Text for headers
            ' "@" - text format - see http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat.aspx
            ' Index 5 - Text for headers rotated
            ' "@" - text format - see http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat.aspx
            ' Index 6 - Alignment Center, Text
            ' "@" - text format - see http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat.aspx
            ' Index 7 - Alignment Right, Text
            ' "@" - text format - see http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat.aspx
            workbookStylesPart.Stylesheet = New Stylesheet(New Fonts(New Font(New FontSize() With { _
                 .Val = 11 _
            }, New Color() With { _
                 .Rgb = New HexBinaryValue() With { _
                     .Value = "00000000" _
                } _
            }, New FontName() With { _
                 .Val = "Calibri" _
            }), New Font(New Bold(), New FontSize() With { _
                 .Val = 11 _
            }, New Color() With { _
                 .Rgb = New HexBinaryValue() With { _
                     .Value = "00000000" _
                } _
            }, New FontName() With { _
                 .Val = "Calibri" _
            })), New Fills(New Fill(New PatternFill() With { _
                 .PatternType = PatternValues.None _
            }), New Fill(New PatternFill() With { _
                 .PatternType = PatternValues.Gray125 _
            }), New Fill(New PatternFill() With { _
                 .PatternType = PatternValues.Solid, _
                 .BackgroundColor = New BackgroundColor() With { _
                     .Indexed = 64UI _
                }, _
                 .ForegroundColor = New ForegroundColor() With { _
                     .Rgb = "FFD9D9D9" _
                } _
            })), New Borders(New Border(New LeftBorder(), New RightBorder(), New TopBorder(), New BottomBorder(), New DiagonalBorder()), New Border(New LeftBorder(New Color() With { _
                 .Auto = True _
            }) With { _
                 .Style = BorderStyleValues.Thin _
            }, New RightBorder(New Color() With { _
                 .Auto = True _
            }) With { _
                 .Style = BorderStyleValues.Thin _
            }, New TopBorder(New Color() With { _
                 .Auto = True _
            }) With { _
                 .Style = BorderStyleValues.Thin _
            }, New BottomBorder(New Color() With { _
                 .Auto = True _
            }) With { _
                 .Style = BorderStyleValues.Thin _
            }, New DiagonalBorder())), New CellFormats(New CellFormat() With { _
                 .NumberFormatId = 0UI, _
                 .FontId = 0UI, _
                 .FillId = 0UI, _
                 .BorderId = 0UI _
            }, New CellFormat(New Alignment() With { _
                 .Horizontal = HorizontalAlignmentValues.Left _
            }) With { _
                 .NumberFormatId = 49UI, _
                 .FontId = 0UI, _
                 .FillId = 0UI, _
                 .BorderId = 1UI, _
                 .ApplyNumberFormat = True, _
                 .ApplyAlignment = True _
            }, New CellFormat() With { _
                 .NumberFormatId = 1UI, _
                 .FontId = 0UI, _
                 .FillId = 0UI, _
                 .BorderId = 1UI, _
                 .ApplyNumberFormat = True _
            }, New CellFormat() With { _
                 .NumberFormatId = 14UI, _
                 .FontId = 0UI, _
                 .FillId = 0UI, _
                 .BorderId = 1UI, _
                 .ApplyNumberFormat = True _
            }, New CellFormat(New Alignment() With { _
                 .Vertical = VerticalAlignmentValues.Center, _
                 .Horizontal = HorizontalAlignmentValues.Center _
            }) With { _
                 .NumberFormatId = 49UI, _
                 .FontId = 1UI, _
                 .FillId = 2UI, _
                 .BorderId = 1UI, _
                 .ApplyNumberFormat = True, _
                 .ApplyAlignment = True _
            }, New CellFormat(New Alignment() With { _
                 .Horizontal = HorizontalAlignmentValues.Center, _
                 .TextRotation = 90UI _
            }) With { _
                 .NumberFormatId = 49UI, _
                 .FontId = 1UI, _
                 .FillId = 2UI, _
                 .BorderId = 1UI, _
                 .ApplyNumberFormat = True, _
                 .ApplyAlignment = True _
            }, _
                New CellFormat(New Alignment() With { _
                 .Horizontal = HorizontalAlignmentValues.Center _
            }) With { _
                 .NumberFormatId = 49UI, _
                 .FontId = 0UI, _
                 .FillId = 0UI, _
                 .BorderId = 1UI, _
                 .ApplyNumberFormat = True, _
                 .ApplyAlignment = True _
            }, New CellFormat(New Alignment() With { _
                 .Horizontal = HorizontalAlignmentValues.Right _
            }) With { _
                 .NumberFormatId = 49UI, _
                 .FontId = 0UI, _
                 .FillId = 0UI, _
                 .BorderId = 1UI, _
                 .ApplyNumberFormat = True, _
                 .ApplyAlignment = True _
            }))
            workbookStylesPart.Stylesheet.Save()

            ' create and fill SheetData
            Dim sheetData = New SheetData()

            ' first row is the header
            Dim iRow As UInteger = 1
            sheetData.AppendChild(CreateColumnHeaderRow(iRow, columnModels))

            ' first of all collect all different strings
            Dim sst = New SharedStringTable()
            Dim sharedStrings = New SortedDictionary(Of String, Integer)()
            For Each dataRow As String() In data
                For iCol As Integer = 0 To dataRow.Length - 1
                    If iCol >= columnModels.Count OrElse columnModels(iCol).Type <> DataType.[Integer] Then
                        Dim text As String = dataRow(iCol)
                        If columnModels(iCol).Type = DataType.Date Then
                            text = ConvertDateToString(text)
                        End If
                        'Dim text As String = If(columnModels(iCol).Type = DataType.[Date], dataRow(iCol), ConvertDateToString(dataRow(iCol)))
                        If Not sharedStrings.ContainsKey(Text) Then
                            sst.AppendChild(New SharedStringItem(New Text(Text)))
                            sharedStrings.Add(Text, sharedStrings.Count)
                        End If
                    End If
                Next
            Next

            Dim shareStringPart = workbookPart.AddNewPart(Of SharedStringTablePart)()
            shareStringPart.SharedStringTable = sst

            shareStringPart.SharedStringTable.Save()

            For Each dataRow As String() In data
                iRow += 1
                sheetData.AppendChild(CreateRow(iRow, dataRow, columnModels, sharedStrings))
            Next

            ' add sheet data to Worksheet
            worksheetPart.Worksheet = New Worksheet(sheetData)
            worksheetPart.Worksheet.Save()

            ' fill workbook with the Worksheet
            ' generate the id for sheet
            spreadsheetDocument.WorkbookPart.Workbook = New Workbook(New FileVersion() With { _
                 .ApplicationName = "Microsoft Office Excel" _
            }, New Sheets(New Sheet() With { _
                 .Name = sheetName, _
                 .SheetId = 1UI, _
                 .Id = workbookPart.GetIdOfPart(worksheetPart) _
            }))
            spreadsheetDocument.WorkbookPart.Workbook.Save()
            spreadsheetDocument.Close()
        End Sub

        Public Shared Sub FillSpreadsheetDocument(stream As Stream, columnModels As ColumnModel(), data As String()(), sheetName As String)
            Using spreadsheetDocument__1 = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)
                FillSpreadsheetDocument(spreadsheetDocument__1, columnModels, data, sheetName)
            End Using
        End Sub

    End Class

End Namespace
