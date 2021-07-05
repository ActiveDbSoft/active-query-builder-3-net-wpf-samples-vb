//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports System
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Namespace Common
    Public Module ExportHelpers
        Private _cellStyleDateTime As ICellStyle

        Public Function ExportToExcel(dataTable As DataTable, pathToSave As String) As Boolean
            Try
                If dataTable Is Nothing OrElse String.IsNullOrEmpty(pathToSave) Then
                    Return False
                End If

                Dim workbook = New XSSFWorkbook()
                Dim sheet = CType(workbook.CreateSheet("Result"), XSSFSheet)

                Dim row = sheet.CreateRow(0)

                For Each column As DataColumn In dataTable.Columns
                    Dim cell = row.CreateCell(dataTable.Columns.IndexOf(column))
                    cell.SetCellValue(column.ColumnName)
                Next column

                For i = 0 To dataTable.Rows.Count - 1
                    Dim tableRow = dataTable.Rows(i)

                    Dim currentRow = sheet.CreateRow(i + 1)

                    Dim values = tableRow.ItemArray.ToList()

                    Dim dateValue As DateTime
                    Dim doubleValue As Double
                    Dim intValue As Integer

                    For Each item In values
                        Dim currentCell = currentRow.CreateCell(values.IndexOf(item))

                        If Date.TryParse(item.ToString(), dateValue) Then
                            If _cellStyleDateTime Is Nothing Then
                                Dim format = workbook.CreateDataFormat()
                                _cellStyleDateTime = workbook.CreateCellStyle()
                                _cellStyleDateTime.DataFormat = format.GetFormat("dd/MM/yyyy")
                            End If

                            currentCell.SetCellValue(dateValue)
                            currentCell.CellStyle = _cellStyleDateTime
                        ElseIf Double.TryParse(item.ToString(), doubleValue) Then
                            currentCell.SetCellValue(doubleValue)
                        ElseIf Integer.TryParse(item.ToString(), intValue) Then
                            currentCell.SetCellValue(intValue)
                        ElseIf TypeOf item Is Byte() Then
                            InsertImage(workbook, sheet, DirectCast(item, Byte()), i + 1, values.IndexOf(item))
                        Else
                            currentCell.SetCellValue(item.ToString())
                        End If
                    Next item
                Next i

                Using fs = New FileStream(pathToSave, FileMode.Create, FileAccess.Write)
                    workbook.Write(fs)
                End Using

                Return True
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Export to Excel", MessageBoxButton.OK)
                Return False
            End Try
        End Function

        Private Sub InsertImage(workbook As XSSFWorkbook, sheet As ISheet, data() As Byte, row As Integer, column As Integer)
            Try
                Dim picInd As Integer = workbook.AddPicture(data, XSSFWorkbook.PICTURE_TYPE_PNG)
                Dim patriarch As IDrawing = sheet.CreateDrawingPatriarch()
                Dim anchor As New XSSFClientAnchor(0, 0, 0, 0, column, row, column + 1, row + 1) With {.AnchorType = AnchorType.MoveAndResize}

                Dim picture As XSSFPicture = CType(patriarch.CreatePicture(anchor, picInd), XSSFPicture)
                'Reset the image to the original size.
                'picture.Resize();   //Note: Resize will reset client anchor you set.

                picture.LineStyle = LineStyle.DashDotGel
            Catch
                'ignore
            End Try
        End Sub

        Public Function ExportToCSV(dataTable As DataTable, strFilePath As String) As Boolean
            If dataTable Is Nothing OrElse String.IsNullOrEmpty(strFilePath) Then
                Return False
            End If

            Try
                Dim sw As New StreamWriter(strFilePath, False, Encoding.UTF8)
                'headers  
                For i As Integer = 0 To dataTable.Columns.Count - 1
                    sw.Write(dataTable.Columns(i))
                    If i < dataTable.Columns.Count - 1 Then
                        sw.Write(";")
                    End If
                Next i

                sw.Write(sw.NewLine)
                For Each dr As DataRow In dataTable.Rows
                    For i As Integer = 0 To dataTable.Columns.Count - 1
                        If Not Convert.IsDBNull(dr(i)) Then
                            Dim value As String = dr(i).ToString()
                            If value.Contains(";"c) Then
                                value = $"""{value}"""
                                sw.Write(value)
                            Else
                                sw.Write(dr(i).ToString())
                            End If
                        End If

                        If i < dataTable.Columns.Count - 1 Then
                            sw.Write(";")
                        End If
                    Next i

                    sw.Write(sw.NewLine)
                Next dr

                sw.Close()
                Return True
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Export to CVS", MessageBoxButton.OK)
                Return False
            End Try
        End Function
    End Module
End Namespace
