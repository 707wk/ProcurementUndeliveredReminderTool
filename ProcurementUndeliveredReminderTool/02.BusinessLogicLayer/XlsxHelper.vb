Imports System.Globalization
Imports System.IO
Imports OfficeOpenXml
''' <summary>
''' xlsx辅助模块
''' </summary>
Public Class XlsxHelper

    Public Shared Sub SaveResultList(outputPath As String,
                                     docs As List(Of DocumentInfo))

        Using tmpExcelPackage As New ExcelPackage()
            Dim tmpWorkBook = tmpExcelPackage.Workbook

            Dim tmpWorkSheet = tmpWorkBook.Worksheets.Add($"{Now.Date:D}")

            ' 创建列标题
            Dim headItems = {
            "采购人员",
            "供应商简称",
            "采购单号",
            "品号",
            "品名",
            "规格",
            "单位",
            "采购数量",
            "已交数量",
            "未交数量",
            "预交货日",
            "备注"
            }
            For i001 = 1 To headItems.Count
                tmpWorkSheet.Cells(1, i001).Value = headItems(i001 - 1)
            Next

            ' 列标题筛选
            tmpWorkSheet.Cells($"1:1").AutoFilter = True

            ' 设置标题背景色
            tmpWorkSheet.Cells($"1:1").Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
            tmpWorkSheet.Cells($"1:1").Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.YellowGreen)

            For tmpIndex = 1 To docs.Count

                tmpWorkSheet.Cells(tmpIndex + 1, 1).Value = docs(tmpIndex - 1).CGRY
                tmpWorkSheet.Cells(tmpIndex + 1, 2).Value = docs(tmpIndex - 1).GYSJC
                tmpWorkSheet.Cells(tmpIndex + 1, 3).Value = docs(tmpIndex - 1).CGDH
                tmpWorkSheet.Cells(tmpIndex + 1, 4).Value = docs(tmpIndex - 1).PH
                tmpWorkSheet.Cells(tmpIndex + 1, 5).Value = docs(tmpIndex - 1).PM
                tmpWorkSheet.Cells(tmpIndex + 1, 6).Value = docs(tmpIndex - 1).GG
                tmpWorkSheet.Cells(tmpIndex + 1, 7).Value = docs(tmpIndex - 1).Unit
                tmpWorkSheet.Cells(tmpIndex + 1, 8).Value = docs(tmpIndex - 1).CGSL
                tmpWorkSheet.Cells(tmpIndex + 1, 9).Value = docs(tmpIndex - 1).YJSL
                tmpWorkSheet.Cells(tmpIndex + 1, 10).Value = docs(tmpIndex - 1).WJSL
                tmpWorkSheet.Cells(tmpIndex + 1, 11).Value = docs(tmpIndex - 1).YJHR
                tmpWorkSheet.Cells(tmpIndex + 1, 12).Value = docs(tmpIndex - 1).Remark

            Next

            ' 首行冻结
            tmpWorkSheet.View.FreezePanes(2, 1)

            ' 设置单元格值格式
            tmpWorkSheet.Cells($"A:G").Style.Numberformat.Format = "@"

            tmpWorkSheet.Cells($"H:J").Style.Numberformat.Format = "#,##0.00"

            tmpWorkSheet.Cells($"K:K").Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern

            ' 自动列宽
            tmpWorkSheet.Cells.AutoFitColumns()

            ' 自动行高
            For i001 = 1 To tmpWorkSheet.Dimension.End.Row
                tmpWorkSheet.Row(i001).CustomHeight = True
            Next

            Using tmpSaveFileStream = New FileStream(outputPath, FileMode.Create)
                tmpExcelPackage.SaveAs(tmpSaveFileStream)
            End Using
        End Using

    End Sub

End Class
