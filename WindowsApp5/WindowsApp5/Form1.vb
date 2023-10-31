Imports System.IO
Imports ExcelDataReader
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop



Public Class Form1

    Dim tables As DataTableCollection
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Using ofd As OpenFileDialog = New OpenFileDialog() With {.Filter = "Excel Workbook|*.xlsx"}
            If ofd.ShowDialog() = DialogResult.OK Then
                TextBox1.Text = ofd.FileName
                Using stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read)
                    Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                        Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                 .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                 .UseHeaderRow = True}})
                        tables = result.Tables
                        ComboBox1.Items.Clear()
                        For Each table As System.Data.DataTable In tables
                            ComboBox1.Items.Add(table.TableName)
                        Next
                    End Using

                End Using
            End If
        End Using


        ' Create OpenFileDialog object
        'Using ofd As OpenFileDialog = New OpenFileDialog() With {.Filter = "Excel Workbook|*.xlsx"}
        '    If ofd.ShowDialog() = DialogResult.OK Then
        '        TextBox1.Text = ofd.FileName
        '        Using stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read)
        '            Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
        '                Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
        '                                                     .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
        '                                                     .UseHeaderRow = True}})
        '                ComboBox1.Items.Clear()
        '                For Each table As System.Data.DataTable In result.Tables
        '                    ComboBox1.Items.Add(table.TableName)
        '                Next
        '            End Using
        '        End Using
        '    End If
        'End Using
    End Sub

    'Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
    '    ' Get the selected table name
    '    Dim selectedTable As String = ComboBox1.SelectedItem.ToString()
    '    ' Find the selected table in the DataSet object
    '    Dim selectedDataTable As DataTable = tables(selectedTable)
    '    ' Bind the data to the DataGridView
    '    DataGridView1.DataSource = selectedDataTable
    'End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim dt As System.Data.DataTable = tables(ComboBox1.SelectedItem.ToString())
        DataGridView1.DataSource = dt
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ''----

        Dim excelApp As New Excel.Application

        ' Open the Excel file
        Dim workbook As Excel.Workbook = excelApp.Workbooks.Open("C:\Users\parkk\OneDrive\Рабочий стол\book7.xlsx")

        ' Get the worksheets in the workbook
        Dim firstTableWorksheet As Excel.Worksheet = workbook.Sheets(1)
        Dim secondTableWorksheet As Excel.Worksheet = workbook.Sheets(2)
        Dim thirdTableWorksheet As Excel.Worksheet = workbook.Sheets(3)
        Dim fourthTableWorksheet As Excel.Worksheet = workbook.Sheets(4)

        ' Loop through the rows in the first table worksheet
        Dim firstTableRows As Excel.Range = firstTableWorksheet.Range("A2:F" & firstTableWorksheet.UsedRange.Rows.Count)
        For Each firstTableRow As Excel.Range In firstTableRows.Rows
            Dim deviceCode As String = firstTableRow.Cells(2).Value ' Get the value of the deviceCode column
            Dim operDataTime As Date = firstTableRow.Cells(3).Value ' Get the value of the operDataTime column
            Dim currency As String = firstTableRow.Cells(4).Value ' Get the value of the currency column
            Dim amountToAdd As Double = firstTableRow.Cells(5).Value ' Get the value of the amount column
            Dim cardNumber As String = firstTableRow.Cells(6).Value ' Get the value of the cardNumber column

            ' Find the row in the second table worksheet where the currency matches
            Dim secondTableCurrencyColumn As Excel.Range = secondTableWorksheet.Range("A:A")
            Dim secondTableCurrencyCell As Excel.Range = secondTableCurrencyColumn.Find(currency)



            ' Find the row in the third table worksheet where the card number matches
            Dim thirdTableCardNumberColumn As Excel.Range = thirdTableWorksheet.Range("A:A")
            Dim thirdTableCardNumberCell As Excel.Range = thirdTableCardNumberColumn.Find(cardNumber)


            Dim thirdTableAmountCell As Excel.Range = thirdTableWorksheet.Cells(thirdTableCardNumberCell.Row, 2) ' Assumes "amount" is in the second column (column B)
            Dim thirdTableAmount As Double = thirdTableAmountCell.Value ' Get the current value of the "amount" cell

            If amountToAdd <= thirdTableAmount Then

                ' If the secondTableCurrencyCell is not null, add the amount to the existing value in the "amount" cell
                If secondTableCurrencyCell IsNot Nothing Then
                    Dim secondTableAmountCell As Excel.Range = secondTableWorksheet.Cells(secondTableCurrencyCell.Row, 2) ' Assumes "amount" is in the second column (column B)
                    Dim currentAmount As Double = secondTableAmountCell.Value ' Get the current value of the "amount" cell
                    secondTableAmountCell.Value = currentAmount + amountToAdd ' Add the amount to the current value
                End If


                ' If the thirdTableCardNumberCell is not null, subtract the amount from the existing value in the "amount" cell
                If thirdTableCardNumberCell IsNot Nothing Then
                    Dim thirdTableAmountCell2 As Excel.Range = thirdTableWorksheet.Cells(thirdTableCardNumberCell.Row, 2) ' Assumes "amount" is in the second column (column B)
                    Dim currentAmount2 As Double = thirdTableAmountCell2.Value ' Get the current value of the "amount" cell
                    thirdTableAmountCell.Value = currentAmount2 - amountToAdd ' Subtract the amount from the current value
                    'firstTableRow.Copy(fourthTableWorksheet.Rows(fourthTableWorksheet.UsedRange.Rows.Count + 1))
                    'firstTableRow.Delete()
                End If

                firstTableRow.Copy(fourthTableWorksheet.Rows(fourthTableWorksheet.UsedRange.Rows.Count + 1))
                firstTableRow.Delete()

            End If
        Next
        firstTableRows = firstTableWorksheet.Range("A2:F" & firstTableWorksheet.UsedRange.Rows.Count)

        For Each firstTableRow As Excel.Range In firstTableRows.Rows
            Dim deviceCode As String = firstTableRow.Cells(2).Value ' Get the value of the deviceCode column
            Dim operDataTime As Date = firstTableRow.Cells(3).Value ' Get the value of the operDataTime column
            Dim currency As String = firstTableRow.Cells(4).Value ' Get the value of the currency column
            Dim amountToAdd As Double = firstTableRow.Cells(5).Value ' Get the value of the amount column
            Dim cardNumber As String = firstTableRow.Cells(6).Value ' Get the value of the cardNumber column

            ' Find the row in the second table worksheet where the currency matches
            Dim secondTableCurrencyColumn As Excel.Range = secondTableWorksheet.Range("A:A")
            Dim secondTableCurrencyCell As Excel.Range = secondTableCurrencyColumn.Find(currency)



            ' Find the row in the third table worksheet where the card number matches
            Dim thirdTableCardNumberColumn As Excel.Range = thirdTableWorksheet.Range("A:A")
            Dim thirdTableCardNumberCell As Excel.Range = thirdTableCardNumberColumn.Find(cardNumber)


            Dim thirdTableAmountCell As Excel.Range = thirdTableWorksheet.Cells(thirdTableCardNumberCell.Row, 2) ' Assumes "amount" is in the second column (column B)
            Dim thirdTableAmount As Double = thirdTableAmountCell.Value ' Get the current value of the "amount" cell

            If amountToAdd <= thirdTableAmount Then

                firstTableRow.Copy(fourthTableWorksheet.Rows(fourthTableWorksheet.UsedRange.Rows.Count + 1))
                firstTableRow.Delete()

            End If
        Next
        ' Save the changes
        workbook.Save()

        ' Close the workbook and Excel application
        workbook.Close()
        excelApp.Quit()

        ' Release the COM objects
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(thirdTableCardNumberCell)
        'System.Runtime.InteropServices
        ''-----------------------

        ''move()


    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        'Open the source Excel file
        Dim excelApp As New Excel.Application
        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Open("C:\Users\parkk\OneDrive\Рабочий стол\book7.xlsx")

        'Get the fourth sheet
        Dim sourceSheet As Excel.Worksheet = excelWorkbook.Sheets(4)

        'Create a new Excel file
        Dim newExcelWorkbook As Excel.Workbook = excelApp.Workbooks.Add()

        'Get the first sheet of the new Excel file
        Dim newSheet As Excel.Worksheet = newExcelWorkbook.Sheets(1)

        'Copy the fourth sheet to the first sheet of the new Excel file
        sourceSheet.Copy(newSheet)

        'Save and close the new Excel file
        newExcelWorkbook.SaveAs("C:\Users\parkk\OneDrive\Рабочий стол\report.xlsx")
        newExcelWorkbook.Close()

        'Close the source Excel file
        excelWorkbook.Close()
        excelApp.Quit()





        'Dim excelApp As New Application()

        '' Open the source Excel file
        'Dim sourceWorkbook As Workbook = excelApp.Workbooks.Open("C:\Users\parkk\OneDrive\Рабочий стол\book6.xlsx")

        '' Get the fourth sheet of the source file
        'Dim sourceSheet As Worksheet = sourceWorkbook.Worksheets(4)

        '' Create a new workbook for the destination file
        'Dim destWorkbook As Workbook = excelApp.Workbooks.Add()

        '' Copy the fourth sheet to the destination workbook
        'sourceSheet.Copy(, destWorkbook.Worksheets(1))

        '' Save the destination workbook
        'destWorkbook.SaveAs("C:\Users\parkk\OneDrive\Рабочий стол\report.xlsx")

        '' Close both workbooks and quit Excel
        'sourceWorkbook.Close()
        'destWorkbook.Close()
        'excelApp.Quit()

        '' Release COM objects from memory
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceSheet)
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceWorkbook)
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(destWorkbook)
        'System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)

        '------------------------------------------------------------------

        'Dim excelApp As New Excel.Application

        '' Open the Excel file
        'Dim workbook As Excel.Workbook = excelApp.Workbooks.Open("C:\Users\parkk\OneDrive\Рабочий стол\book3.xlsx")

        '' Get the third worksheet in the workbook
        'Dim worksheet As Excel.Worksheet = workbook.Sheets(1)

        '' Get the used range of the worksheet
        'Dim usedRange As Excel.Range = worksheet.UsedRange

        '' Get the number of rows and columns in the used range
        'Dim rowCount As Integer = usedRange.Rows.Count
        'Dim columnCount As Integer = usedRange.Columns.Count

        '' Create a new PDF document
        'Dim document As New Document

        '' Create a new PdfWriter to write the document to a file
        'Dim writer As PdfWriter = PdfWriter.GetInstance(document, New FileStream("C:\Users\parkk\OneDrive\Рабочий стол\bank.pdf", FileMode.Create))

        '' Open the document
        'document.Open()

        '' Create a new PdfPTable with the same number of columns as the Excel sheet
        'Dim pdfTable As New PdfPTable(columnCount)

        '' Loop through the rows and columns in the used range
        'For row As Integer = 1 To rowCount
        '    For column As Integer = 1 To columnCount
        '        ' Get the value of the current cell
        '        Dim cellValue As Object = usedRange.Cells(row, column).Value

        '        ' If the cell value is not null, add it to the PDF table
        '        If cellValue IsNot Nothing Then
        '            pdfTable.AddCell(cellValue.ToString)
        '        End If
        '    Next
        'Next

        '' Add the PDF table to the document
        'document.Add(pdfTable)

        '' Close the document and release the resources
        'document.Close()
        'writer.Close()
        'workbook.Close(False)
        'excelApp.Quit()
        ''ReleaseObject(usedRange)
        ''ReleaseObject(worksheet)
        ''ReleaseObject(workbook)
        ''ReleaseObject(excelApp)

        '' Display a message box indicating that the PDF report has been created
        'MessageBox.Show("PDF report created successfully!")

    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Class
