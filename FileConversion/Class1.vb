
Imports Syncfusion.DocIO
Imports Syncfusion.DocIO.DLS

Class Class1
    Shared Sub Main()
        Dim dotxFilePath As String = "C:\Users\RajeA\Documents\Custom Office Templates\Employee details.dotx"
        Dim docxFilePath As String = "C:\Users\RajeA\Documents\ConvertedFile.docx"
        Dim rowCount As Integer = 5
        Using document As New WordDocument(dotxFilePath)
            Dim bookmarkNavigator As New BookmarksNavigator(document)
            AddEmployeeDetails(bookmarkNavigator, "FirstName", "Atharva")
            AddEmployeeDetails(bookmarkNavigator, "lastName", "Raje")
            AddEmployeeDetails(bookmarkNavigator, "CompanyName", "AdtechCorp")
            AddEmployeeDetails(bookmarkNavigator, "Worklocation", "Hyderabad,Telangana,India")
            GenerateTable(bookmarkNavigator, "OrderTable", rowCount, AddressOf GenerateOrderRow)
            GenerateTable(bookmarkNavigator, "JobTable", rowCount, AddressOf GenerateJobRow)
            GenerateTable(bookmarkNavigator, "InventoryTable", rowCount, AddressOf GenerateInventoryRow)
            document.Save(docxFilePath, FormatType.Docx)
        End Using
        Console.WriteLine("Changed to docx,filled bookmarks and added rows to the tables.")
    End Sub
    Private Shared Sub AddEmployeeDetails(bookmarkNavigator As BookmarksNavigator, bookmarknName As String, inputValue As String)
        bookmarkNavigator.MoveToBookmark(bookmarknName)
        bookmarkNavigator.DeleteBookmarkContent(True)
        bookmarkNavigator.InsertText(inputValue)
    End Sub
    Private Shared Sub GenerateTable(bookmarkNavigator As BookmarksNavigator, bookmarkName As String, rowCount As Integer, rowGenerator As Action(Of Integer, WTableRow))
        bookmarkNavigator.MoveToBookmark(bookmarkName)
        'Dim table As WTable = CType(bookmarkNavigator.CurrentBookmark.BookmarkStart.Owner.Owner.Owner.Owner, WTable)
        Dim templateRow = CType(bookmarkNavigator.CurrentBookmark.BookmarkStart.OwnerParagraph.OwnerTextBody, WTableCell).OwnerRow
        Dim table = CType(templateRow.Owner, WTable)
        For i As Integer = 1 To rowCount
            Dim row As WTableRow
            If i = 1 Then
                row = table.Rows(i)
                'For Each cell As WTableCell In row.Cells
                '    cell.Paragraphs.RemoveAt(0)
                'Next
                row.Cells.Cast(Of WTableCell)().ToList().ForEach(Sub(cell) cell.Paragraphs.RemoveAt(0))
            Else
                row = table.AddRow()
            End If
            rowGenerator(i, row)
        Next
    End Sub
    Private Shared Sub GenerateOrderRow(rowIndex As Integer, row As WTableRow)
        Dim random As New Random()
        row.Cells(0).AddParagraph().AppendText(rowIndex.ToString())
        row.Cells(1).AddParagraph().AppendText(random.Next(1000, 9999).ToString())
        row.Cells(2).AddParagraph().AppendText(random.Next(10, 100).ToString("F2"))
        row.Cells(3).AddParagraph().AppendText(random.Next(1, 10).ToString())
        row.Cells(4).AddParagraph().AppendText((Convert.ToDecimal(row.Cells(2).LastParagraph.Text) * Convert.ToDecimal(row.Cells(3).LastParagraph.Text)).ToString("F2"))
    End Sub
    Private Shared Sub GenerateJobRow(rowIndex As Integer, row As WTableRow)
        Dim random As New Random()
        row.Cells(0).AddParagraph().AppendText(rowIndex.ToString())
        row.Cells(1).AddParagraph().AppendText(random.Next(2000, 3000).ToString())
        row.Cells(2).AddParagraph().AppendText("Job " & rowIndex)
        row.Cells(3).AddParagraph().AppendText(random.Next(100, 200).ToString())
        row.Cells(4).AddParagraph().AppendText(random.Next(1, 100).ToString())
    End Sub
    Private Shared Sub GenerateInventoryRow(rowIndex As Integer, row As WTableRow)
        Dim random As New Random()
        row.Cells(0).AddParagraph().AppendText(rowIndex.ToString())
        row.Cells(1).AddParagraph().AppendText(random.Next(3000, 4000).ToString())
        row.Cells(2).AddParagraph().AppendText("Quality " & Chr(65 + random.Next(0, 5)).ToString())
        row.Cells(3).AddParagraph().AppendText(random.Next(1, 50).ToString())
        row.Cells(4).AddParagraph().AppendText(random.Next(1000, 40000).ToString())
    End Sub
End Class
