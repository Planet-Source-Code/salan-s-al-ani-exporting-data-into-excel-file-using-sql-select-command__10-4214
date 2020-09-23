<div align="center">

## Exporting data into Excel file using SQL SELECT command


</div>

### Description

If you want to have you selection of data in your database to be available in an Excel format, this code will help you.

I tried this code usin VB.NET 2003, .NET Framework v1.1, and Microsoft Excel v11.0 Object Library.

Please feel free to email me if you have any comments.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Salan S\. Al\-Ani](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/salan-s-al-ani.md)
**Level**          |Intermediate
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |VB\.NET
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__10-5.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/salan-s-al-ani-exporting-data-into-excel-file-using-sql-select-command__10-4214/archive/master.zip)





### Source Code

```
'You should add the reference "Microsoft Excel 11.0 Object Library" to your project
Imports System
Imports System.IO
Public Class frmMain
 Inherits System.Windows.Forms.Form
 Dim My_Connection As New System.Data.OleDb.OleDbConnection
 Dim sql As String
 Dim cmd As System.Data.OleDb.OleDbCommand
 Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
	'Defining the connection string (you can use your own database provider and data source type)
	My_Connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=C:\Test.mdb"
  Try
   My_Connection.Open()
  Catch ex As Exception
   MsgBox("Failed to connect to database file: " & ex.ToString(), MsgBoxStyle.Critical, "Error")
   End
  End Try
 End Sub
 Private Sub frmMain_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
  If My_Connection.State <> ConnectionState.Closed Then My_Connection.Close()
 End Sub
 Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
  On Error GoTo ErrorHandler
  Dim ExcelFile_Path As String
  Dim MySaveFileDialog As New SaveFileDialog
	'Setting the attributes for "MySaveFileDialog" object
  MySaveFileDialog.Filter = "Microsoft Office Excel Files (*.xls)|*.xls"
  MySaveFileDialog.FilterIndex = 1
  MySaveFileDialog.CheckFileExists = False
  MySaveFileDialog.CheckPathExists = True
  MySaveFileDialog.CreatePrompt = False
  MySaveFileDialog.DefaultExt = "xls"
  MySaveFileDialog.DereferenceLinks = True
  MySaveFileDialog.InitialDirectory = App_Path
  MySaveFileDialog.OverwritePrompt = True
  MySaveFileDialog.RestoreDirectory = True
  MySaveFileDialog.ValidateNames = True
  If MySaveFileDialog.ShowDialog() <> DialogResult.OK Then Exit Sub
  ExcelFile_Path = MySaveFileDialog.FileName
	If File.Exists(ExcelFile_Path) = True Then File.Delete(ExcelFile_Path)
  Me.Cursor = Cursors.WaitCursor
  Dim xlApp As Microsoft.Office.Interop.Excel.Application
  Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
  Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet
	'Create an empty Excel file in the path specified by the "ExcelFile_Path" string which obtained from the "MySaveFileDialog" object
	'Usually the new Excel file will contain three empty sheets (sheet1, sheet2, and sheet3)
	xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
  xlBook = CType(xlApp.Workbooks.Add, Microsoft.Office.Interop.Excel.Workbook)
  xlSheet = CType(xlBook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
  xlSheet.SaveAs(ExcelFile_Path)
  xlBook.Close()
	'Exporting the data into Excel file using SQL SELECT command
	'In this example the data will be exported in a new sheet named "EXPORT"
  sql = "SELECT * INTO [Excel 8.0;Database=" & ExcelFile_Path & "].[EXPORT] FROM TEST_TABLE ORDER BY FIELD1,FIELD2"
  cmd = New System.Data.OleDb.OleDbCommand(sql, My_Connection)
  cmd.ExecuteNonQuery()
	'Reopen the Excel file
  xlApp = New Microsoft.Office.Interop.Excel.Application
  xlApp = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
  xlBook = xlApp.Workbooks.Open(ExcelFile_Path)
	'Delete all sheets in the Excel file except the "EXPORT" sheet
  For Each xlSheet In xlBook.Worksheets
   If xlSheet.Name <> "EXPORT" Then xlSheet.Delete()
  Next
  xlSheet = New Microsoft.Office.Interop.Excel.Worksheet
  xlSheet = CType(xlBook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
  Dim xlRange As Microsoft.Office.Interop.Excel.Range
	'Make the first row in the "EXPORT" sheet as bold (this is recommended if the first row is a header row)
  xlRange = DirectCast(xlSheet.Rows(1), Microsoft.Office.Interop.Excel.Range)
  xlRange.Font.Bold = True
	'Auto fit the whole columns in the "EXPORT" sheet
  xlRange = DirectCast(xlSheet.Columns, Microsoft.Office.Interop.Excel.Range)
  xlRange.AutoFit()
	'Save and close the Excel file
  xlBook.Save()
  xlBook.Close()
  Me.Cursor = Cursors.Default
  Exit Sub
ErrorHandler:
  Me.Cursor = Cursors.Default
  xlBook.Close(SaveChanges:=False)
  MsgBox("Operation failed", MsgBoxStyle.Critical, "Error")
 End Sub
End Class
```

