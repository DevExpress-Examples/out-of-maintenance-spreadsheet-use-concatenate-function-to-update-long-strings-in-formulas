Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Linq

Namespace SpreadsheetApp
	Partial Public Class Form1
		Inherits DevExpress.XtraBars.Ribbon.RibbonForm
		Public Sub New()
			InitializeComponent()
			spreadsheetControl1.Options.Compatibility.TruncateLongStringsInFormulas = False
			spreadsheetControl1.LoadDocument("FileWithLongStringInFormula.xlsx")
		End Sub

		Private Sub spreadsheetControl1_DocumentLoaded(ByVal sender As Object, ByVal e As EventArgs) Handles spreadsheetControl1.DocumentLoaded
			Dim workbook = spreadsheetControl1.Document
			workbook.BeginUpdate()
			For Each sheet In workbook.Worksheets
				For Each cell In sheet.GetExistingCells().Where(Function(x) x.HasFormula)
					Dim walker = New ReplaceLongStringsWalker(workbook)
					Dim parsedExpression = cell.ParsedExpression
					walker.Walk(parsedExpression)
					If walker.IsExpressionModified Then
						cell.ParsedExpression = parsedExpression
					End If
				Next cell
			Next sheet
			workbook.EndUpdate()
		End Sub
	End Class
End Namespace
