﻿Imports Microsoft.VisualBasic
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace DXRichEditControlAPISample.CodeExamples
	Friend Class TableActions
		Private Shared Sub CreateTable(ByVal document As Document)
'			#Region "#CreateTable"
			' Insert new table.
			Dim tbl As Table = document.Tables.Create(document.Range.Start, 1, 3, AutoFitBehaviorType.AutoFitToWindow)
			' Create a table header.
			document.InsertText(tbl(0, 0).Range.Start, "Name")
			document.InsertText(tbl(0, 1).Range.Start, "Size")
			document.InsertText(tbl(0, 2).Range.Start, "DateTime")
			' Insert table data.
			Dim dirinfo As New DirectoryInfo("C:\")
			Try
				tbl.BeginUpdate()
				For Each fi As FileInfo In dirinfo.GetFiles()
					Dim row As TableRow = tbl.Rows.Append()
					Dim cell As TableCell = row.FirstCell
					Dim fileName As String = fi.Name
					Dim fileLength As String = String.Format("{0:N0}", fi.Length)
					Dim fileLastTime As String = String.Format("{0:g}", fi.LastWriteTime)
					document.InsertSingleLineText(cell.Range.Start, fileName)
					document.InsertSingleLineText(cell.Next.Range.Start, fileLength)
					document.InsertSingleLineText(cell.Next.Next.Range.Start, fileLastTime)
				Next fi
				' Center the table header.
				For Each p As Paragraph In document.Paragraphs.Get(tbl.FirstRow.Range)
					p.Alignment = ParagraphAlignment.Center
				Next p
			Finally
				tbl.EndUpdate()
			End Try
'			#End Region ' #CreateTable
		End Sub

		Private Shared Sub CreateFixedTable(ByVal document As Document)
'			#Region "#CreateFixedTable"
			Dim table As Table = document.Tables.Create(document.Range.Start, 3, 3)

			table.TableAlignment = TableRowAlignment.Center
			table.TableLayout = TableLayoutType.Fixed
			table.PreferredWidthType = WidthType.Fixed
			table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4f)

			table.Rows(1).HeightType = HeightType.Exact
			table.Rows(1).Height = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.8f)

			table(1, 1).PreferredWidthType = WidthType.Fixed
			table(1, 1).PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f)
'			#End Region ' #CreateFixedTable
		End Sub

		Private Shared Sub CreateAndApplyTableStyle(ByVal document As Document)
'			#Region "#CreateAndApplyTableStyle"
			document.BeginUpdate()
			' Create a new table style.
			Dim tStyleMain As TableStyle = document.TableStyles.CreateNew()
			' Specify style characteristics.
			tStyleMain.AllCaps = True
			tStyleMain.FontName = "Segoe Condensed"
			tStyleMain.FontSize = 14
			tStyleMain.Alignment = ParagraphAlignment.Center
			tStyleMain.TableBorders.InsideHorizontalBorder.LineStyle = TableBorderLineStyle.Dotted
			tStyleMain.TableBorders.InsideVerticalBorder.LineStyle = TableBorderLineStyle.Dotted
			tStyleMain.TableBorders.Top.LineThickness = 1.5f
			tStyleMain.TableBorders.Top.LineStyle = TableBorderLineStyle.Double
			tStyleMain.TableBorders.Left.LineThickness = 1.5f
			tStyleMain.TableBorders.Left.LineStyle = TableBorderLineStyle.Double
			tStyleMain.TableBorders.Bottom.LineThickness = 1.5f
			tStyleMain.TableBorders.Bottom.LineStyle = TableBorderLineStyle.Double
			tStyleMain.TableBorders.Right.LineThickness = 1.5f
			tStyleMain.TableBorders.Right.LineStyle = TableBorderLineStyle.Double
			tStyleMain.CellBackgroundColor = System.Drawing.Color.LightBlue
			tStyleMain.TableLayout = TableLayoutType.Fixed
			tStyleMain.Name = "MyTableStyle"
			'Add the style to the document.
			document.TableStyles.Add(tStyleMain)
			document.EndUpdate()
			document.BeginUpdate()
			' Create a table.
			Dim table As Table = document.Tables.Create(document.Range.Start, 3, 3)
			table.TableLayout = TableLayoutType.Fixed
			table.PreferredWidthType = WidthType.Fixed
			table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4.5f)
			table(1, 1).PreferredWidthType = WidthType.Fixed
			table(1, 1).PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f)
			' Apply a previously defined style.
			table.Style = tStyleMain
			document.EndUpdate()

			document.InsertText(table(1, 1).Range.Start, "STYLED")
'			#End Region ' #CreateAndApplyTableStyle
		End Sub

		Private Shared Sub UseConditionalStyle(ByVal document As Document)
'			#Region "#UseConditionalStyle"
			document.LoadDocument("TableStyles.docx", DocumentFormat.OpenXml)
			document.BeginUpdate()

			' Create a new style that is based on the 'Grid Table 5 Dark Accent 1' style defined in the loaded document.
			Dim myNewStyle As TableStyle = document.TableStyles.CreateNew()
			myNewStyle.Parent = document.TableStyles("Grid Table 5 Dark Accent 1")
			' Create conditional styles (styles for table elements)
			Dim myNewStyleForFirstRow As TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstRow)
			myNewStyleForFirstRow.CellBackgroundColor = Color.PaleVioletRed
			Dim myNewStyleForFirstColumn As TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstColumn)
			myNewStyleForFirstColumn.CellBackgroundColor = Color.PaleVioletRed
			Dim myNewStyleForOddColumns As TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.OddColumnBanding)
			myNewStyleForOddColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.Light(Color.PaleVioletRed)
			Dim myNewStyleForEvenColumns As TableConditionalStyle = myNewStyle.ConditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.EvenColumnBanding)
			myNewStyleForEvenColumns.CellBackgroundColor = System.Windows.Forms.ControlPaint.LightLight(Color.PaleVioletRed)
			document.TableStyles.Add(myNewStyle)
			' Create a new table and apply a new style.
			Dim table As Table = document.Tables.Create(document.Range.End, 4, 4, AutoFitBehaviorType.AutoFitToWindow)
			table.Style = myNewStyle
			' Specify which conditonal styles are in effect.
			table.TableLook = TableLookTypes.ApplyFirstRow Or TableLookTypes.ApplyFirstColumn

			document.EndUpdate()
'			#End Region ' #UseConditionalStyle
		End Sub

		Private Shared Sub ChangeColumnAppearance(ByVal document As Document)
'			#Region "#ChangeColumnAppearance"
			Dim table As Table = document.Tables.Create(document.Range.Start, 3, 10)
			table.BeginUpdate()
			'Change cell background color and vertical alignment in the third column.
			table.ForEachRow(New TableRowProcessorDelegate(AddressOf ChangeColumnAppearanceHelper.ChangeColumnColor))
			table.EndUpdate()
'			#End Region ' #ChangeColumnAppearance

		End Sub
		#Region "#@ChangeColumnAppearance"
		Public Class ChangeColumnAppearanceHelper
			Public Shared Sub ChangeColumnColor(ByVal row As TableRow, ByVal rowIndex As Integer)
				row(2).BackgroundColor = System.Drawing.Color.LightCyan
				row(2).VerticalAlignment = TableCellVerticalAlignment.Center
			End Sub
		End Class
		#End Region ' #@ChangeColumnAppearance

		Private Shared Sub UseTableCellProcessor(ByVal document As Document)
'			#Region "#UseTableCellProcessor"
			Dim table As Table = document.Tables.Create(document.Range.Start, 8, 8)
			table.BeginUpdate()
			table.ForEachCell(New TableCellProcessorDelegate(AddressOf UseTableCellProcessorHelper.MakeMultiplicationCell))
			table.EndUpdate()
'			#End Region ' #UseTableCellProcessor
		End Sub
		#Region "#@UseTableCellProcessor"
		Public Class UseTableCellProcessorHelper
			Public Shared Sub MakeMultiplicationCell(ByVal cell As TableCell, ByVal i As Integer, ByVal j As Integer)
				Dim doc As SubDocument = cell.Range.BeginUpdateDocument()
				doc.InsertText(cell.Range.Start, String.Format("{0}*{1} = {2}", i + 2, j + 2, (i + 2) * (j + 2)))
				cell.Range.EndUpdateDocument(doc)
			End Sub
		End Class
		#End Region ' #@UseTableCellProcessor
	End Class
End Namespace
