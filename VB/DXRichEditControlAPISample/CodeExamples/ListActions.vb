Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace DXRichEditControlAPISample.CodeExamples
    Friend Class ListActions
        Private Shared Sub CreateBulletedList(ByVal document As Document)
            '            #Region "#CreateBulletedList"

            document.LoadDocument("Documents//List.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            document.BeginUpdate()
            ' Define an abstract list that is the pattern for lists used in the document.
            Dim list As AbstractNumberingList = document.AbstractNumberingLists.Add()
            list.NumberingType = NumberingType.Bullet

            ' Specify parameters for each list level.

            Dim level As ListLevel = list.Levels(0)
            level.ParagraphProperties.LeftIndent = 100
            level.CharacterProperties.FontName = "Symbol"
            level.DisplayFormatString = New String(ChrW(&H00B7), 1)

            level = list.Levels(1)
            level.ParagraphProperties.LeftIndent = 250
            level.CharacterProperties.FontName = "Symbol"
            level.DisplayFormatString = New String(ChrW(&H006F), 1)

            level = list.Levels(2)
            level.ParagraphProperties.LeftIndent = 450
            level.CharacterProperties.FontName = "Symbol"
            level.DisplayFormatString = New String(ChrW(&H00B7), 1)

            ' Create a list for use in the document. It is based on a previously defined abstract list with ID = 0.
            document.NumberingLists.Add(0)
            document.EndUpdate()

            document.AppendText("Line 1" & vbLf & "Line 2" & vbLf & "Line 3")
            ' Convert all paragraphs to list items.
            document.BeginUpdate()
            Dim paragraphs As ParagraphCollection = document.Paragraphs
            For Each pgf As Paragraph In paragraphs
                pgf.ListIndex = 0
                pgf.ListLevel = 1
            Next pgf
            document.EndUpdate()
'            #End Region ' #CreateBulletedList
        End Sub


        Private Shared Sub CreateNumberedList(ByVal document As Document)
            '            #Region "#CreateNumberedList"
            document.LoadDocument("Documents//List.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)
            document.BeginUpdate()
            'Create a new pattern object
            Dim abstractListNumberingRoman As AbstractNumberingList = document.AbstractNumberingLists.Add()

            'Specify the list's type
            abstractListNumberingRoman.NumberingType = NumberingType.Simple

            'Define the first level's properties
            Dim level As ListLevel = abstractListNumberingRoman.Levels(0)
            level.ParagraphProperties.LeftIndent = 150
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 75
            level.Start = 1
            level.NumberingFormat = NumberingFormat.LowerRoman
            level.DisplayFormatString = "{0}."

            'Create a new list based on the specific pattern
            Dim numberingList As NumberingList = document.NumberingLists.Add(0)
            document.EndUpdate()

            document.BeginUpdate()
            Dim paragraphs As ParagraphCollection = document.Paragraphs

            'Add paragraphs to the list
            paragraphs.AddParagraphsToList(document.Range, numberingList, 0)
            document.EndUpdate()
            '            #End Region ' #CreateNumberedList
        End Sub

        Private Shared Sub CreateMultilevelList(ByVal document As Document)
            ' #Region "#CreateMultilevelList"
            document.LoadDocument("Documents//List.docx", DevExpress.XtraRichEdit.DocumentFormat.OpenXml)

            document.BeginUpdate()

            'Create a new list pattern object
            Dim list As AbstractNumberingList = document.AbstractNumberingLists.Add()

            'Specify the list's type
            list.NumberingType = NumberingType.MultiLevel

            'Specify parameters for each level
            Dim level As ListLevel = list.Levels(0)
            level.ParagraphProperties.LeftIndent = 105
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 55
            level.Start = 1
            level.NumberingFormat = NumberingFormat.UpperRoman
            level.DisplayFormatString = "{0}"

            level = list.Levels(1)
            level.ParagraphProperties.LeftIndent = 125
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 65
            level.Start = 1
            level.NumberingFormat = NumberingFormat.LowerRoman
            level.DisplayFormatString = "{1})"

            level = list.Levels(2)
            level.ParagraphProperties.LeftIndent = 145
            level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging
            level.ParagraphProperties.FirstLineIndent = 75
            level.Start = 1
            level.NumberingFormat = NumberingFormat.LowerLetter
            level.DisplayFormatString = "{2}."

            'Create a new list object based on the specified pattern
            document.NumberingLists.Add(0)
            document.EndUpdate()

            document.BeginUpdate()

            ' Apply numbering to a list
            Dim paragraphs As ParagraphCollection = document.Paragraphs

            For Each pgf As Paragraph In paragraphs
                pgf.ListIndex = 0
                pgf.ListLevel = pgf.Index
            Next

            document.EndUpdate()
            ' #End Region ' #CreateMultilevelList
        End Sub



    End Class
End Namespace
