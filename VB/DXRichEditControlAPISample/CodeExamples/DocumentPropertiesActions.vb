Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.Xpf.RichEdit

Namespace DXRichEditControlAPISample.CodeExamples
    Friend Class DocumentPropertiesActions
        Private Shared Sub StandardDocumentProperties(ByVal document As Document)
'            #Region "#StandardDocumentProperties"
            document.BeginUpdate()

            document.DocumentProperties.Creator = "John Doe"
            document.DocumentProperties.Title = "Inserting Custom Properties"
            document.DocumentProperties.Category = "TestDoc"
            document.DocumentProperties.Description = "This code demonstrates API to modify and display standard document properties."

            document.Fields.Create(document.AppendText(ControlChars.Lf & "AUTHOR: ").End, "AUTHOR")
            document.Fields.Create(document.AppendText(ControlChars.Lf & "TITLE: ").End, "TITLE")
            document.Fields.Create(document.AppendText(ControlChars.Lf & "COMMENTS: ").End, "COMMENTS")
            document.Fields.Create(document.AppendText(ControlChars.Lf & "CREATEDATE: ").End, "CREATEDATE")
            document.Fields.Create(document.AppendText(ControlChars.Lf & "Category: ").End, "DOCPROPERTY Category")
            document.Fields.Update()
            document.EndUpdate()
'            #End Region ' #StandardDocumentProperties
        End Sub


        Private Shared Sub CustomDocumentProperties(ByVal document As Document)
'            #Region "#CustomDocumentProperties"
            document.BeginUpdate()
            document.AppendText("A new value of MyBookmarkProperty is obtained from here: NEWVALUE!" & ControlChars.Lf)
            document.Bookmarks.Create(document.FindAll("NEWVALUE!", SearchOptions.CaseSensitive)(0), "bmOne")
            document.AppendText(ControlChars.Lf & "MyNumericProperty: ")
            document.Fields.Create(document.Range.End, "DOCVARIABLE CustomProperty MyNumericProperty")
            document.AppendText(ControlChars.Lf & "MyStringProperty: ")
            document.Fields.Create(document.Range.End, "DOCVARIABLE CustomProperty MyStringProperty")
            document.AppendText(ControlChars.Lf & "MyBooleanProperty: ")
            document.Fields.Create(document.Range.End, "DOCVARIABLE CustomProperty MyBooleanProperty")
            document.AppendText(ControlChars.Lf & "MyBookmarkProperty: ")
            document.Fields.Create(document.Range.End, "DOCVARIABLE CustomProperty MyBookmarkProperty")
            document.EndUpdate()

            document.CustomProperties("MyNumericProperty") = 123.45
            document.CustomProperties("MyStringProperty") = "The Final Answer"
            document.CustomProperties("MyBookmarkProperty") = document.Bookmarks(0)
            document.CustomProperties("MyBooleanProperty") = True

            AddHandler document.CalculateDocumentVariable, AddressOf DocumentPropertyDisplayHelper.OnCalculateDocumentVariable
            document.Fields.Update()
'            #End Region ' #CustomDocumentProperties
        End Sub

        #Region "#@CustomDocumentProperties"
        Private Class DocumentPropertyDisplayHelper
            Public Shared Sub OnCalculateDocumentVariable(ByVal sender As Object, ByVal e As DevExpress.XtraRichEdit.CalculateDocumentVariableEventArgs)
                If e.Arguments.Count = 0 OrElse e.VariableName <> "CustomProperty" Then
                    Return
                End If
                Dim name As String = e.Arguments(0).Value
                Dim customProperty = DirectCast(sender, DevExpress.Xpf.RichEdit.RichEditControl).Document.CustomProperties(name)
                If customProperty IsNot Nothing Then
                    e.Value = customProperty.ToString()
                End If
                e.Handled = True
            End Sub
        End Class
        #End Region ' #@CustomDocumentProperties
    End Class
End Namespace
