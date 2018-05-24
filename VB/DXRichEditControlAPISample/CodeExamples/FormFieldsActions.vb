Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace DXRichEditControlAPISample.CodeExamples
    Friend Class FormFieldsActions
        Private Shared Sub CreateCheckbox(ByVal document As Document)
'            #Region "#CreateCheckbox"
            Dim currentPosition As DocumentPosition = document.CaretPosition
            Dim checkBox As DevExpress.XtraRichEdit.API.Native.CheckBox = document.FormFields.InsertCheckBox(currentPosition)
            checkBox.Name = "check1"
            checkBox.State = CheckBoxState.Checked
            checkBox.SizeMode = CheckBoxSizeMode.Auto
            checkBox.HelpTextType = FormFieldTextType.Custom
            checkBox.HelpText = "help text"
'            #End Region ' #CreateCheckbox
        End Sub
    End Class
End Namespace
