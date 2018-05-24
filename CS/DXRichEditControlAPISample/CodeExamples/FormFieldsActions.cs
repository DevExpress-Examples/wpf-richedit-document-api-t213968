using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXRichEditControlAPISample.CodeExamples
{
    class FormFieldsActions
    {
        static void CreateCheckbox(Document document)
        {
            #region #CreateCheckbox
            DocumentPosition currentPosition = document.CaretPosition;
            DevExpress.XtraRichEdit.API.Native.CheckBox checkBox = document.FormFields.InsertCheckBox(currentPosition);
            checkBox.Name = "check1";
            checkBox.State = CheckBoxState.Checked;
            checkBox.SizeMode = CheckBoxSizeMode.Auto;
            checkBox.HelpTextType = FormFieldTextType.Custom;
            checkBox.HelpText = "help text";
            #endregion #CreateCheckbox
        }
    }
}
