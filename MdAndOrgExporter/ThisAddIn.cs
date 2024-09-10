using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using MdAndOrgExporter.Functions;

namespace MdAndOrgExporter
{
    public partial class ThisAddIn
    {
        public Word.Document Document { get; set; }

        public Microsoft.Office.Tools.Word.Document VstoDocument { get; set; }

      protected override Microsoft.Office.Core.IRibbonExtensibility
      CreateRibbonExtensibilityObject()
        {
                return Globals.Factory.GetRibbonFactory().CreateRibbonManager(
                    new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { new Export() });
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Documents.Count > 0)
            {
                Microsoft.Office.Interop.Word.Document nativeDocument =
                    Globals.ThisAddIn.Application.ActiveDocument;
                Microsoft.Office.Tools.Word.Document vstoDocument =
                    Globals.Factory.GetVstoObject(nativeDocument);

                Document = nativeDocument;
                VstoDocument = vstoDocument;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
