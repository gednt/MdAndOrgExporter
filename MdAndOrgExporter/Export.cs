using MdAndOrgExporter.Functions;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Paragraph = Microsoft.Office.Interop.Word.Paragraph;

namespace MdAndOrgExporter
{

    public partial class Export
    {

        public Microsoft.Office.Interop.Word.Document Document { get; set; }

        public Microsoft.Office.Tools.Word.Document VstoDocument { get; set; }

        public List<Tuple<String, String, Paragraph>> Images = new List<Tuple<string, string,Paragraph>>();

        private void Export_Load(object sender, RibbonUIEventArgs e)
        {
            Images = new List<Tuple<string, string, Paragraph>>();
        }
        

        private void ExportToOrg_Click(object sender, RibbonControlEventArgs e)
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

            using (SaveFileDialog fd = new SaveFileDialog())
            {
                Images = new List<Tuple<string, string, Paragraph>>();
                fd.Title = "Export to Org";
                fd.DefaultExt = ".org";
                fd.FileName = Document.Name.Split('.')[0] + ".org";
                fd.InitialDirectory = "%Documents%";
                fd.Filter = "|*.org";
                DialogResult result = fd.ShowDialog();

                for (int i = 0; i < Document.InlineShapes.Count; i++)
                {
                    InlineShape image = Document.InlineShapes[i + 1];
                    var imageRange = image.Range;
                    Microsoft.Office.Interop.Word.Paragraph info = null;
                    foreach (Microsoft.Office.Interop.Word.Paragraph range in imageRange.Paragraphs)
                    {
                        info = range;
                    }
                    if (image.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        System.Drawing.Image clipboardImage = null;
                        var imageName = Document.Name;
                        imageRange.Select();
                        image.Application.Selection.Copy();
                        try
                        {
                            if (Clipboard.ContainsImage())
                            {
                                clipboardImage = Clipboard.GetImage();

                                if (!Directory.Exists(Path.Combine(Path.GetDirectoryName(fd.FileName), "assets")))
                                {
                                    Directory.CreateDirectory(Path.Combine(Path.GetDirectoryName(fd.FileName), "assets"));
                                }
                                var path = Path.Combine(Path.Combine(Path.GetDirectoryName(fd.FileName), "assets"), $"{imageName}{i}.jpg");
                                if (!File.Exists($"{imageName}{i}") && clipboardImage != null)
                                {
                                    clipboardImage.Save(path);
                                }
                                Images.Add(new Tuple<String, String, Paragraph>($"image{i}.jpg", $"{imageName}{i}.jpg", info));


                            }
                        }
                        catch
                        {

                        }

                    }



                }
                List<Functions.Paragraph> paragraphs = new List<Functions.Paragraph>();
                foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in Document.Paragraphs)
                {
                    if (paragraph != null)
                    {
                        foreach(var image in Images)
                        {
                            if ((paragraph.Range.Start == image.Item3.Range.Start) &&paragraph.Range.InlineShapes.Count>0)
                            {
                                paragraphs.Add(new Functions.Paragraph()
                                {
                                    ContainsImage = true,
                                    Text = image.Item2,
                                    ListFormat = paragraph.Range.ListFormat,
                                    List = paragraph.Range.ListFormat.List
                                });
                            }
                        }
                        paragraphs.Add(new Functions.Paragraph()
                        {
                            Text = paragraph.Range.Text,
                            Identation = paragraph.LeftIndent,
                            ListFormat = paragraph.Range.ListFormat,
                            List = paragraph.Range.ListFormat.List
                        });
                    }

                }
                var textToExport = new StringBuilder();
                OrgModeUtilities orgModeUtilities = new OrgModeUtilities();
                textToExport.Append(orgModeUtilities.CreatePreamble(Document.Name.Replace(".docx", "").Replace(".doc", ""), "DMF - Export to MD"));
                System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        textToExport.Append("\n* tags: ");
                        if (chkSeparateTags.Checked)
                        {
                            foreach (var tag in txtTags.Text.Split('/'))
                            {
                                textToExport.Append(" [[" + tag + "]] ");
                            }
                        }
                        else
                        {
                            textToExport.Append(" [[" + txtTags.Text + "]] ");
                        }
                        textToExport.Append('\n');
                    }
                    catch
                    {

                    }

                    textToExport.Append(orgModeUtilities.CreateHeading(paragraphs));
                    var textReadyToBeExported = textToExport.ToString();
                    if (result == DialogResult.OK)
                    {
                        if(!Directory.Exists(Path.Combine(Path.GetDirectoryName(fd.FileName),"pages")))
                        {
                            Directory.CreateDirectory(Path.Combine(Path.GetDirectoryName(fd.FileName), "pages"));
                        }
                        var saveTo = Path.Combine(Path.GetDirectoryName(fd.FileName), "pages", Path.GetFileName(fd.FileName));
                        using (StreamWriter sw = new StreamWriter(saveTo, false))
                        {
                            sw.Write(textReadyToBeExported);
                            sw.Close();
                        }
                    }
                });
            }


        }

        private void btnExportToMd_Click(object sender, RibbonControlEventArgs e)
        {
            using (SaveFileDialog fd = new SaveFileDialog())
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
                fd.Title = "Export to Markdown";
                fd.DefaultExt = ".md";
                fd.FileName = Document.Name.Split('.')[0] + ".md";
                fd.InitialDirectory = "%Documents%";
                fd.Filter = "|*.md";
                DialogResult result = fd.ShowDialog();


                List<Functions.Paragraph> paragraphs = new List<Functions.Paragraph>();
                foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in Document.Paragraphs)
                {
                    if (paragraph != null)
                    {
                        paragraphs.Add(new Functions.Paragraph()
                        {
                            Text = paragraph.Range.Text,
                            Identation = paragraph.LeftIndent,
                            ListFormat = paragraph.Range.ListFormat,
                            List = paragraph.Range.ListFormat.List
                        });
                    }
                }
                var textToExport = new StringBuilder();

                MdModeUtilities mdModeUtilities = new MdModeUtilities();
                textToExport.Append(mdModeUtilities.CreatePreamble(Document.Name.Replace(".docx", "").Replace(".doc", ""), "DMF - Export to MD"));
                try
                {
                    textToExport.Append("\n* tags: ");
                    if (chkSeparateTags.Checked)
                    {
                        foreach (var tag in txtTags.Text.Split('/'))
                        {
                            textToExport.Append(" [[" + tag + "]] ");
                        }
                    }
                    else
                    {
                        textToExport.Append(" [[" + txtTags.Text + "]] ");
                    }
                    textToExport.Append("\n- ");
                }
                catch
                {

                }

                textToExport.Append(mdModeUtilities.CreateHeading(paragraphs));
                var textReadyToBeExported = textToExport.ToString();
                if (result == DialogResult.OK)
                {

                    using (StreamWriter sw = new StreamWriter(fd.FileName, false))
                    {
                        sw.Write(textReadyToBeExported);
                        sw.Close();
                    }
                }
            }

        }
    }
}
