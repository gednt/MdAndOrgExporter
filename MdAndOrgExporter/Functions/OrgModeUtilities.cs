using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExtensionMethods;
using Microsoft.Office.Interop.Word;
namespace MdAndOrgExporter.Functions
{
    public class OrgModeUtilities
    {
        public String CreatePreamble(String document_title, String document_author)
        {
            return $"#+alias: {document_title}{System.Environment.NewLine}" +
                $"#+title: {document_title}{System.Environment.NewLine}" +
                $"#+author: {document_author}{System.Environment.NewLine}";
        }

    /// <summary>
    /// The first value is the text. The second value is the identation_level
    /// </summary>
    /// <param name="Paragraphs"></param>
    /// <returns></returns>
        public String CreateHeading(List<Paragraph>Paragraphs)
        {
            var textToReturn = new StringBuilder();
            List<int> outLineIndentLevel = new List<int>();
            for (int cont=0;cont<Paragraphs.Count;cont++)
            {
                System.Windows.Forms.Application.DoEvents();
                WdStyleType type = Paragraphs[cont].Style.Type;
                var identLevel = 0;
                identLevel = Paragraphs[cont].Type() > 0 ? Paragraphs[cont].Type():3 + (int)Paragraphs[cont].Identation;
                identLevel = Paragraphs[cont].ParagraphFormat.Alignment==WdParagraphAlignment.wdAlignParagraphCenter?1:
                    Paragraphs[cont].ParagraphFormat.Alignment == WdParagraphAlignment.wdAlignParagraphRight?1:identLevel;

                    
                
                bool justNullOrEmpty = String.IsNullOrWhiteSpace(Paragraphs[cont].Text);
                if(!justNullOrEmpty)
                    textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(identLevel, Paragraphs[cont].IdentationCharacter));

                //for (int i = 0; i < levels; i++)
                //{
                //    textToReturn.Append(Paragraphs[cont].IdentationCharacter);
                //}
                if (!justNullOrEmpty)
                {
                    if (Paragraphs[cont].Text != "\r" && Paragraphs[cont].Text != "/\r")
                    {
                        if (Paragraphs[cont].ContainsImage == false)
                        {
                            textToReturn.Append(" " + (Paragraphs[cont].ListFormat != null ? Paragraphs[cont].ListFormat.ListString + " " : "") + Paragraphs[cont].Text + " \n");
                        }
                        else
                        {
                            textToReturn.Append(" " + (Paragraphs[cont].ListFormat != null ? Paragraphs[cont].ListFormat.ListString + " " : "") + $"![{Path.GetFileName(Paragraphs[cont].Text)}](../assets/{Paragraphs[cont].Text})" + " \n");
                        }
                    }
                    if (Paragraphs[cont].Footnotes != null)
                    {
                        foreach(Footnote footnote in Paragraphs[cont].Footnotes)
                        {
                            textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(identLevel + 1, Paragraphs[cont].IdentationCharacter));
                            textToReturn.Append(" #+BEGIN_PINNED");
                            textToReturn.AppendLine();
                            textToReturn.Append(footnote.Range.Text);
                            textToReturn.AppendLine();
                            textToReturn.Append(" #+END_PINNED");
                            textToReturn.AppendLine();

                        }
                    }
 

                }



            }
            return textToReturn.ToString();
        }

    }
}
