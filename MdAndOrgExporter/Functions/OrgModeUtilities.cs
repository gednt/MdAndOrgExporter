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
                if (Paragraphs[cont].Style.NameLocal.Contains("Heading"))
                {
                    identLevel = Paragraphs[cont].Type();
                    outLineIndentLevel.Add(identLevel);
                }
                else
                {
                    if (outLineIndentLevel.Count > 0)
                    {
                        identLevel = outLineIndentLevel.Last() + 1;
                    }
                    else
                    {
                        identLevel = 1 + (int)Paragraphs[cont].Identation;
                    }
                }

                textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(identLevel, Paragraphs[cont].IdentationCharacter));
                //for (int i = 0; i < levels; i++)
                //{
                //    textToReturn.Append(Paragraphs[cont].IdentationCharacter);
                //}
                if(Paragraphs[cont].Text != "\r" && Paragraphs[cont].Text != "/\r")
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


            }
            return textToReturn.ToString();
        }
    }
}
