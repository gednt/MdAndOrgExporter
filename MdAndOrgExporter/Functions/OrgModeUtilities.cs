using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExtensionMethods;
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
            foreach (var paragraph in Paragraphs)
            {
                switch (paragraph.ParagraphFormat.OutlineLevel)
                {
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(1, '*'));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(2 + (int)Math.Round(paragraph.Identation), '*'));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(3 + (int)Math.Round(paragraph.Identation), '*'));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(4 + (int)Math.Round(paragraph.Identation), '*'));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel5:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(4 + (int)Math.Round(paragraph.Identation), '*'));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel6:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(4 + (int)Math.Round(paragraph.Identation), '*'));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel7:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(4 + (int)Math.Round(paragraph.Identation), '*'));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel8:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(4 + (int)Math.Round(paragraph.Identation), '*'));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel9:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(4 + (int)Math.Round(paragraph.Identation), '*'));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(5 + (int)Math.Round(paragraph.Identation), '*'));
                        break;
                    default:
                        break;
                }
                //for (int i = 0; i < levels; i++)
                //{
                //    textToReturn.Append(paragraph.IdentationCharacter);
                //}
                if(paragraph.Text != "\r" && paragraph.Text != "/\r")
                {
                    if (paragraph.ContainsImage == false)
                    {
                        textToReturn.Append(" " + (paragraph.ListFormat != null ? paragraph.ListFormat.ListString + " " : "") + paragraph.Text + " \n");
                    }
                    else
                    {
                        textToReturn.Append(" " + (paragraph.ListFormat != null ? paragraph.ListFormat.ListString + " " : "") + $"![{Path.GetFileName(paragraph.Text)}](../assets/{paragraph.Text})" + " \n");
                    }
                }


            }
            return textToReturn.ToString();
        }
    }
}
