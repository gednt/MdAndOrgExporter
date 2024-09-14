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
            for (int cont=0;cont<Paragraphs.Count;cont++)
            {
                switch (Paragraphs[cont].ParagraphFormat.OutlineLevel)
                {
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(1, '*'));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(cont>0 &&!Paragraphs[cont - 1].ParagraphFormat.Equals(Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2) ? (int)Math.Round(Paragraphs[cont-1].Identation)+1: 1, Paragraphs[cont].IdentationCharacter));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(cont > 0 && !Paragraphs[cont - 1].ParagraphFormat.Equals(Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3) ? (int)Math.Round(Paragraphs[cont - 1].Identation) + 1 : 2, Paragraphs[cont].IdentationCharacter));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(cont > 0 && !Paragraphs[cont - 1].ParagraphFormat.Equals(Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4) ? (int)Math.Round(Paragraphs[cont - 1].Identation) + 2 : 3, Paragraphs[cont].IdentationCharacter));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel5:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(cont > 0 && !Paragraphs[cont - 1].ParagraphFormat.Equals(Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel5) ? (int)Math.Round(Paragraphs[cont - 1].Identation) + 3 : 4, Paragraphs[cont].IdentationCharacter));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel6:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(cont > 0 && !Paragraphs[cont - 1].ParagraphFormat.Equals(Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel6) ? (int)Math.Round(Paragraphs[cont - 1].Identation) + 4 : 4, Paragraphs[cont].IdentationCharacter));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel7:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(cont > 0 && !Paragraphs[cont - 1].ParagraphFormat.Equals(Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel7) ? (int)Math.Round(Paragraphs[cont - 1].Identation) + 5 : 4, Paragraphs[cont].IdentationCharacter));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel8:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(cont > 0 && !Paragraphs[cont - 1].ParagraphFormat.Equals(Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel8) ? (int)Math.Round(Paragraphs[cont - 1].Identation) + 6 : 4, Paragraphs[cont].IdentationCharacter));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel9:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(cont > 0 && !Paragraphs[cont - 1].ParagraphFormat.Equals(Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel9) ? (int)Math.Round(Paragraphs[cont - 1].Identation) + 7 : 4, Paragraphs[cont].IdentationCharacter));
                        break;
                    case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText:
                        textToReturn.Append(textToReturn.ToString().ReturnIteratedChars(cont > 0 && !Paragraphs[cont - 1].ParagraphFormat.Equals(Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText) ? (int)Math.Round(Paragraphs[cont-1].Identation)+1:5, Paragraphs[cont].IdentationCharacter));
                        break;
                    default:
                        break;
                }
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
