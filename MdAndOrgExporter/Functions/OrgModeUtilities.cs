using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                textToReturn.Append(paragraph.IdentationCharacter);
                for(int i = 0;i<Math.Round(paragraph.Identation,0);i++)
                {
                    textToReturn.Append(paragraph.IdentationCharacter);
                }
                if (paragraph.ContainsImage == false)
                {
                    textToReturn.Append(" " + (paragraph.ListFormat != null ? paragraph.ListFormat.ListString + " " : "") + paragraph.Text + System.Environment.NewLine);
                }
                else
                {
                    textToReturn.Append(" " + (paragraph.ListFormat != null ? paragraph.ListFormat.ListString + " " : "") + $"![{ Path.GetFileName(paragraph.Text)}](../assets/{paragraph.Text})" + System.Environment.NewLine);
                }

            }
            return textToReturn.ToString();
        }
    }
}
