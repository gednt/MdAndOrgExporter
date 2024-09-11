using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MdAndOrgExporter.Models
{
    public class Paragraph
    {
        public readonly string IdentationCharacter = "*";
        public readonly string IdentationCharacter_md = "-";
        public enum MarkupType
        {
            NONE = 0,
            BOLD = 1,
            ITALIC = 2
        };
        public float Identation { get; set; }
        public string Text { get; set; }

        public ListFormat ListFormat { get; set; }
        public List List { get; set; }
        public MarkupType MarkupTypeApplied { get; set; }

        public bool ContainsImage { get; set; } = false;

    }
}
