using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MdAndOrgExporter.Functions
{
    public class Paragraph : Models.Paragraph
    {
        public String Markup(String text_portion, MarkupType type)
        {
            switch (type)
            {
                case MarkupType.BOLD: 
                    return "*"+text_portion+"*";
                case MarkupType.ITALIC:
                    return "/"+text_portion+"/";
                default:
                case MarkupType.NONE:
                    return text_portion;


            }
        }

        public int Type()
        {
            if(this.Style.NameLocal.Contains("Heading 1"))
            {
                return 1;
            }
            if(this.Style.NameLocal.Contains("Heading 2"))
            {
                return 2;
            }
            if(this.Style.NameLocal.Contains("Heading 3"))
            {
                return 3;
            }
            if(this.Style.NameLocal.Contains("Heading 4"))
            {
                return 4;
            }
            if (this.Style.NameLocal.Contains("Heading 5"))
            {
                return 5;
            }
            if (this.Style.NameLocal.Contains("Heading 6"))
            {
                return 6;
            }
            if (this.Style.NameLocal.Contains("Heading 7"))
            {
                return 7;
            }
            if (this.Style.NameLocal.Contains("Heading 8"))
            {
                return 8;
            }
            if (this.Style.NameLocal.Contains("Heading 9"))
            {
                return 9;
            }
            return 0;
        }
    }
}
