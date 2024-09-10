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
    }
}
