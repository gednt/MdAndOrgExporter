using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtensionMethods
{
    public static class Utilities
    {
        public static String ReturnIteratedChars(this String iterationCharacter, int timesToIterate,char character)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < timesToIterate; i++)
            {
                sb.Append(character);
            }
            return sb.ToString();
        } 
    }
}
