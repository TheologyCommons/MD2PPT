using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Md_Ppt
{
    public class Paragraph:List<string>
    {
        public int Level()
        {
            if (this.Count == 2 && this[1].All(c => c == '='))
            {
                return 1;
            }

            if (this.Count == 2 && this[1].All(c => c == '-'))
            {
                return 2;
            }

            if (this.Any(l=>l.StartsWith('#')))
            {
                int level = 2;
                foreach (var c in this.First(l => l.StartsWith('#')))
                {
                    if (c != '#')
                    {
                        return level;
                    }
                    level++;
                }
            }
            return 0;
        }
    }
}
