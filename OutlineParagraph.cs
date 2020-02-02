using Md_Ppt;
using System.Collections.Generic;
using System.Linq;

namespace Md_Ppt
{
    internal class OutlineParagraph
    {
        public OutlineParagraph(Paragraph paragraph,List<Paragraph> children)
        { 
            Text = paragraph?.First();
            Level = paragraph?.Level()??0;
            Children = new List<OutlineParagraph>();
            List<Paragraph> grandchildren = new List<Paragraph>();

            Paragraph workingPara = null;


            foreach (var para in children)
            {

                if (para.Level() > 0)
                {
                    if (workingPara == null)
                    {
                        workingPara = para;
                    }
                    else if (para.Level() <= workingPara.Level())
                    {
                        Children.Add(new OutlineParagraph(workingPara,grandchildren));
                        grandchildren = new List<Paragraph>();
                        workingPara = para;
                    }
                    else
                    {
                        grandchildren.Add(para);
                    }
                }
            }

            if (workingPara != null)
            {
                Children.Add(new OutlineParagraph(workingPara, grandchildren));
            }
        }

        public string Text { get; }

        public int Level { get; }

        public List<OutlineParagraph> Children { get; }

    }
}