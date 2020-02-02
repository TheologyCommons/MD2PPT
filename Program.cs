using System;
using CommandLine;
using Md_Ppt;

namespace Docx_Ppt
{
    class Program
    {
        static void Main(string[] args)
        {
           
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(o => { PPTGenerator.Create(o.MarkDownFile,o.PowerPointFile); }
            );
        }
    }
}
