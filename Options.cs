using System;
using System.Collections.Generic;
using System.Text;
using CommandLine;

namespace Docx_Ppt
{
    public class Options
    {
        [Option('m', "MarkDown", Required = true, HelpText = "MarkDown File Name")]
        public string MarkDownFile { get; set; }

        [Option('p', "PPT", Required = true, HelpText = "Powerpoint File Name")]
        public string PowerPointFile { get; set; }

    }
}
