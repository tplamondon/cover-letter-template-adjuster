using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;

namespace cover_letter_template_adjuster
{
    public class Options
    {
        [Option('p', "path", Required = true, HelpText = "The file path of the word document")]
        public required string Path { get; set; }

        [Option('c', "company", Required = true, HelpText = "The name of the company")]
        public required string CompanyName { get; set; }

        [Option('r', "role", Required = true, HelpText = "The name of the role")]
        public required string RoleName { get; set; }
    }
}
