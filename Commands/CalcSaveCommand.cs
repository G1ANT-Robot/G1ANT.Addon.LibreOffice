using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.save", Tooltip = "Command to save the currently open Calc spreadsheet")]
    class CalcSaveCommand : Command
    {
        public CalcSaveCommand(AbstractScripter scripter) : base(scripter)
        {

        }

        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Where to save the file", Required = true)]
            public TextStructure savepath { get; set; } = new TextStructure();
        }

        public void Execute(Arguments arguments)
        { 
            CalcManager.CurrentCalc.Save(arguments.savepath.Value);
        }
    }
}
