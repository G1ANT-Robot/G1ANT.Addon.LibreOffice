using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.addsheet", Tooltip = "Add a new sheet with the name specified")]
    class CalcAddSheetCommand : Command
    {
        public CalcAddSheetCommand(AbstractScripter scripter) : base(scripter)
        {

        }

        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Name of the Sheet to Add", Required = true)]

            public TextStructure sheetName { get; set; }
        }

        public void Execute(Arguments arguments)
        {

            CalcManager.CurrentCalc.AddSheet(arguments.sheetName.Value);
        }
    }
}
