using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.activatesheet", Tooltip = "Activate the specified sheet")]
    class CalcActivateSheetCommand : Command
    {
       public CalcActivateSheetCommand(AbstractScripter scripter) : base(scripter)
        {

        }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "sheet", Tooltip = "Enter the Name of the sheet", Required = true)]
            public TextStructure SheetName { get; set; } = new TextStructure();
        }

        public void Execute(Arguments arguments)
        {
            CalcManager.CurrentCalc.ActivateSheet(arguments.SheetName.Value);
        }
    }
}
