using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.switch", Tooltip = "Switch to the specified instance of a calc spreadsheet")]
    class CalcSwitchCommand : Command
    {
        public CalcSwitchCommand(AbstractScripter scripter) : base(scripter)
        {

        }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "id", Tooltip = "Enter the ID of the sheet to switch to.")]
            public IntegerStructure id { get; set; } = new IntegerStructure();
        }

        public void Execute(Arguments arguments)
        {
            CalcManager.SwitchCalc(arguments.id.Value);
        }
    }
}
