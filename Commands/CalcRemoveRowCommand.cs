using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;
namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.removerow", Tooltip = "Removes a row with the specified row number")]
    class CalcRemoveRowCommand : Command
    {
        public CalcRemoveRowCommand(AbstractScripter scripter) : base(scripter)
        {

        }
        public class Arguments : CommandArguments
        {
            [Argument(Name = "rownumber", Required = true, Tooltip = "Enter the row number")]
            public IntegerStructure RowNumber { get; set; } = new IntegerStructure();
        }
        public void Execute(Arguments arguments)
        {
            CalcManager.CurrentCalc.RemoveRow(arguments.RowNumber.Value);
        }
    }
}
