using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;
namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.removecolumn", Tooltip = "Removes a column with the specified column number")]
    class CalcRemoveColumnCommand : Command
    {
        public CalcRemoveColumnCommand(AbstractScripter scripter) : base(scripter)
        {

        }
        public class Arguments : CommandArguments
        {
            [Argument(Name = "colnumber", Required = true, Tooltip = "Enter the column number")]
            public IntegerStructure colNumber { get; set; } = new IntegerStructure();
        }
        public void Execute(Arguments arguments)
        {
            CalcManager.CurrentCalc.RemoveRow(arguments.colNumber.Value);
        }
    }
}
