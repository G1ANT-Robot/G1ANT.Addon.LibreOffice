using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;
namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.insertcolumn", Tooltip = "Inserts a column before or after the specified column")]
    class CalcInsertColumnCommand : Command
    {
        public CalcInsertColumnCommand(AbstractScripter scripter) : base(scripter)
        {

        }
        public class Arguments : CommandArguments
        {
            [Argument(Name = "colnumber", Tooltip = "Enter the column number")]
            public IntegerStructure colnumber { get; set; } = new IntegerStructure();
            [Argument(Name = "before", Required = true, Tooltip = "Set to true to insert before the specified column number, false to insert after")]
            public BooleanStructure before { get; set; } = new BooleanStructure();
        }

        public void Execute(Arguments arguments)
        {
            CalcManager.CurrentCalc.InsertColumn(arguments.colnumber.Value, arguments.before.Value);
        }
    }
}
