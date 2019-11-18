using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.getvalue", Tooltip = "Get the value of a specified cell")]
    class CalcGetValueCommand : Command
    {
        public CalcGetValueCommand(AbstractScripter scripter) : base(scripter)
        {

        }

        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Name = "colnumber", Tooltip = "Enter the column number of the cell")]
            public IntegerStructure ColNum { get; set; } = new IntegerStructure();
            [Argument(Required = true, Name = "rownumber", Tooltip = "Enter the row number of the cell")]
            public IntegerStructure RowNum { get; set; } = new IntegerStructure();
            [Argument(Tooltip = "Contains the value stored in the cell")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public void Execute(Arguments arguments)
        {
            var result = CalcManager.CurrentCalc.GetValue(arguments.ColNum.Value, arguments.RowNum.Value);
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new TextStructure(result));
        }
            
    }
}
