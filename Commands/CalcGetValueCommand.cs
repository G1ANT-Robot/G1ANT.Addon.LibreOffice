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
            public IntegerStructure colNum { get; set; } = new IntegerStructure();
            [Argument(Required = true, Name = "rownumber", Tooltip = "Enter the row number of the cell")]
            public IntegerStructure rowNum { get; set; } = new IntegerStructure();
            [Argument(Required = false, Name = "sheetname", Tooltip = "Enter the name of sheet, the default sheet is \"Sheet1\"")]
            public TextStructure sheetName { get; set; } = new TextStructure("Sheet1");

            [Argument(Tooltip = "Contains the value stored in the cell")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public void Execute(Arguments arguments)
        {
            var result = CalcManager.CurrentCalc.GetValue(arguments.sheetName.Value, arguments.colNum.Value, arguments.rowNum.Value);
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new TextStructure(result));
        }
            
    }
}
