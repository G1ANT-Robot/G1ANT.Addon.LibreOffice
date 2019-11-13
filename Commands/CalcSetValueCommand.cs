using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.setvalue", Tooltip = "Set the value of a specified cell")]
    class CalcSetValueCommand : Command
    {
        public CalcSetValueCommand(AbstractScripter scripter) : base(scripter)
        {

        }

        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Name = "colnumber", Tooltip = "Enter the column number of the cell")]
            public IntegerStructure colNum { get; set; } = new IntegerStructure();
            [Argument(Required = true, Name = "rownumber", Tooltip = "Enter the row number of the cell")]
            public IntegerStructure rowNum { get; set; } = new IntegerStructure();
            [Argument(Required = true, Name = "value", Tooltip = "Enter the value to insert into the cell")]
            public TextStructure value { get; set; } = new TextStructure();
            [Argument(Required = false, Name = "sheetname", Tooltip = "Enter the name of sheet, the default sheet is \"Sheet1\"")]
            public TextStructure sheetName { get; set; } = new TextStructure("Sheet1");
        }
        
        public void Execute(Arguments arguments)
        {
            CalcManager.CurrentCalc.SetValue(arguments.value.Value, arguments.sheetName.Value, arguments.colNum.Value, arguments.rowNum.Value);
        }
    }
}
