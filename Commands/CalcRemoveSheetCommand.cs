using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.removesheet", Tooltip = "Removes the specified sheet")]
    class CalcRemoveSheetCommand : Command
    {
        public CalcRemoveSheetCommand(AbstractScripter scripter) : base(scripter)
        {

        }
        public class Arguments : CommandArguments
        {
            [Argument(Name = "sheetname", Tooltip = "Enter the name of the sheet to remove.", Required = true)]
            public TextStructure sheetName { get; set; } = new TextStructure();
        }

        public void Execute(Arguments arguments)
        {
            CalcManager.CurrentCalc.RemoveSheet(arguments.sheetName.Value);
        }
    }
}
