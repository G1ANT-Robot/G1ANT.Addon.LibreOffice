using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice
{
    [Command(Name = "calc.open", Tooltip = "This command creates a new instance of Calc")]
    public class CalcOpenCommand : Command
    {
        public CalcOpenCommand(AbstractScripter scripter): base(scripter)
        {

        }
        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Contains ID of the opened Instance, can be used with calc.switch")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }

        public void Execute(Arguments arguments)
        {
            CalcWrapper calcWrapper = CalcManager.CreateInstance();
            var id = calcWrapper.Open();
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.IntegerStructure(calcWrapper.Id));
        }
    }
}
