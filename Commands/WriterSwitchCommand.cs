using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "writer.switch", Tooltip = "Switch to the specified instance of a writer document")]
    class WriterSwitchCommand : Command
    {
        public WriterSwitchCommand(AbstractScripter scripter) : base(scripter)
        {

        }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "id", Tooltip = "Enter the ID of the document to switch to", Required = true)]
            public IntegerStructure id { get; set; } = new IntegerStructure();
        }
        
        public void Execute(Arguments arguments)
        {
            WriterManager.SwitchWriter(arguments.id.Value);
        }
    }
}
