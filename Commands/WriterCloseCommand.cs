using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "writer.close", Tooltip = "Close the Document with the Specified ID")]
    class WriterCloseCommand : Command
    {
        public WriterCloseCommand(AbstractScripter scripter) : base(scripter)
        {

        }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "ID", Tooltip = "Enter the ID of the Document to close", Required = true)]
            public IntegerStructure ID { get; set; } = new IntegerStructure();
        }

        public void Execute(Arguments arguments)
        {
            WriterManager.RemoveInstance(arguments.ID.Value);
        }
    }
}
