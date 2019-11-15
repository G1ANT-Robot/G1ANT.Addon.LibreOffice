using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;
namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "writer.open", Tooltip = "This command creates a new instance of Writer")]
    class WriterOpenCommand : Command
    {
        public WriterOpenCommand(AbstractScripter scripter) : base(scripter)
        {

        }

        public class Arguments : CommandArguments
        {
            [Argument(Tooltip = "Whether to open the file in hidden mode", Required = true)]
            public BooleanStructure Hidden { get; set; } = new BooleanStructure(false);

            [Argument(Tooltip = "Path of an existing file to open, if left empty, creates a new document", Required = false)]
            public TextStructure path { get; set; } = new TextStructure();

            [Argument(Tooltip = "Contains ID of the opened Instance, can be used with writer.switch")]
            public VariableStructure Result { get; set; } = new VariableStructure("result");
        }
        public void Execute(Arguments arguments)
        {
            WriterWrapper writerWrapper = WriterManager.CreateInstance();
            WriterManager.CurrentWriter.Open(arguments.Hidden.Value, arguments.path.Value);
            Scripter.Variables.SetVariableValue(arguments.Result.Value, new Language.IntegerStructure(writerWrapper.Id));
        }
    }
}
