using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "writer.replace", Tooltip = "Replaces text in the document")]
    class WriterReplaceCommand : Command
    {
        public WriterReplaceCommand(AbstractScripter scripter) : base(scripter)
        {

        }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "word", Tooltip = "Enter the word to replace", Required = true)]
            public TextStructure word { get; set; } = new TextStructure();
            [Argument(Name = "replacewith", Tooltip = "Enter the word to replace with", Required = true)]
            public TextStructure ReplaceWith { get; set; } = new TextStructure();
        }

        public void Execute(Arguments arguments)
        {
            WriterManager.CurrentWriter.ReplaceWith(arguments.word.Value, arguments.ReplaceWith.Value);
        }
    }
}
