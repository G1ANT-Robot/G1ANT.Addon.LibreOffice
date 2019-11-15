using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "writer.inserttext", Tooltip = "Inserts text into the currently open Writer document")]
    class WriterInsertTextCommand : Command
    {
        public WriterInsertTextCommand(AbstractScripter scripter) : base(scripter)
        {

        }
        public class Arguments : CommandArguments
        {
            [Argument(Name = "text", Tooltip = "Enter the text to be insert", Required = true)]
            public TextStructure text { get; set; } = new TextStructure();
            [Argument(Name = "append", Tooltip = "True if you want to append the text, false to replace text")]
            public BooleanStructure append { get; set; } = new BooleanStructure();
        }

        public void Execute(Arguments arguments)
        {
            WriterManager.CurrentWriter.InsertText(arguments.text.Value, arguments.append.Value);
        }
    }
}
