using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "writer.switch", Tooltip = "Switch to the specified instance of a writer document")]
    class WriterSwitchCommand : Command
    {
        public WriterSwitchCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "Id", Tooltip = "Enter the ID of the document to switch to", Required = true)]
            public IntegerStructure Id { get; set; }
        }
        
        public void Execute(Arguments arguments)
        {
            WriterManager.Instance.SwitchWriter(arguments.Id.Value);
        }
    }
}
