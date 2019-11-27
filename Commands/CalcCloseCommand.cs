using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.close", Tooltip = "Close the Document with the Specified ID")]
    class CalcCloseCommand : Command
    {
        public CalcCloseCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "Id", Tooltip = "Enter the ID of the Document to close", Required = true)]
            public IntegerStructure Id { get; set; }
        }

        public void Execute(Arguments arguments)
        {
            CalcManager.Instance.RemoveInstance(arguments.Id.Value);
        }
    }
}
