using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.switch", Tooltip = "Switch to the specified instance of a calc spreadsheet")]
    class CalcSwitchCommand : Command
    {
        public CalcSwitchCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "Id", Tooltip = "Enter the ID of the sheet to switch to.")]
            public IntegerStructure Id { get; set; }
        }

        public void Execute(Arguments arguments)
        {
           CalcManager.Instance.SwitchCalc(arguments.Id.Value);
        }
    }
}
