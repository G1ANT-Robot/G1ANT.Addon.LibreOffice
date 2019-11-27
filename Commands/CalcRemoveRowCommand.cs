using G1ANT.Language;
namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.removerow", Tooltip = "Removes a row with the specified row number")]
    class CalcRemoveRowCommand : Command
    {
        public CalcRemoveRowCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "rownumber", Required = true, Tooltip = "Enter the row number")]
            public IntegerStructure RowNumber { get; set; }
        }
        public void Execute(Arguments arguments)
        {
            CalcManager.Instance.CurrentCalc.RemoveRow(arguments.RowNumber.Value);
        }
    }
}
