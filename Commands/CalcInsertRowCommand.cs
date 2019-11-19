using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.insertrow", Tooltip = "Adds a new row before or after the specified row number")]
    class CalcInsertRowCommand : Command
    {
        public CalcInsertRowCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "rownumber", Required = true, Tooltip = "Enter the row number")]
            public IntegerStructure RowNumber { get; set; } = new IntegerStructure();

            [Argument(Name = "before", Required = true, Tooltip = "Set to true to insert before the specified row number, false to insert after")]
            public BooleanStructure before { get; set; } = new BooleanStructure(true);
        }

        public void Execute(Arguments arguments)
        {
            CalcManager.Instance.CurrentCalc.InsertRow(arguments.RowNumber.Value, arguments.before.Value);
        }
    }
}
