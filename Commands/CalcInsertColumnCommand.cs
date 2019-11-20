using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.insertcolumn", Tooltip = "Inserts a column before or after the specified column")]
    class CalcInsertColumnCommand : Command
    {
        public CalcInsertColumnCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "colindex", Tooltip = "Enter the column number", Required = true)]
            public IntegerStructure ColNumber { get; set; }

            [Argument(Name = "before", Required = true, Tooltip = "Set to true to insert before the specified column number, false to insert after")]
            public BooleanStructure before { get; set; } = new BooleanStructure(false);
        }

        public void Execute(Arguments arguments)
        {
            CalcManager.Instance.CurrentCalc.InsertColumn(arguments.ColNumber.Value, arguments.before.Value);
        }
    }
}
