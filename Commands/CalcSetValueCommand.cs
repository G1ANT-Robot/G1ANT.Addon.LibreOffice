using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.setvalue", Tooltip = "Set the value of a specified cell")]
    class CalcSetValueCommand : Command
    {
        public CalcSetValueCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(Required = true, Name = "colnumber", Tooltip = "Enter the column number of the cell")]
            public IntegerStructure ColNum { get; set; } = new IntegerStructure();

            [Argument(Required = true, Name = "rownumber", Tooltip = "Enter the row number of the cell")]
            public IntegerStructure RowNum { get; set; } = new IntegerStructure();

            [Argument(Required = true, Name = "value", Tooltip = "Enter the value to insert into the cell")]
            public TextStructure value { get; set; } = new TextStructure();
        }
        
        public void Execute(Arguments arguments)
        {
            CalcManager.Instance.CurrentCalc.SetValue(arguments.value.Value, arguments.ColNum.Value, arguments.RowNum.Value);
        }
    }
}
