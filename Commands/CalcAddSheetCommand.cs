using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.addsheet", Tooltip = "Add a new sheet with the name specified")]
    class CalcAddSheetCommand : Command
    {
        public CalcAddSheetCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "sheetname", Tooltip = "Name of the Sheet to Add", Required = true)]
            public TextStructure SheetName { get; set; } = new TextStructure();
        }

        public void Execute(Arguments arguments)
        {
            CalcManager.Instance.CurrentCalc.AddSheet(arguments.SheetName.Value);
        }
    }
}
