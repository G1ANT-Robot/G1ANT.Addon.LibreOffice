using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice.Commands
{
    [Command(Name = "calc.removesheet", Tooltip = "Removes the specified sheet")]
    class CalcRemoveSheetCommand : Command
    {
        public CalcRemoveSheetCommand(AbstractScripter scripter) : base(scripter) { }

        public class Arguments : CommandArguments
        {
            [Argument(Name = "sheetname", Tooltip = "Enter the name of the sheet to remove.", Required = true)]
            public TextStructure SheetName { get; set; } = new TextStructure();
        }

        public void Execute(Arguments arguments)
        {
            CalcManager.Instance.CurrentCalc.RemoveSheet(arguments.SheetName.Value);
        }
    }
}
