using System;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.sheet;
using unoidl.com.sun.star.util;
using unoidl.com.sun.star.text;
using unoidl.com.sun.star.table;
using unoidl.com.sun.star.beans;

namespace G1ANT.Addon.LibreOffice
{
    public class CalcWrapper
    {
        public int Id { get; set; }
        public CalcWrapper(int id)
        {
            this.Id = id;
        }

        public XSpreadsheetDocument mxDocument;
        public XSpreadsheetDocument MxDocument
        {
            get { return mxDocument ?? throw new ApplicationException("Calc instance must be opened first using calc.open command"); }
            private set { mxDocument = value; }
        }

        private XSpreadsheetDocument InitDocument(bool hidden, string path)
        {
            var aLoader = CreateLoader();

            if (hidden)
            {
                var loadProperties = CreatePropertiesForHiddenReadonlyLoad();
                return IsNewDocument(path) ? CreateNewDocument(aLoader, loadProperties) : LoadDocument(path, aLoader, loadProperties);
            }
            else
            {
                return IsNewDocument(path) ? CreateNewDocument(aLoader) : LoadDocument(path, aLoader);
            }
        }

        private static bool IsNewDocument(string path)
        {
            return string.IsNullOrEmpty(path);
        }

        private static PropertyValue[] CreatePropertiesForHiddenReadonlyLoad()
        {
            return new PropertyValue[2] {
            new PropertyValue()
            {
            Name = "Hidden",
            Value = new uno.Any(true)
            },
            new PropertyValue()
            {
                Name = "ReadOnly",
                Value = new uno.Any(false)
            }
            };
        }

        private static XSpreadsheetDocument LoadDocument(string path, XComponentLoader aLoader, PropertyValue[] loadProperties = null)
        {
            path = string.Format("file:///{0}", path);
            return (XSpreadsheetDocument)aLoader.loadComponentFromURL(path, "_blank", 0, loadProperties ?? new PropertyValue[0]);
        }

        private static XSpreadsheetDocument CreateNewDocument(XComponentLoader aLoader, PropertyValue[] loadProperties = null)
        {
            return (XSpreadsheetDocument)aLoader.loadComponentFromURL("private:factory/scalc", "_blank", 0, loadProperties ?? new PropertyValue[0]);
        }

        private XComponentLoader CreateLoader()
        {
            var mxContext = uno.util.Bootstrap.bootstrap();
            var mxMSFactory = (XMultiServiceFactory)mxContext.getServiceManager();
            return (XComponentLoader)mxMSFactory.createInstance("com.sun.star.frame.Desktop");
        }

        private void SaveDocument(string path)
        {
            path = path.Replace("\\", "/"); // Convert forward slashes to backslashes, converting it to the correct format storeToURL expects. 
            path = string.Concat("file:///", path);
            XStorable xStorable = (XStorable)MxDocument; // Typecast the currently open document to XStorable type.
            xStorable.storeToURL(path, new PropertyValue[1]); //Creating an empty PropertyValue array saves the document in the default .ods format.
        }

        private void AddNewSheet(string name)
        {
            var xSheets = MxDocument.getSheets();
            short sheetIndex = Convert.ToInt16(xSheets.getElementNames().Length + 1); //Get the number of sheets already existing in the document, add one to get the new index. 
            xSheets.insertNewByName(name, sheetIndex);
        }

        public void RemoveSheet(string name)
        {
            XSpreadsheets xSheets = MxDocument.getSheets();
            xSheets.removeByName(name);
        }

        private void SetActiveSheet(string sheetName)
        {
            XSpreadsheets xSheets = MxDocument.getSheets();
            XModel xModel = (XModel)MxDocument;
            XController xController = xModel.getCurrentController();
            XSpreadsheetView sheetView = (XSpreadsheetView)xController;
            XSpreadsheet xSpreadsheet = (XSpreadsheet)xSheets.getByName(sheetName).Value;
            sheetView.setActiveSheet(xSpreadsheet);
        }

        private XSpreadsheet GetActiveSheet()
        {
            XModel xModel = (XModel)MxDocument;
            XController xController = xModel.getCurrentController();
            XSpreadsheetView sheetView = (XSpreadsheetView)xController;
            return sheetView.getActiveSheet();
        }

        public void Save(string path)
        {
            SaveDocument(path);
        }

        public int Open(bool hidden, string path)
        {
            MxDocument = InitDocument(hidden, path);
            return Id;
        }

        public void Close()
        {
            XModel xModel = (XModel)MxDocument;
            XCloseable xCloseable = (XCloseable)xModel;
            xCloseable.close(true);
        }

        public void AddSheet(string name)
        {
            AddNewSheet(name);
        }

        public void ActivateSheet(string sheetName)
        {
            SetActiveSheet(sheetName);
        }

        public string GetValue(int colNum, int rowNum)
        {
            var xSheet = GetActiveSheet();
            var xCell = (XText)xSheet.getCellByPosition(colNum - 1, rowNum - 1);
            string Value = xCell.getString();
            return Value;
        }

        public void SetValue(string value, int colNum, int rowNum)
        {
            var xSheet = GetActiveSheet();
            var xCell = (XText)xSheet.getCellByPosition(colNum - 1, rowNum -1 );
            XTextCursor xTextCursor = xCell.createTextCursor();
            xTextCursor.setString(value);
        }

        public void InsertRow(int rowNumber, bool before)
        {
            var xSheet = GetActiveSheet();
            XColumnRowRange xCRRange = (XColumnRowRange)xSheet;
            var xRows = xCRRange.getRows();
            if (before)
            {
                xRows.insertByIndex(rowNumber - 1, 1);
            }
            else
            {
                xRows.insertByIndex(rowNumber, 1);
            }            
        }

        public void InsertColumn(int colNumber, bool before)
        {
            var xSheet = GetActiveSheet();
            XColumnRowRange xCRRange = (XColumnRowRange)xSheet;
            var xColumns = xCRRange.getColumns();
            if (before)
            {
                xColumns.insertByIndex(colNumber - 1, 1);
            }
            else
            {
                xColumns.insertByIndex(colNumber, 1);
            }
        }

        public void RemoveRow(int rowNumber)
        {
            var xSheet = GetActiveSheet();
            XColumnRowRange xCRRange = (XColumnRowRange)xSheet;
            var xRows = xCRRange.getRows();
            xRows.removeByIndex(rowNumber - 1, 1);
        }

        public void RemoveColumn(int colNumber)
        {
            var xSheet = GetActiveSheet();
            XColumnRowRange xCRRange = (XColumnRowRange)xSheet;
            var xColumns = (XTableColumns)xCRRange.getColumns();
            xColumns.removeByIndex(colNumber - 1, 1);
        }
    }
}
