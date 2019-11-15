using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using unoidl.com.sun.star.bridge;
using unoidl.com.sun.star.frame;
using G1ANT.Language;

namespace G1ANT.Addon.LibreOffice
{
    public class CalcWrapper
    {
        public int Id { get; set; }

        public CalcWrapper(int id)
        {
            this.Id = id;
        }

        private unoidl.com.sun.star.uno.XComponentContext m_xContext;
        private unoidl.com.sun.star.lang.XMultiServiceFactory mxMSFactory;
        public unoidl.com.sun.star.sheet.XSpreadsheetDocument mxDocument;

        private unoidl.com.sun.star.sheet.XSpreadsheetDocument InitDocument(bool Hidden, String Path)
        {
            mxMSFactory = Connect();
            XComponentLoader aLoader = (XComponentLoader)mxMSFactory.createInstance("com.sun.star.frame.Desktop");
            XComponent xComponent;
            unoidl.com.sun.star.beans.PropertyValue[] loadProperties = new unoidl.com.sun.star.beans.PropertyValue[2];
            loadProperties[0] = new unoidl.com.sun.star.beans.PropertyValue();
            loadProperties[1] = new unoidl.com.sun.star.beans.PropertyValue();
            loadProperties[1].Name = "ReadOnly";
            loadProperties[1].Value = new uno.Any(false);

            if (Hidden)
            {
                loadProperties[0].Name = "Hidden";
                loadProperties[0].Value = new uno.Any(true);

                //Check to see if the path is provided by the user, if it's empty, create a new document. 
                if (String.IsNullOrEmpty(Path))
                {
                    xComponent = aLoader.loadComponentFromURL("private:factory/scalc", "_blank", 0, loadProperties);
                }
                else
                {
                    try
                    {
                        Path = String.Format("file:///{0}", Path);
                        xComponent = aLoader.loadComponentFromURL(Path, "_blank", 0, loadProperties);
                    }
                    catch (unoidl.com.sun.star.uno.Exception ex)
                    {
                        throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
                    }
                }
            }
            else
            {
                loadProperties[0].Name = "Hidden";
                loadProperties[0].Value = new uno.Any(false);
                if (String.IsNullOrEmpty(Path))
                {
                    xComponent = aLoader.loadComponentFromURL("private:factory/scalc", "_blank", 0, new unoidl.com.sun.star.beans.PropertyValue[0]);
                }
                else
                {
                    try
                    {
                        Path = String.Format("file:///{0}", Path);
                        xComponent = aLoader.loadComponentFromURL(Path, "_blank", 0, new unoidl.com.sun.star.beans.PropertyValue[0]);
                    }
                    catch (unoidl.com.sun.star.uno.Exception ex)
                    {
                        throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
                    }
                }
            }

            return (unoidl.com.sun.star.sheet.XSpreadsheetDocument)xComponent;
        }

        private void SaveDocument(String Path)
        {
            try
            {
                Path = Path.Replace("\\", "/"); // Convert forward slashes to backslashes, converting it to the correct format storeToURL expects. 
                Path = String.Concat("file:///", Path);
                XStorable xStorable = (XStorable)mxDocument; // Typecast the currently open document to XStorable type.
                xStorable.storeToURL(Path, new unoidl.com.sun.star.beans.PropertyValue[1]); //Creating an empty PropertyValue array saves the document in the default .ods format.
            }
            catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }

        }

        private void AddNewSheet(String Name)
        {
            try
            {
                unoidl.com.sun.star.sheet.XSpreadsheets xSheets = mxDocument.getSheets();
                short sheetIndex = Convert.ToInt16(xSheets.getElementNames().Length + 1); //Get the number of sheets already existing in the document, add one to get the new index. 
                xSheets.insertNewByName(Name, sheetIndex);
            }

            catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }
        }

        public void RemoveSheet(String Name)
        {
            try
            {
                unoidl.com.sun.star.sheet.XSpreadsheets xSheets = mxDocument.getSheets();
                xSheets.removeByName(Name);
            }
            catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }
        }


        private void SetActiveSheet(String sheetName)
        {
            try
            {
                unoidl.com.sun.star.sheet.XSpreadsheets xSheets = mxDocument.getSheets();
                unoidl.com.sun.star.frame.XModel xModel = (unoidl.com.sun.star.frame.XModel)mxDocument;
                unoidl.com.sun.star.frame.XController xController = xModel.getCurrentController();
                unoidl.com.sun.star.sheet.XSpreadsheetView sheetView = (unoidl.com.sun.star.sheet.XSpreadsheetView)xController;
                unoidl.com.sun.star.sheet.XSpreadsheet xSpreadsheet = (unoidl.com.sun.star.sheet.XSpreadsheet)xSheets.getByName(sheetName).Value;
                sheetView.setActiveSheet(xSpreadsheet);
            }
           catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }
        }
        private unoidl.com.sun.star.sheet.XSpreadsheet GetActiveSheet()
        {
            unoidl.com.sun.star.frame.XModel xModel = (unoidl.com.sun.star.frame.XModel)mxDocument;
            unoidl.com.sun.star.frame.XController xController = xModel.getCurrentController();
            unoidl.com.sun.star.sheet.XSpreadsheetView sheetView = (unoidl.com.sun.star.sheet.XSpreadsheetView)xController;
            return sheetView.getActiveSheet();
        }

        private XMultiServiceFactory Connect()
        {
            m_xContext = uno.util.Bootstrap.bootstrap();
            return (XMultiServiceFactory)m_xContext.getServiceManager();
        }

        public int Open(bool Hidden, String Path)
        {
            mxDocument = InitDocument(Hidden, Path);
            return this.Id;
        }

        public void Save(String Path)
        {
            SaveDocument(Path);
        }

        public void Close()
        {
            try
            {
                XModel xModel = (XModel)mxDocument;
                unoidl.com.sun.star.util.XCloseable xCloseable = (unoidl.com.sun.star.util.XCloseable)xModel;
                xCloseable.close(true);
            }
            catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }
        }

        public void AddSheet(String Name)
        {
            AddNewSheet(Name);
        }

        public void ActivateSheet(String sheetName)
        {
            SetActiveSheet(sheetName);
        }

        public String GetValue(String sheetName, int colNum, int rowNum)
        {
            try
            {
                var xSheets = mxDocument.getSheets();
                var xSheet = (unoidl.com.sun.star.sheet.XSpreadsheet)xSheets.getByName(sheetName).Value;
                var xCell = (unoidl.com.sun.star.text.XText)xSheet.getCellByPosition(colNum - 1, rowNum - 1);
                String Value = xCell.getString();
                return Value;
            }
            catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }
        }

        public void SetValue(String Value, String sheetName, int colNum, int rowNum)
        {
            try
            {
                var xSheets = mxDocument.getSheets();
                var xSheet = (unoidl.com.sun.star.sheet.XSpreadsheet)xSheets.getByName(sheetName).Value;
                var xCell = (unoidl.com.sun.star.text.XText)xSheet.getCellByPosition(colNum - 1, rowNum -1 );
                unoidl.com.sun.star.text.XTextCursor xTextCursor = xCell.createTextCursor();
                xTextCursor.setString(Value);
            }
            catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }
        }

        public void InsertRow(int rowNumber, bool before)
        {
            try
            {
                var xSheet = GetActiveSheet();
                unoidl.com.sun.star.table.XColumnRowRange xCRRange = (unoidl.com.sun.star.table.XColumnRowRange)xSheet;
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
            catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }
        }

        public void InsertColumn(int colNumber, bool before)
        {
            try
            {
                var xSheet = GetActiveSheet();
                unoidl.com.sun.star.table.XColumnRowRange xCRRange = (unoidl.com.sun.star.table.XColumnRowRange)xSheet;
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
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);
            }
        }

        public void RemoveRow(int rowNumber)
        {
            try
            {
                var xSheet = GetActiveSheet();
                unoidl.com.sun.star.table.XColumnRowRange xCRRange = (unoidl.com.sun.star.table.XColumnRowRange)xSheet;
                var xRows = xCRRange.getRows();
                xRows.removeByIndex(rowNumber - 1, 1);
            }
            catch(unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);

            }
        }

        public void RemoveColumn(int colNumber)
        {
            try
            {
                var xSheet = GetActiveSheet();
                unoidl.com.sun.star.table.XColumnRowRange xCRRange = (unoidl.com.sun.star.table.XColumnRowRange)xSheet;
                var xColumns = (unoidl.com.sun.star.table.XTableColumns)xCRRange.getColumns();
                xColumns.removeByIndex(colNumber - 1, 1);
            }
            catch (unoidl.com.sun.star.uno.Exception ex)
            {
                throw new unoidl.com.sun.star.uno.Exception(ex.Message, ex);

            }
        }
    }
}
