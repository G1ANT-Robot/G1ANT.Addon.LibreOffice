# G1ANT.Addon.LibreOffice
LibreOffice Addon for G1ANT Studio

## Commands

### Calc Commands:

- **calc.open:** Opens a new instance of Calc. 

  - **Parameters:** 
    - hidden: Whether to open the file in hidden mode
    - path: Path of an existing file to open, if left empty, creates a new document. Must contain file extension. 

- **calc.close**: Close the Document with the Specified ID

  - **Parameters:** 
    - id: ID of the Document to close

- **calc.save:** Saves the currently open Calc Spreadsheet

  - **Parameters:** 
    - savepath: Where to save the file. Must contain file extension. 

- **calc.switch:** Switch to the specified instance of a calc spreadsheet

  - **Parameters:**
    - id: ID of the sheet to switch to. 

- **calc.addsheet:** Add a new sheet with the name specified

  - **Parameters:** 
    - sheetname: Name of the Sheet to Add

- **calc.removesheet**: Removes the specified sheet

  - **Parameters:** 
    - sheetname: Name of the sheet to remove

- **calc.activatesheet:** Activate the Specified sheet

  - **Parameters:** 
    - sheet: Name of the sheet to activate

- **calc.getvalue:** Get the value of a specified cell

  - **Parameters:** 
    - colnumber: Column number of the cell
    - rownumber: Row number of the cell
    - sheetname: Name of the sheet to retrieve from, optional, default value is "Sheet1"

- **calc.setvalue:** Set the value of a specified cell

  - **Parameters:** 
    - colnumber: Column number of the cell
    - rownumber: Row number of the cell
    - sheetname: Name of the sheet to retrieve from, optional, default value is "Sheet1"
  - **calc.insertcolumn:** Insert a column before or after the specified column
    - **Parameters:** 
      - colnumber: Column number to insert before or after
      - before: Set to true to insert before the specified column number, false to insert after
  - **calc.insertrow:** Adds a new row before or after the specified row number
    - **Parameters:**
      - rownumber: Row number to insert before or after
      - before: Set to true to insert before the specified row number, false to insert after
  - **calc.removerow:** Removes a row with the specified row number
    - **Parameters:**
      - rownumber: Row number to delete
  - **calc.removecolumn**: Removes a column with the specified column number
    - **Parameters:** 
      - colnumber: Column number to delete

  **Writer Commands:**

  - **writer.open:** Creates a new instance of writer
    - **Parameters:** 
      - hidden: Whether to open the file in hidden mode
      - path: Path of an existing file to open, if left empty, creates a new document. Must contain file extension. 

  - **writer.close:** Close the document with the specified ID
    - **Parameters:** 
      - id: ID of the document to close
  - **writer.save:** Saves the currently open document
    - **Parameters:**
      - savepath: Where to save the file, requires file extension.
  - **writer.switch**: Switch to the specified instance of a writer document
    - **Parameters:**
      - id: ID of the document to switch to
  - **writer.inserttext:** Inserts text into the currently open Writer document
    - **Parameters:**
      - text: Text to insert
      - append: True if you want to append the text, false to replace text 
  - **writer.gettext:** Get the text in the document
  - **writer.replace:** Replaces text in the document
    - **Parameters:** 
      - word: Word to replace
      - replacewith: Word to replace with

  **NOTE:** For the addon to work correctly, File Locking in LibreOffice must be disabled. Click [here](https://ask.libreoffice.org/en/question/135501/how-do-i-prevent-these-lock-files-from-being-createddeleted/?answer=135503#post-id-135503) for instructions on how to disable locking. 