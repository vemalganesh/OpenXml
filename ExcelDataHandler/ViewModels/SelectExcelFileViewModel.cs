using Caliburn.Micro;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataHandler.Models;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelDataHandler.ViewModels
{
    class SelectExcelFileViewModel: Screen
    {
        private string _sheetName;
        private SpreadSheetHelper _mainSheetData;
        private SpreadSheetHelper _sheetData;
        private List<InformationSpreadSheet> _mainSheetObjectToCompare;
        private List<InformationSpreadSheet> _newSheetObject;
        private string _filePath;
        public SelectExcelFileViewModel()
        {
           
        }

        public void ClearText()
        {
            _filePath = @"C:\Users\vemal\Documents\DevHandover_corrected.xlsm";
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_filePath, true))
            {

                _sheetName = "Sheet1";  //Sheet name should be changed to dynamic
                _mainSheetData = GetSingleSheetByName(spreadsheetDocument, _sheetName);
                _mainSheetData = GetRowsAndStringTableFromSheet(_mainSheetData);
                _mainSheetObjectToCompare = GetAllNamesFromSheet(_mainSheetData.sharedStringTable, _mainSheetData.Rows);

                _sheetName = "EBAttributes";
                _sheetData = GetSingleSheetByName(spreadsheetDocument, _sheetName);
                _sheetData = GetRowsAndStringTableFromSheet(_sheetData);
                _newSheetObject = GetAllSimilerData(_mainSheetObjectToCompare, _sheetData.sharedStringTable, _sheetData.Rows);

                _sheetName = "API";
                _sheetData = GetSingleSheetByName(spreadsheetDocument, _sheetName);
                _sheetData = GetRowsAndStringTableFromSheet(_sheetData);
                _newSheetObject = GetAllSimilerData(_mainSheetObjectToCompare, _sheetData.sharedStringTable, _sheetData.Rows);

                _sheetName = "ISA";
                _sheetData = GetSingleSheetByName(spreadsheetDocument, _sheetName);
                _sheetData = GetRowsAndStringTableFromSheet(_sheetData);
                _newSheetObject = GetAllSimilerData(_mainSheetObjectToCompare, _sheetData.sharedStringTable, _sheetData.Rows);

                //DeleteSheet(_mainSheetData);
               var newSheet = CreateNewWorkSheet(spreadsheetDocument, "Sheet2");
               spreadsheetDocument.Save();
               AddrowsToNewWorkSheet(_newSheetObject, spreadsheetDocument);
            }
        }

        private SpreadSheetHelper GetRowsAndStringTableFromSheet(SpreadSheetHelper spreadSheet)
        {

            spreadSheet.SharedStringTablePart = spreadSheet.workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            spreadSheet.sharedStringTable = spreadSheet.SharedStringTablePart.SharedStringTable;
            //List<Cell> cells = sheetData.workSheet.Descendants<Cell>().ToList();
            spreadSheet.Rows = spreadSheet.workSheet.Descendants<Row>().ToList();
            return spreadSheet;

        }

        public void DeleteSheet(SpreadSheetHelper spreadSheet)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_filePath, false))
            {
                var workbookPart = spreadsheetDocument.WorkbookPart;

                // Get the SheetToDelete from workbook.xml
                var theSheet = workbookPart.Workbook.Descendants<Sheet>()
                                           .FirstOrDefault(s => s.Id == spreadSheet.SheetId);

                if (theSheet == null)
                {
                    return;
                }

                // Remove the sheet reference from the workbook.
                var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(spreadSheet.SheetId));
                theSheet.Remove();

                // Delete the worksheet part.
                workbookPart.DeletePart(worksheetPart);
            }
        }

        private void AddrowsToNewWorkSheet(List<InformationSpreadSheet> sheetObject, SpreadsheetDocument spreadsheetDocument)
        {

            IEnumerable<Sheet> Sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == "Sheet29");
            string relationshipId = Sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            //First Time Row Creation for Header
            int rowCount = 1;
            Row HeaderRow = new Row();
            HeaderRow.Append(CreateNewCell("A", rowCount, "Attributes with Lower Case"));
            HeaderRow.Append(CreateNewCell("B", rowCount, "EBAttributes"));
            HeaderRow.Append(CreateNewCell("C", rowCount, "API"));
            HeaderRow.Append(CreateNewCell("D", rowCount, "ISA"));
            sheetData.Append(HeaderRow);
            foreach (var obj in sheetObject)
            {
                rowCount++;
                Row row = new Row();
                row.Append(CreateNewCell("A", rowCount, obj.AttributesLCNames));
                row.Append(CreateNewCell("B", rowCount, obj.EBAttributesId));
                row.Append(CreateNewCell("C", rowCount, obj.APIId));
                row.Append(CreateNewCell("D", rowCount, obj.ISAID));
                sheetData.Append(row);
            }

        }

        private SpreadSheetHelper GetSingleSheetByName(SpreadsheetDocument document, string sheetName)
        {
            SpreadSheetHelper helper = new SpreadSheetHelper();
            IEnumerable<Sheet> Sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            helper.SheetId = Sheets.First().Id.Value;
            helper.workbookPart = document.WorkbookPart;

            helper.WorkSheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(helper.SheetId);
            helper.workSheet = helper.WorkSheetPart.Worksheet;

            return helper;
        }
        private Sheet CreateNewWorkSheet(SpreadsheetDocument spreadSheet, string NewSheetName)
        {
            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Give the new worksheet a name.
            string sheetName = NewSheetName + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            newWorksheetPart.Worksheet.Save();
            spreadSheet.WorkbookPart.Workbook.Save();
            return sheet;
        }

        private Cell CreateNewCell(string columnName, int rowNumber, string cellValue)
        {
            Cell cell = new Cell()
            {
                CellReference = columnName + rowNumber,
                CellValue = new CellValue(cellValue),
            };
            

            return cell;
        }

        private List<InformationSpreadSheet> GetAllSimilerData(List<InformationSpreadSheet> newSpreadSheetData, SharedStringTable sstforsheet,List<Row> SheetRows)
        {

            foreach (Row row in SheetRows)
            {
                Cell AttributeName = new Cell();
                if (_sheetName.Equals("EBAttributes"))
                    AttributeName = row.Descendants<Cell>().Where(x => x.CellReference == "A" + row.RowIndex).FirstOrDefault();
                else
                    AttributeName = row.Descendants<Cell>().Where(x => x.CellReference == "C" + row.RowIndex).FirstOrDefault();

                if (AttributeName != null && (AttributeName.DataType != null) && (AttributeName.DataType == CellValues.SharedString))
                    {
                        int ssid = int.Parse(AttributeName.CellValue.Text);
                        string TextInsideTheCell = sstforsheet.ChildElements[ssid].InnerText;

                        if (newSpreadSheetData.Where(s => s.AttributesLCNames.Equals(TextInsideTheCell)).FirstOrDefault() != null && !String.IsNullOrEmpty(TextInsideTheCell) && TextInsideTheCell != " ")
                        {
                            Cell cellvalue = new Cell();

                            //This is to chose the cell to be taken from which column so that we dont have to manually change the placement in excel
                            if (_sheetName.Equals("EBAttributes"))// This Should be dynamic
                            {
                                cellvalue = row.Descendants<Cell>().Where(x => x.CellReference == "B" + row.RowIndex).FirstOrDefault();
                                if (cellvalue != null)
                                {
                                    newSpreadSheetData.Single(s => s.AttributesLCNames.Equals(TextInsideTheCell)).EBAttributesId = cellvalue.CellValue.Text;
                                }
                            }

                            if (_sheetName.Equals("API"))
                            {
                                cellvalue = row.Descendants<Cell>().Where(x => x.CellReference == "B" + row.RowIndex).FirstOrDefault();
                                if (cellvalue != null)
                                {
                                    newSpreadSheetData.Single(s => s.AttributesLCNames.Equals(TextInsideTheCell)).APIId = cellvalue.CellValue.Text;
                                }
                            }

                            if (_sheetName.Equals("ISA"))
                            {
                                cellvalue = row.Descendants<Cell>().Where(x => x.CellReference == "B" + row.RowIndex).FirstOrDefault();
                                if (cellvalue != null)
                                {
                                    newSpreadSheetData.Single(s => s.AttributesLCNames.Equals(TextInsideTheCell)).ISAID = cellvalue.CellValue.Text;
                                }
                            }

                            //This takes the cell text value and saves it to a object

                        }
                    }
                
            }

            return newSpreadSheetData;
        }

       

      
        private List<InformationSpreadSheet> GetAllNamesFromSheet(SharedStringTable sstforsheet1, List<Row> sheet1Rows)
        {
            List<InformationSpreadSheet> newSpreadSheetData = new List<InformationSpreadSheet>();
            foreach (Row row in sheet1Rows)
            {
                foreach (Cell c in row.Elements<Cell>())
                {
                    if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                    {
                        InformationSpreadSheet SheetOneData = new InformationSpreadSheet();
                        int ssid = int.Parse(c.CellValue.Text);
                        string str = sstforsheet1.ChildElements[ssid].InnerText;
                        SheetOneData.AttributesLCNames = str;
                        newSpreadSheetData.Add(SheetOneData);
                    }
                }
            }
            return newSpreadSheetData;
        }

    
    }
}
