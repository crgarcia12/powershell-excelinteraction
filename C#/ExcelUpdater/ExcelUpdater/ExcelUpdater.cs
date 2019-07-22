namespace ExcelUpdater
{
    using Microsoft.Office.Interop.Excel;
    using System;

    class ExcelUpdater
    {
        public void UpdateExcel(string filePath, int initialRow = 1)
        {
            Application oXL;
            _Workbook oWB;
            _Worksheet inputSheet, outputSheet;

            //Start Excel and get Application object.
            oXL = new Application();
            oXL.Visible = true;

            //Get a new workbook.
            oWB = (_Workbook)(oXL.Workbooks.Open(filePath));
            inputSheet = (_Worksheet)oWB.Sheets.Item["input"];
            outputSheet = (_Worksheet)oWB.Sheets.Item["output"];

            int maxRow = initialRow;
            while(!string.IsNullOrWhiteSpace(inputSheet.Cells[maxRow+1, 1].Text))
            {
                maxRow++;
            }

            int initialDataRow = initialRow + 1;
            int dataRowCount = maxRow - initialDataRow;

            char columnLetter = 'B';
            int columnIndex = 2;
            int nextOutRowStart = 1;
            string columnName = inputSheet.Cells[initialRow, columnIndex].Text;

            while (!string.IsNullOrWhiteSpace(columnName))
            {
                    Range copyFrom;
                    Range copyTo;

                    // Copy columnName
                    string[] columnNameColumn = new string[dataRowCount];
                    for (int i = 0; i < dataRowCount; i++) columnNameColumn[i] = columnName;
                    copyTo = outputSheet.Range['A' + nextOutRowStart.ToString(), 'A' + (nextOutRowStart + dataRowCount).ToString()];
                    copyTo.Value2 = columnNameColumn;

                    // Copy time
                    copyFrom = inputSheet.Range['A' + initialDataRow.ToString(), 'A' + maxRow.ToString()];
                    copyTo = outputSheet.Range['B' + nextOutRowStart.ToString(), 'B' + (nextOutRowStart + dataRowCount).ToString()];
                    copyTo.Value2 = copyFrom.Value2;
                        
                    // Copy data
                    copyFrom = inputSheet.Range[columnLetter + initialDataRow.ToString(), columnLetter + maxRow.ToString()];
                    copyTo = outputSheet.Range['C' + nextOutRowStart.ToString(), 'C' + (nextOutRowStart + dataRowCount).ToString()];
                    copyTo.Value2 = copyFrom.Value2;

                    // Next iteration
                    columnLetter++;
                    columnIndex++;
                    nextOutRowStart += dataRowCount + 1;
                    columnName = inputSheet.Cells[initialRow, columnIndex].Text;
            }
        }
    }
}
