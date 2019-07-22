namespace ExcelUpdater
{
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Collections;
    using System.Diagnostics;

    class ExcelUpdater
    {
        public void UpdateExcel(string filePath, int initialRow, int copyUntilMaxPlusRows)
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
            int dataRowCount = maxRow - initialDataRow + 1;

            char columnLetter = 'B';
            int columnIndex = 2;
            int nextOutRowStart = 1;
            string columnName = inputSheet.Cells[initialRow, columnIndex].Text;

            while (!string.IsNullOrWhiteSpace(columnName))
            {
                Range copyFrom;
                Range copyTo;
                int dataRowCountToCopy = dataRowCount;

                // We need to find the max value and copy until maxvalue + copyUntilMaxPlusRows
                if (copyUntilMaxPlusRows >= 0)
                {
                    float maxValue = -1;
                    int maxValueIndex = 1;
                    int index = 1;
                    // Copy data column
                    copyFrom = inputSheet.Range[columnLetter + initialDataRow.ToString(), columnLetter + maxRow.ToString()];

                    foreach (Range cell in copyFrom)
                    {
                        float cellValue = float.Parse(cell.Value2.ToString());
                        if(cellValue > maxValue)
                        {
                            maxValue = cellValue;
                            maxValueIndex = index;
                        }
                        index++;
                    }

                    // we will copy only (maxValueIndex + copyUntilMaxPlusRows), or dataRowCount if we get out of range
                    dataRowCountToCopy = maxValueIndex + copyUntilMaxPlusRows;
                    dataRowCountToCopy = dataRowCountToCopy < dataRowCount ? dataRowCountToCopy : dataRowCount;
                }

                var sourceDataInitialIndex      = initialDataRow.ToString();
                var sourceDataEndIndex          = (initialDataRow + dataRowCountToCopy - 1).ToString();
                var destinationDataInitialIndex = nextOutRowStart.ToString();
                var destinationDataEndIndex     = (nextOutRowStart + dataRowCountToCopy -1).ToString();

                // Copy data column
                copyFrom = inputSheet.Range[columnLetter + sourceDataInitialIndex, columnLetter + sourceDataEndIndex];
                copyTo = outputSheet.Range['C' + destinationDataInitialIndex, 'C' + destinationDataEndIndex];
                copyTo.Value2 = copyFrom.Value2;

                // Copy columnName column
                string[] columnNameColumn = new string[dataRowCountToCopy];
                for (int i = 0; i < dataRowCountToCopy; i++) columnNameColumn[i] = columnName;
                copyTo = outputSheet.Range['A' + destinationDataInitialIndex, 'A' + destinationDataEndIndex];
                copyTo.Value2 = columnNameColumn;

                // Copy time column
                copyFrom = inputSheet.Range['A' + sourceDataInitialIndex, 'A' + sourceDataEndIndex];
                copyTo = outputSheet.Range['B' + destinationDataInitialIndex, 'B' + destinationDataEndIndex];
                copyTo.Value2 = copyFrom.Value2;

                // Next iteration indexes
                columnLetter++;
                columnIndex++;
                nextOutRowStart += dataRowCountToCopy + 1;
                columnName = inputSheet.Cells[initialRow, columnIndex].Text;
            }
        }
    }
}
