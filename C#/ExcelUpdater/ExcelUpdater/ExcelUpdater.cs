namespace ExcelUpdater
{
    using Microsoft.Office.Interop.Excel;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Diagnostics;

    class ExcelUpdater
    {


        private string CalculateColumnLettersFromIndex(int columnIndex)
        {
            // Calculate the column letter excel style: 1 => A, 2 => B
            string columnLetter = "";
            int tempIndex = columnIndex;
            while (tempIndex > 26)
            {
                tempIndex -= 26;
                columnLetter = string.IsNullOrEmpty(columnLetter) ? "A" : ((char)(columnLetter[0] + 1)).ToString();
            }

            columnLetter += ((char)((columnIndex - 1) % 26 + 65)).ToString();

            return columnLetter;
        }

        private int FindMaxItemIndexInRange(Range copyFrom)
        {
            float maxValue = -1;
            int maxValueIndex = 1;
            int index = 1;

            foreach (Range cell in copyFrom)
            {
                float cellValue = float.Parse(cell.Value2.ToString());
                if (cellValue > maxValue)
                {
                    maxValue = cellValue;
                    maxValueIndex = index;
                }
                index++;
            }

            return maxValueIndex;
        }

        public void Play(string filePath, int initialRowIndex, int copyUntilMaxPlusRows)
        {
            Application oXL;
            _Workbook oWB;
            _Worksheet inputSheet, outputSheet;

            // Start Excel and get Application object.
            oXL = new Application();
            oXL.Visible = true;

            // Get a new workbook.
            oWB = (_Workbook)(oXL.Workbooks.Open(filePath));

            foreach(_Worksheet sheet2 in oWB.Sheets)
            {
                Debugger.Log(1,"asd", sheet2.Name);
            }

            _Worksheet sheet = oWB.Sheets.Add();
            sheet.Name = "Carlos";
        }


        public void UpdateExcel(string filePath, int initialRowIndex, int copyUntilMaxPlusRows)
        {
            Application oXL;
            _Workbook oWB;
            _Worksheet inputSheet, outputSheet;

            // Start Excel and get Application object.
            oXL = new Application();
            oXL.Visible = true;

            // Get a new workbook.
            oWB = (_Workbook)(oXL.Workbooks.Open(filePath));


            List<string> inputSheets = new List<string>();
            foreach (_Worksheet sheet in oWB.Sheets)
            {
                inputSheets.Add(sheet.Name);
            }

            foreach(string inputSheetName in inputSheets)
            {
                // Create output
                _Worksheet sheet = oWB.Sheets.Add();
                sheet.Name = inputSheetName + "-out";

                inputSheet = (_Worksheet)oWB.Sheets.Item[inputSheetName];
                outputSheet = (_Worksheet)oWB.Sheets.Item[inputSheetName + "-out"];

                UpdateExcel(initialRowIndex, copyUntilMaxPlusRows, inputSheet, outputSheet);
            }
        }

        public void UpdateExcel(int initialRowIndex, int copyUntilMaxPlusRows, _Worksheet inputSheet, _Worksheet outputSheet)
        {
            // maxRowIndex = index of the first blank cell in the first column
            int maxRowIndex = initialRowIndex;
            while(!string.IsNullOrWhiteSpace(inputSheet.Cells[maxRowIndex+1, 1].Text))
            {
                maxRowIndex++;
            }

            // INITIALIZATION DATA
            int initialDataRow = initialRowIndex + 1;               // First are column names titles
            int dataRowCount = maxRowIndex - initialDataRow + 1;    // How many rows of pure data do we have per column
            int columnIndex = 2;                                    // First column is datetime
            int nextOutRowStart = 1;                                // Which row we start writing in the Output sheet
            string columnName = inputSheet.Cells[initialRowIndex, columnIndex].Text;

            while (!string.IsNullOrWhiteSpace(columnName))
            {
                Range copyFrom;
                Range copyTo;
                int dataRowCountToCopy = dataRowCount;

                // Calculate the column letter excel style: 1 => A, 2 => B
                string columnLetter = CalculateColumnLettersFromIndex(columnIndex);

                // We need to find the max value and copy until maxvalue + copyUntilMaxPlusRows
                if (copyUntilMaxPlusRows >= 0)
                {
                     // Copy data column
                    copyFrom = inputSheet.Range[columnLetter + initialDataRow.ToString(), columnLetter + maxRowIndex.ToString()];

                    int maxValueIndex = FindMaxItemIndexInRange(copyFrom);

                    // we will copy only (maxValueIndex + copyUntilMaxPlusRows), or dataRowCount if we get out of range
                    dataRowCountToCopy = maxValueIndex + copyUntilMaxPlusRows;
                    dataRowCountToCopy = dataRowCountToCopy < dataRowCount ? dataRowCountToCopy : dataRowCount;
                }

                // Important row index that we will use to copy all columns
                var sourceDataInitialIndex      = initialDataRow.ToString();
                var sourceDataEndIndex          = (initialDataRow + dataRowCountToCopy - 1).ToString();
                var destinationDataInitialIndex = nextOutRowStart.ToString();
                var destinationDataEndIndex     = (nextOutRowStart + dataRowCountToCopy -1).ToString();

                // Copy columnName column
                string[] columnNameColumn = new string[dataRowCountToCopy];
                for (int i = 0; i < dataRowCountToCopy; i++) columnNameColumn[i] = columnName;
                copyTo = outputSheet.Range['A' + destinationDataInitialIndex, 'A' + destinationDataEndIndex];
                copyTo.Value2 = columnNameColumn;

                // Copy time column
                copyFrom = inputSheet.Range['A' + sourceDataInitialIndex, 'A' + sourceDataEndIndex];
                copyTo = outputSheet.Range['B' + destinationDataInitialIndex, 'B' + destinationDataEndIndex];
                copyTo.Value2 = copyFrom.Value2;

                // Copy data column
                copyFrom = inputSheet.Range[columnLetter + sourceDataInitialIndex, columnLetter + sourceDataEndIndex];
                copyTo = outputSheet.Range['C' + destinationDataInitialIndex, 'C' + destinationDataEndIndex];
                copyTo.Value2 = copyFrom.Value2;

                // Next iteration indexes
                columnIndex++;
                columnLetter = "";               
                nextOutRowStart += dataRowCountToCopy;
                columnName = inputSheet.Cells[initialRowIndex, columnIndex].Text;
            }
        }
    }
}
