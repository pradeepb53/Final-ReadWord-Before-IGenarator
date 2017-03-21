using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using ReadWord;
using System.Text.RegularExpressions;

namespace ReadWord
{
    class Program
    {
        private Application app = new Application();



        private static List<TranscriptWord> CustomWordDirectory = new List<TranscriptWord>();

        private string currentWord = "";



        static void Main(string[] args)
        {

            // Microsoft.Office.Interop.Word.Range rng = document.Range(Start: i+1, End: i+1);
            //int lineNumber = rng.Information[WdInformation.wdActiveEndPageNumber];
            //Microsoft.Office.Interop.Word.WdUnits.

            // Range rng = app.ActiveDocument.Range(0, lengthOfCurrentWord);
            // int currentPageNumber = rng.Information[WdInformation.wdActiveEndPageNumber];
            // int pageNumber = document.Words[i ].ComputeStatistics(WdStatistic.wdStatisticPages);
            //int currentPageNumber = app.Selection.Information[WdInformation.wdActiveEndPageNumber];

            //Range rng2 = document.Content;
            //int inumber = rng2.ListFormat.ListLevelNumber;
            //string iString = rng2.ListFormat.ListString;

            //Microsoft.Office.Interop.Word.WdStatistic stat = Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages;
            //int pages = document.ComputeStatistics(stat, Type.Missing);
            //Console.WriteLine(pages);
            //Console.ReadLine();
            //_________________________________________________________________________________________________________________________-



            Program WordIndex = new Program();
               WordIndex.processDocument();  // Read main document

             WordIndex.printWordIndex(); // Old print method
            // WordIndex.createIndexTable(); // Index Table POC

            //WordIndex.createWrodIndexTable(); // Real Index
            // WordIndex.createWrodNewIndexTable(); // New index - page at a time 12/25

            // WordIndex.createIndex2017();
           // WordIndex.createSampleText();

        }

        private void createSampleText()
        {
            Microsoft.Office.Interop.Excel.Application oxApp = new Microsoft.Office.Interop.Excel.Application();
            //oxApp.Visible = true;
            oxApp.Workbooks.Add();

            Microsoft.Office.Interop.Excel._Worksheet oxWorkSheet = oxApp.ActiveSheet;

            int lineNumber = 0;
            float pageRowNumber = 0;
            int columnNumber = 1;
            int initialRowOnAllColumns = 0;
            int rowOnColumnB = 0;
            int rowOnColumnC = 0;
            int addCount = 0;
            

            List<String> testArray = new List<string>();

            for (int i = 0; i < 300; i++)
            {
                testArray.Add("This is : " + i);
            }

            //StreamWriter textWriter = new StreamWriter("C:\\user_pradeep\\fvrs\\index testing\\LogText.txt");
            foreach (var item in testArray)
            {
                lineNumber++;
                pageRowNumber++;
                //if (lineNumber <= 47)
                    if (pageRowNumber/47 <= 1)
                    {
                    oxWorkSheet.Cells[lineNumber, "A"] = item;
                    addCount++;
                    if (addCount == 1)
                        initialRowOnAllColumns = lineNumber;
                   // columnNumber++;
                  //  if (lineNumber == 47)
                        if (pageRowNumber/47 == 1)
                        {
                         rowOnColumnB = initialRowOnAllColumns;
                         rowOnColumnC = initialRowOnAllColumns;
                         addCount = 0;
                    }
                        
                }
               // if((lineNumber > 47) && (lineNumber <= 94))
                    if ((pageRowNumber/47 > 1) && (pageRowNumber/47 <= 2))
                    {
                    oxWorkSheet.Cells[rowOnColumnB, "B"] = item;
                    rowOnColumnB++;
                   // if (lineNumber == 94)
                        if (pageRowNumber/47 == 2)
                        {
                        rowOnColumnB = 0;
                    }
                }

               // if ((lineNumber > 93) && (lineNumber <= 140))
                    if ((pageRowNumber/47 > 2) && (pageRowNumber/47 <= 3))
                    {
                    oxWorkSheet.Cells[rowOnColumnC, "C"] = item;
                    rowOnColumnC++;
                    //if (lineNumber == 140)
                        if (pageRowNumber/47 == 3)
                        {
                        rowOnColumnC = 0;
                        pageRowNumber = 0;
                    }
                }
            }

            oxWorkSheet.SaveAs("C:\\user_pradeep\\fvrs\\index testing\\LogTextNew3.xlsx");
           
            //textWriter.Close();
            //textWriter.Dispose();
        }

        private void createIndex2017()
        {
            List<TranscriptWord> SCustomWordDirectory = CustomWordDirectory.OrderBy(o => o.Name).ToList();
            string[,] indexList = new string[5, 3];
            int i = 0;
            int wordCount = 0;

            foreach (TranscriptWord item in SCustomWordDirectory)
            {
                wordCount++;

                if (wordCount <=5 )
                {
                    indexList[i, 0] = item.Name;
                    i++;
                    if (wordCount == 5)
                    {
                        i = 0;
                    } 
                }
                if ((wordCount >5) && (wordCount <=10))
                {
                    indexList[i, 1] = item.Name;
                    i++;

                    if (wordCount == 10)
                    {
                        i = 0;
                    }
                }

                if ((wordCount > 10) && (wordCount <= 15))
                {
                    indexList[i, 2] = item.Name;
                    i++;

                    if (wordCount == 15)
                    {
                        i = 0;
                    }
                }
                if (wordCount ==15)
                {
                    break;
                }
            }

            for (int y = 0; y < 5; y++)
            {
                
                    Console.WriteLine(indexList[y, 0] + "|" + indexList[y, 1] + "|" + indexList[y, 2]);
                
                
            }
            Console.ReadLine();
        }

        private void createWrodNewIndexTable()   // 12/08
        {
            List<TranscriptWord> SCustomWordDirectory = CustomWordDirectory.OrderBy(o => o.Name).ToList();

            //Document properties
            //===================
            Document indexTableDoc = app.Documents.Add();
            Range indexTableRange = indexTableDoc.Range();

            // Table properties
            //=================
            Table myTable = indexTableDoc.Tables.Add(indexTableRange, 1, 3);
            int columnNumber = 1;
            int rowNumber = 0;
            int rowNumberOnNewPage = 0;
            int totalRowsOnCurrentPage = 0;
            int startingRowOfTheActivePage = 0;

            // Page and Range properties
            //==========================
            object oMissing = System.Reflection.Missing.Value;   
            Range cellRange = myTable.Cell(1, 1).Range;

            int previousPageNumber = myTable.Rows[1].Cells[1].Range.Information[WdInformation.wdActiveEndPageNumber];
            int currentPageNumber = previousPageNumber; // Both set to same value to get started
            
            
            //Temp storage properties
            //=======================
            int rowCountPageAndLineNumber = 0; 
            int rowOnePageNumber = 0;
            int rowOneLineNumber = 0;
            int pageNumber = 0;

            foreach (TranscriptWord item in SCustomWordDirectory)
            {

                if (columnNumber == 1) //Add new rows, only when range is on column #1
                {
                    myTable.Rows.Add(ref oMissing);
                    rowNumber++;
                    currentPageNumber = myTable.Rows[rowNumber].Cells[columnNumber].Range.Information[WdInformation.wdActiveEndPageNumber];

                    //if (currentPageNumber != previousPageNumber)
                        if (rowNumber == 44)
                        {
                        rowNumberOnNewPage = rowNumber;
                        rowNumber--;
                        myTable.Cell(rowNumber, columnNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                        totalRowsOnCurrentPage = rowNumber;
                        startingRowOfTheActivePage = (rowNumberOnNewPage - totalRowsOnCurrentPage) - 1;
                        rowNumber = startingRowOfTheActivePage;
                        columnNumber++;
                        previousPageNumber = currentPageNumber;

                    }
                }

                //Column >1 and < 3 - Need to naviagte through all rows untill row count is == total rows on the current page 

                if ((columnNumber != 1) && (columnNumber < 3))
                {
                    if (rowNumber < totalRowsOnCurrentPage)
                    {
                        rowNumber++;
                    }
                    else
                    {
                        myTable.Cell(rowNumber, columnNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                        rowNumber = startingRowOfTheActivePage;
                        columnNumber++;
                    }
                }

                if (columnNumber == 3)
                {
                    if (rowNumber < totalRowsOnCurrentPage)
                    {
                        rowNumber++;
                    }
                    else
                    {
                        myTable.Cell(rowNumber, columnNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                        rowNumber = rowNumberOnNewPage;
                        columnNumber = 1;
                        totalRowsOnCurrentPage = 0;
                    }
                    
                }

                cellRange = myTable.Rows[rowNumber].Cells[columnNumber].Range;
                cellRange.Font.Size = 10;
                cellRange.Font.Bold = 1;
                pageNumber = myTable.Rows[rowNumber].Cells[columnNumber].Range.Information[WdInformation.wdActiveEndPageNumber];

                myTable.Rows[rowNumber].Cells[columnNumber].Range.Text = item.Name + " [" + item.Frequency + "]" + " Page# " + pageNumber;

                rowCountPageAndLineNumber = 0;
                rowOnePageNumber = 0;
                rowOneLineNumber = 0;

                foreach (Occurrence step in item.PageAndLine)
                {

                    rowCountPageAndLineNumber++;

                    if (rowCountPageAndLineNumber == 2)
                    {
                        //if (columnNumber == 1) //Add new rows, only when range is on column #1
                        if (rowNumber == 44)
                        {
                            myTable.Rows.Add(ref oMissing);
                            rowNumber++;
                            currentPageNumber = myTable.Rows[rowNumber].Cells[columnNumber].Range.Information[WdInformation.wdActiveEndPageNumber];

                            if (currentPageNumber != previousPageNumber)
                            {
                                rowNumberOnNewPage = rowNumber;
                                rowNumber--;
                                myTable.Cell(rowNumber, columnNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                                totalRowsOnCurrentPage = rowNumber;
                                startingRowOfTheActivePage = (rowNumberOnNewPage - totalRowsOnCurrentPage) - 1;
                                rowNumber = startingRowOfTheActivePage;
                                columnNumber++;
                                previousPageNumber = currentPageNumber;

                            }
                        }

                        if ((columnNumber != 1) && (columnNumber < 3))
                        {
                            if (rowNumber < totalRowsOnCurrentPage)
                            {
                                rowNumber++;
                            }
                            else
                            {
                                myTable.Cell(rowNumber, columnNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                                rowNumber = startingRowOfTheActivePage;
                                columnNumber++;
                            }
                        }

                        if (columnNumber == 3)
                        {
                            if (rowNumber < totalRowsOnCurrentPage)
                            {
                                rowNumber++;
                            }
                            else
                            {
                                myTable.Cell(rowNumber, columnNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                                rowNumber = rowNumberOnNewPage;
                                columnNumber = 1;
                                totalRowsOnCurrentPage = 0;
                            }

                        }


                        cellRange = myTable.Cell(rowNumber, columnNumber).Range;
                        cellRange.Font.Size = 7;
                        cellRange.Font.Bold = 0;
                        pageNumber = myTable.Rows[rowNumber].Cells[columnNumber].Range.Information[WdInformation.wdActiveEndPageNumber];

                      //  myTable.Cell(rowNumber, columnNumber).TopPadding = 0.00f;
                        cellRange.ParagraphFormat.SpaceBefore = 0.00f;

                        cellRange.Text = "[P" + rowOnePageNumber + ":" + "L" + rowOneLineNumber + "] [P" + step.CustomPageNumber + ":" + "L" + step.CustomLineNumber + "]" + " Page# " + pageNumber;

                        cellRange.ParagraphFormat.SpaceAfter = 0.00f;
                       // myTable.Cell(rowNumber, columnNumber).BottomPadding = 0.00f;
                       
                        rowCountPageAndLineNumber = 0;
                        rowOnePageNumber = 0;
                        rowOneLineNumber = 0;
                    }
                    else
                    {
                        rowOnePageNumber = step.CustomPageNumber;
                        rowOneLineNumber = step.CustomLineNumber;
                    }
                   

                }

                if (rowCountPageAndLineNumber == 1)
                {
                    if (columnNumber == 1) //Add new rows, only when range is on column #1
                    {
                        myTable.Rows.Add(ref oMissing);
                        rowNumber++;
                        currentPageNumber = myTable.Rows[rowNumber].Cells[columnNumber].Range.Information[WdInformation.wdActiveEndPageNumber];

                        // if (currentPageNumber != previousPageNumber)
                        if (rowNumber == 44)
                        {
                            rowNumberOnNewPage = rowNumber;
                            rowNumber--;
                            myTable.Cell(rowNumber, columnNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                            totalRowsOnCurrentPage = rowNumber;
                            startingRowOfTheActivePage = (rowNumberOnNewPage - totalRowsOnCurrentPage) - 1;
                            rowNumber = startingRowOfTheActivePage;
                            columnNumber++;
                            previousPageNumber = currentPageNumber;

                        }
                    }

                    if ((columnNumber != 1) && (columnNumber < 3))
                    {
                        if (rowNumber < totalRowsOnCurrentPage)
                        {
                            rowNumber++;
                        }
                        else
                        {
                            myTable.Cell(rowNumber, columnNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                            rowNumber = startingRowOfTheActivePage;
                            columnNumber++;
                        }
                    }

                    if (columnNumber == 3)
                    {
                        if (rowNumber < totalRowsOnCurrentPage)
                        {
                            rowNumber++;
                        }
                        else
                        {
                            myTable.Cell(rowNumber, columnNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                            rowNumber = rowNumberOnNewPage;
                            columnNumber = 1;
                            totalRowsOnCurrentPage = 0;
                        }

                    }

                    cellRange = myTable.Cell(rowNumber, columnNumber).Range;
                    cellRange.Font.Size = 7;
                    cellRange.Font.Bold = 0;
                    pageNumber = myTable.Rows[rowNumber].Cells[columnNumber].Range.Information[WdInformation.wdActiveEndPageNumber];

                   // myTable.Cell(rowNumber, columnNumber).TopPadding = 0.00f;
                    cellRange.ParagraphFormat.SpaceBefore = 0.00f;

                    cellRange.Text = "[P" + rowOnePageNumber + ":" + "L" + rowOneLineNumber + "] " + " Page# " + pageNumber;

                    cellRange.ParagraphFormat.SpaceAfter = 0.00f;
                    //myTable.Cell(rowNumber, columnNumber).BottomPadding = 0.00f;
                }

            }


            indexTableDoc.Save();
            indexTableDoc.Close();
            app.Quit();
        }
        private void createWrodIndexTable()   // 11/03
        {
            List<TranscriptWord> SCustomWordDirectory = CustomWordDirectory.OrderBy(o => o.Name).ToList();


            Document indexTableDoc = app.Documents.Add();
            Range indexTableRange = indexTableDoc.Range();

            Table myTable = indexTableDoc.Tables.Add(indexTableRange, 600, 3);

           // myTable.BottomPadding = 0.00f;
            
            Range cellRange = myTable.Cell(1, 1).Range;

            int colNumber = 1;
            int rowNumber = 0;
            int actualRowCount = 0;
            int tableRowCount = 0;

            int columnCountPageAndLine = 0; // Two rows of page and line numbers grouped to make 1 row with 2 columns
            int columnOnePageNumber = 0;
            int columnOneLineNumber = 0;

            foreach (TranscriptWord item in SCustomWordDirectory)
            {
                

                if ((tableRowCount == 51) && (colNumber < 3))
                {
                    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

                    rowNumber = (actualRowCount - 51);
                    colNumber++;
                    tableRowCount = 0;

                }

                if ((tableRowCount == 51) && (colNumber == 3))
                {
                    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;


                    //indexTableRange = indexTableDoc.Range(Start: myTable.Cell(rowNumber, colNumber).Range, End: myTable.Cell(rowNumber +1, colNumber).Range);
                    indexTableRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    indexTableRange.InsertBreak(WdBreakType.wdPageBreak);
                    indexTableRange.Collapse(WdCollapseDirection.wdCollapseEnd);


                    rowNumber = actualRowCount;
                    tableRowCount = 0;
                    colNumber = 1;
                }

                if (colNumber == 1)
                {
                    actualRowCount++;

                }

                rowNumber++;
                tableRowCount++;

                cellRange = myTable.Cell(rowNumber, colNumber).Range;
                
                cellRange.Font.Size = 10;
                cellRange.Font.Bold = 1;

                Cell cell1 = myTable.Cell(rowNumber, colNumber);

                cell1.Range.ParagraphFormat.SpaceBefore = 0.00f;

                cellRange.Text = item.Name + " [" + item.Frequency + "]";

                cell1.Range.ParagraphFormat.SpaceAfter = 0.00f;

                // cellRange.ParagraphFormat.SpaceAfter = 0;

                myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;

                if (actualRowCount == 1) // Top border only on top of the page - for now
                {
                    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                }

                //if (actualRowCount == 32)
                //{
                //    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                //}

                columnCountPageAndLine = 0;
                 columnOnePageNumber = 0;
                 columnOneLineNumber = 0;        


                foreach (Occurrence step in item.PageAndLine)
                {       

                    columnCountPageAndLine++; // Row count for Page and Line numbers

                    if (columnCountPageAndLine == 2)
                    {

                        if ((tableRowCount == 51) && (colNumber < 3))
                        {
                            myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

                            rowNumber = (actualRowCount - 51);
                            colNumber++;
                            tableRowCount = 0;

                        }

                        if ((tableRowCount == 51) && (colNumber == 3))
                        {
                             myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

                            //indexTableRange = indexTableDoc.Range(Start: myTable.Cell(rowNumber, colNumber).Range, End: myTable.Cell(rowNumber + 1, colNumber).Range);
                            indexTableRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                            indexTableRange.InsertBreak(WdBreakType.wdPageBreak);
                            indexTableRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                            rowNumber = actualRowCount;
                            tableRowCount = 0;
                            colNumber = 1;
                        }

                        if (colNumber == 1)
                        {
                            actualRowCount++;

                        }

                        rowNumber++;
                        tableRowCount++;
                        
                        cellRange = myTable.Cell(rowNumber, colNumber).Range;
                       
                        cellRange.Font.Size = 7;
                        cellRange.Font.Bold = 0;

                        cellRange.ParagraphFormat.SpaceBefore = 0.00f;           // Today
                        myTable.Cell(rowNumber, colNumber).TopPadding = 0.00f;

                        Cell cell2 = myTable.Cell(rowNumber, colNumber);
                       
                        cell2.Range.ParagraphFormat.SpaceBefore = 0.00f;

                        cellRange.Text = "[P" + columnOnePageNumber + ":" + "L" + columnOneLineNumber + "] [P" + step.CustomPageNumber + ":" + "L" + step.CustomLineNumber + "]";
                        

                        cellRange.ParagraphFormat.SpaceAfter = 0.00f;        // Today
                        cell2.Range.ParagraphFormat.SpaceAfter = 0.00f;

                        myTable.Cell(rowNumber, colNumber).BottomPadding = 0.00f;

                        myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                        myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;


                        //if (tableRowCount == 1)
                        //{
                        //    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                        //}           

                        columnCountPageAndLine = 0;
                        columnOnePageNumber = 0;
                        columnOneLineNumber = 0;

                    }
                    else
                    {
                        columnOnePageNumber = step.CustomPageNumber;
                        columnOneLineNumber = step.CustomLineNumber;

                    }

                    

                }

                //If columnCount is 1, only 1 occurrence found, then print one line and reset counter 

                if (columnCountPageAndLine == 1)
                {
                    if ((tableRowCount == 51) && (colNumber < 3))
                    {
                        myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

                        rowNumber = (actualRowCount - 51);
                        colNumber++;
                        tableRowCount = 0;

                    }

                    if ((tableRowCount == 51) && (colNumber == 3))
                    {
                        myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;

                       // indexTableRange = indexTableDoc.Range(Start: myTable.Cell(rowNumber, colNumber).Range, End: myTable.Cell(rowNumber + 1, colNumber).Range);
                        indexTableRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                        indexTableRange.InsertBreak(WdBreakType.wdPageBreak);
                        indexTableRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                        rowNumber = actualRowCount;
                        tableRowCount = 0;
                        colNumber = 1;
                    }

                    if (colNumber == 1)
                    {actualRowCount++;

                    }

                    rowNumber++;
                    tableRowCount++;                

                    cellRange = myTable.Cell(rowNumber, colNumber).Range;

                    cellRange.Font.Size = 7;
                    cellRange.Font.Bold = 0;

                    cellRange.ParagraphFormat.SpaceBefore = 0.00f;           // Today
                    myTable.Cell(rowNumber, colNumber).TopPadding = 0.00f;

                    Cell cell3 = myTable.Cell(rowNumber, colNumber);

                    cell3.Range.ParagraphFormat.SpaceBefore = 0.00f;

                    cellRange.Text = "[P" + columnOnePageNumber + ":" + "L" + columnOneLineNumber + "]";

                    cellRange.ParagraphFormat.SpaceAfter = 0;          //Today
                    cell3.Range.ParagraphFormat.SpaceAfter = 0.00f;

                    myTable.Cell(rowNumber, colNumber).BottomPadding = 0.00f;

                    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;

                }
                
            }
   

            indexTableDoc.Save();
            indexTableDoc.Close();
            app.Quit();
        }

        private void createIndexTable()
        {
            Document indexTableDoc = app.Documents.Add();
            Range indexTableRange = indexTableDoc.Range();

            Table myTable = indexTableDoc.Tables.Add(indexTableRange, 35, 3);

            Range cellRange = myTable.Cell(1, 1).Range;

            int colNumber = 1;
            int rowNumber = 0;
            int actualRowCount = 0;
            int tableRowCount = 0;

            for (int i = 1; i <= 126; i++)
            {

                if (colNumber == 1)
                {
                    actualRowCount++;

                }

                rowNumber++;
                tableRowCount++;

               // myTable.Cell(rowNumber, colNumber).BottomPadding = 0.00f;
                //myTable.Cell(rowNumber, colNumber).TopPadding = 0.00f;

                cellRange = myTable.Cell(rowNumber, colNumber).Range;

                //cellRange.ParagraphFormat.SpaceBefore = 0.0f;

                cellRange.Text = "The value is :" + i + " column:" + colNumber ;

                //cellRange.ParagraphFormat.SpaceAfter = 0.0f;

                myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
               // if ((actualRowCount == 9) || (actualRowCount == 21) || (actualRowCount == 30))
              // {
                  //  cellRange.Font.Size = 10;
                 //   cellRange.Font.Bold = 1;
                    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
               //}


                if ((tableRowCount == 28) && (colNumber < 3))
                {

                    rowNumber = (actualRowCount - 28);
                    colNumber++;
                    tableRowCount = 0;

                }

                if ((tableRowCount == 28) && (colNumber == 3))
                {

                    rowNumber = actualRowCount;
                    tableRowCount = 0;
                    colNumber = 1;
                }
            }

            indexTableDoc.Save();
            indexTableDoc.Close();
            app.Quit();
        }

        private void printWordIndex()
        {
            List<TranscriptWord> SCustomWordDirectory = CustomWordDirectory.OrderBy(o => o.Name).ToList();

           // int totalElements = SCustomWordDirectory.Count;


            //Console.WriteLine("Total number of elements in the List : " + totalElements);

            Document indexDoc = app.Documents.Add();
            Range indexRange = indexDoc.Range();
            indexDoc.PageSetup.TextColumns.SetCount(5);
            indexDoc.Activate();
            // indexDoc.Range().Select();
            indexRange.Select();

            //int mainElements = 0;
           // int subElements = 0;
            string previousLabel = ""; //03/17/17
            bool isNumber = false; //03/17/17
            bool isLetter = false; //03/17/17
            bool numberSignPrinted = false; //02/17/17
            bool firstLetterPrinted = false; //02/17/17

            foreach (TranscriptWord item in SCustomWordDirectory)
            {
                //Check first charactor of the word to determine type and identification
                //-----------------------03/17/17--------------------------------------

                string firstCharOfTheWord = item.Name.Substring(0, 1);
                int valueOne = 0;
                
                bool parsed = Int32.TryParse(firstCharOfTheWord, out valueOne);

                if (parsed)
                {
                    isNumber = true;
                    isLetter = false;
                   
                    if (previousLabel == "#")
                    {
                        numberSignPrinted = true;
                    }
                }
                else
                {
                    if (Regex.IsMatch(firstCharOfTheWord, @"^[a-zA-Z]"))
                    {
                        isLetter = true;
                        isNumber = false;

                        if (previousLabel == firstCharOfTheWord)
                        {
                            firstLetterPrinted = true;
                        }else
                        {
                            firstLetterPrinted = false;
                        }
                    }

                }
                  
                if((isNumber) && (!numberSignPrinted))
                {
                    previousLabel = "#";

                    Paragraph para0 = indexDoc.Paragraphs.Add();
                    para0.Range.Font.Size = 15;
                    para0.Range.Font.Bold = 1;
                    para0.Range.Text = " " + previousLabel + " " +  "\r\n";
                    numberSignPrinted = true;
                }

                if ((isLetter) && (!firstLetterPrinted))
                {
                    previousLabel = firstCharOfTheWord;

                    Paragraph para0 = indexDoc.Paragraphs.Add();
                    para0.Range.Font.Size = 15;
                    para0.Range.Font.Bold = 1;
                    para0.Range.Text = "- " + previousLabel.ToUpper() + " -" + "\r\n";
                    firstLetterPrinted = true;
                }
                
                //-----------------------03/17/17--------------------------------------

                Paragraph para1 = indexDoc.Paragraphs.Add();
                para1.Range.Font.Size = 10;
                para1.Range.Font.Bold = 1;

                
               
                para1.Range.Text = item.Name + " [" + item.Frequency + "]" + "\r\n";  // 10/26

                int columnCount = 0;
                int columnOnePageNumber = 0;
                int columnOneLineNumber = 0;

               // mainElements++;

                foreach (Occurrence step in item.PageAndLine)
                {
                    columnCount++; // Actually row count

                   // subElements++;

                    if (columnCount == 2)
                    {
                        Paragraph para2 = indexDoc.Paragraphs.Add();
                        para2.Range.Font.Size = 7;
                        para2.Range.Font.Bold = 0;

                        para2.Range.Text = "[P" + columnOnePageNumber + ":" + "L" + columnOneLineNumber + "] [P" + step.CustomPageNumber + ":" + "L" + step.CustomLineNumber + "]" + "\r\n";
                        para2.Range.ParagraphFormat.SpaceAfter = 0;

                        columnCount = 0;
                        columnOnePageNumber = 0;
                        columnOneLineNumber = 0;
                    }
                    else
                    {
                         columnOnePageNumber = step.CustomPageNumber;
                         columnOneLineNumber = step.CustomLineNumber;

                    }
                    
                }

                //If columnCount is 1, then print one line and reset counter 
                if (columnCount == 1)
                {
                    Paragraph para3 = indexDoc.Paragraphs.Add();
                    para3.Range.Font.Size = 7;
                    para3.Range.Font.Bold = 0;
                    para3.Range.Text = "[P" + columnOnePageNumber + ":" + "L" + columnOneLineNumber + "]" + "\r\n";
                    para3.Range.ParagraphFormat.SpaceAfter = 0;

                    columnCount = 0;
                    columnOnePageNumber = 0;
                    columnOneLineNumber = 0;
                }

            }

           // Console.WriteLine("Number of Words : " + mainElements + " and sub items : " + subElements);

            indexDoc.Save();
            indexDoc.Close();
            app.Quit();
        }

        private void processDocument()
        {

            Document document = app.Documents.Open(@"C:\User_Pradeep\Transcript2.doc", ReadOnly: true);
            document.Activate();

            //int totalWordCount = document.Words.Count; 10/25

            //int wordCountArrayLength = 0; 10/24

            //string[] processedWordList = new string[totalWordCount]; 10/24

            //List<string> processedWordList = new List<string>(); 10/25

            HashSet<string> processedWordList = new HashSet<string>(); // 10/25
           

            String textOftheWord = "";


            for (int i = 1; i < document.Words.Count; i++)
            {

                textOftheWord = document.Words[i].Text.ToString().Trim().ToLower(); // 10/26

               

                //if ((textOftheWord.Length > 3) && (textOftheWord.All(Char.IsLetterOrDigit))) // 10/23 // Comma issues



                if ((textOftheWord.Length > 3) && (textOftheWord.Contains("_________") == false) && (textOftheWord != "-")) // 10/26
                    {

                    string processedWord = processedWordList.FirstOrDefault(w => w == textOftheWord); // 10/25

                    //int wordFoundIntheProcessedWordList = Array.IndexOf(processedWordList, textOftheWord); // 10/23 10/24

                    //if ((wordFoundIntheProcessedWordList < 0)) // 10/23 10/24

                   
                        if (processedWord == null) // 10/24
                        {

                        int valueInTheWord = 0; // 10/22

                        bool parsed = Int32.TryParse(textOftheWord, out valueInTheWord); // 10/22

                        //Ignore *All numbers and continues and dotted lines

                        if ((!parsed) )  // 10/22
                        {

                            //curRange = document.Words[i];
                            Range currentRange = document.Range(Start: document.Words[i].Start, End: document.Words[i].End); // 10/22

                            int pageNumberOfTheCurrentRange = currentRange.Information[WdInformation.wdActiveEndPageNumber]; // 10/22


                            //if ((textOftheWord.Length > 3) && (pageNumberOfTheCurrentRange > 1) && (textOftheWord != "__") && (textOftheWord != "-")) // 10/22

                            //if (pageNumberOfTheCurrentRange > 1) // 10/23


                            if ((pageNumberOfTheCurrentRange > 1) ) // 10/26
                            {


                                Console.WriteLine("Now processing word # " + i + "  In page # " + pageNumberOfTheCurrentRange);


                                //Check whether word is already processed?

                              //  int wordPosition = Array.IndexOf(processedWordList, textOftheWord); 10/24

                                var CustomWord = new TranscriptWord();

                                // Not in the array, it is a new word

                                processedWordList.Add(textOftheWord); // 10/24

                               // if (wordPosition < 0) //10/24
                               // {
                                    //Add new word to processed list of words

                                   // if (wordCountArrayLength <= totalWordCount) // 10/24
                                    //{
                                      //  processedWordList[wordCountArrayLength] = textOftheWord; 10/24
                                       // wordCountArrayLength++; 10/24

                                        // Update CustomWord name property
                                        CustomWord.Name = textOftheWord;

                                    //}

                                    int wordFoundFrequency = 0;

                                    // Range fullRange = document.Content; // 10/23

                                    Range searchRange = document.Range(Start: document.Words[i].Start, End: document.Content.End); // 10/23

                                    //Range fullRange = curRange;
                                    searchRange.Find.Forward = true; // 10/26 all range words
                                    searchRange.Find.Text = textOftheWord;

                                    currentWord = textOftheWord;

                                    searchRange.Find.Execute(MatchWholeWord: true);
                                    int currentLineNumber = 0;
                                    int currentPageNumber = 0;
                                    int pageNumberOfTheWord = 0;
                                    int lineNumberOfTheWord = 0;


                                    while (searchRange.Find.Found)
                                    {
                                        wordFoundFrequency++;

                                        Console.WriteLine("Looking for word : " + currentWord);

                                        //___________________________________________
                                        //  Sentences sent = fullRange.Sentences;
                                        // Range rng = document.Range(Start: sent, End: sent);

                                        //Microsoft.Office.Interop.Word.Selection sel = document.S

                                        //Total papagraphs in the documnet

                                        // int paraCountTotal = document.Paragraphs.Count; 10/22
                                        //Paragraph count in the selected range = Last paragraph read 
                                        // int totalParagraphsRead = app.ActiveDocument.Range(0, fullRange.End).Paragraphs.Count; 10/22



                                        // Range rangeFirstSentence = fullRange.Sentences.First;  10/22
                                        //Text of the current sentence
                                        // string textOfTheCurrentSentence = rangeFirstSentence.Text;    // Entire Line   // 10/22
                                        //------------------------------------------------------------------------------------------------------
                                        //int lineNumberOfTheCurrentSentence = getLineNumberOfTheWord(textOfTheCurrentSentence);



                                        // string text2 = fullRange.Text;   // Selected Word   // 10/22

                                        currentPageNumber = searchRange.Information[WdInformation.wdActiveEndPageNumber];

                                        currentLineNumber = searchRange.Information[WdInformation.wdFirstCharacterLineNumber];

                                        if (currentPageNumber != 1) // Page # 1 is cover page 10/22
                                        {
                                            //Check whether current word is repeating in the same page and line number, if not, create "Occurrence" object. 
                                            if (wordFoundFrequency > 1)
                                            {
                                                if (pageNumberOfTheWord != currentPageNumber || lineNumberOfTheWord != currentLineNumber)
                                                {
                                                    pageNumberOfTheWord = currentPageNumber;
                                                    lineNumberOfTheWord = currentLineNumber;

                                                    var CustomOccurrence = new Occurrence { CustomPageNumber = pageNumberOfTheWord, CustomLineNumber = lineNumberOfTheWord };
                                                    CustomWord.PageAndLine.Add(CustomOccurrence);
                                                }

                                            }
                                            else
                                            {
                                                pageNumberOfTheWord = currentPageNumber;
                                                lineNumberOfTheWord = currentLineNumber;

                                                var CustomOccurrence = new Occurrence { CustomPageNumber = pageNumberOfTheWord, CustomLineNumber = lineNumberOfTheWord };
                                                CustomWord.PageAndLine.Add(CustomOccurrence);
                                            }
                                        }


                                        searchRange.Find.Execute(MatchWholeWord: true);

                                    }

                                    CustomWord.Frequency = wordFoundFrequency;
                                    CustomWordDirectory.Add(CustomWord);

                                }

                            }
                        }

                    }
                }
            //}

            document.Close();

        }

        private int getLineNumberOfTheWord(string textOfTheCurrentRead)
        {
            string firstCharOfTheText = textOfTheCurrentRead.Substring(0, 1);
            int valueOne = 0;
            int lineNumber = 0;
            bool parsed = Int32.TryParse(firstCharOfTheText, out valueOne);

            if (!parsed)
            {
                return lineNumber;
            }
            else
            {
                lineNumber = valueOne;

                string secondCharOfTheText = textOfTheCurrentRead.Substring(1, 1);
                int valueTwo = 0;
                parsed = Int32.TryParse(secondCharOfTheText, out valueTwo);

                if (parsed)
                {
                    lineNumber = int.Parse(valueOne.ToString() + valueTwo.ToString());
                }
            }

            return lineNumber;
        }


    }

    class TranscriptWord
    {
        public string Name { get; set; }
        public int Frequency { get; set; }
        public List<Occurrence> PageAndLine { get; set; }

        public TranscriptWord()
        {
            PageAndLine = new List<Occurrence>();
        }

    }

    class Occurrence
    {
        public int CustomPageNumber { get; set; }
        public int CustomLineNumber { get; set; }

    }
}
