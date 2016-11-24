using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using ReadWord;

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
            WordIndex.processDocument();

           // WordIndex.printWordIndex(); // Old print method
           //WordIndex.createIndexTable(); // Index Table POC

           WordIndex.createWrodIndexTable(); // Real Index
            //New routine needed - 11/23/2016


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
                if (colNumber == 1)
                {
                    actualRowCount++;

                }

                rowNumber++;
                tableRowCount++;

                cellRange = myTable.Cell(rowNumber, colNumber).Range;
                
                cellRange.Font.Size = 10;
                cellRange.Font.Bold = 1;
                cellRange.Text = item.Name + " [" + item.Frequency + "]";

                

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


                if ((tableRowCount == 31) && (colNumber < 3))
                {
                   
                    rowNumber = (actualRowCount - 31);
                    colNumber++;
                    tableRowCount = 0;

                }

                if ((tableRowCount == 31) && (colNumber == 3))
                {

                    rowNumber = actualRowCount;
                    tableRowCount = 0;
                    colNumber = 1;
                }


                foreach (Occurrence step in item.PageAndLine)
                {       

                    columnCountPageAndLine++; // Row count for Page and Line numbers

                    if (columnCountPageAndLine == 2)
                    {

                        if (colNumber == 1)
                        {
                            actualRowCount++;

                        }

                        rowNumber++;
                        tableRowCount++;

                        //  cellRange.ParagraphFormat.SpaceBefore = 0;

                        myTable.Cell(rowNumber, colNumber).TopPadding = 0.00f;
                        



                        cellRange = myTable.Cell(rowNumber, colNumber).Range;
                       
                        cellRange.Font.Size = 7;
                        cellRange.Font.Bold = 0;

                        //cellRange.ParagraphFormat.SpaceBefore = 0.00f;

                        Cell cell = myTable.Cell(rowNumber, colNumber);
                       
                        cell.Range.ParagraphFormat.SpaceBefore = 0.00f;

                        cellRange.Text = "[P" + columnOnePageNumber + ":" + "L" + columnOneLineNumber + "] [P" + step.CustomPageNumber + ":" + "L" + step.CustomLineNumber + "]";
                        

                       // cellRange.ParagraphFormat.SpaceAfter = 0.00f;
                        cell.Range.ParagraphFormat.SpaceAfter = 0.00f;

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

                        if ((tableRowCount == 31) && (colNumber < 3))
                        {
                            
                            rowNumber = (actualRowCount - 31);
                            colNumber++;
                            tableRowCount = 0;

                        }

                        if ((tableRowCount == 31) && (colNumber == 3))
                        {

                            rowNumber = actualRowCount;
                            tableRowCount = 0;
                            colNumber = 1;
                        }

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

                    if (colNumber == 1)
                    {
                        actualRowCount++;

                    }

                    rowNumber++;
                    tableRowCount++;

                    cellRange = myTable.Cell(rowNumber, colNumber).Range;
                    cellRange.Font.Size = 7;
                    cellRange.Font.Bold = 0;
                    cellRange.Text = "[P" + columnOnePageNumber + ":" + "L" + columnOneLineNumber + "]" ;

                   // cellRange.ParagraphFormat.SpaceAfter = 0;

                    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;


                    if ((tableRowCount == 31) && (colNumber < 3))
                    {
                        
                        rowNumber = (actualRowCount - 31);
                        colNumber++;
                        tableRowCount = 0;

                    }

                    if ((tableRowCount == 31) && (colNumber == 3))
                    {

                        rowNumber = actualRowCount;
                        tableRowCount = 0;
                        colNumber = 1;
                    }

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

            for (int i = 1; i <= 86; i++)
            {

                if (colNumber == 1)
                {
                    actualRowCount++;

                }

                rowNumber++;
                tableRowCount++;

                myTable.Cell(rowNumber, colNumber).BottomPadding = 0.00f;
                myTable.Cell(rowNumber, colNumber).TopPadding = 0.00f;

                cellRange = myTable.Cell(rowNumber, colNumber).Range;

                cellRange.ParagraphFormat.SpaceBefore = 0.0f;

                cellRange.Text = "The value is :" + i + " column:" + colNumber;

                cellRange.ParagraphFormat.SpaceAfter = 0.0f;

                myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                if ((actualRowCount == 9) || (actualRowCount == 21) || (actualRowCount == 30))
                {
                    cellRange.Font.Size = 10;
                    cellRange.Font.Bold = 1;
                    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    myTable.Cell(rowNumber, colNumber).Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                }


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

            int totalElements = SCustomWordDirectory.Count;


            Console.WriteLine("Total number of elements in the List : " + totalElements);

            Document indexDoc = app.Documents.Add();
            Range indexRange = indexDoc.Range();
            indexDoc.Activate();
            // indexDoc.Range().Select();
            indexRange.Select();

            int mainElements = 0;
            int subElements = 0;

            foreach (TranscriptWord item in SCustomWordDirectory)
            {
                Paragraph para1 = indexDoc.Paragraphs.Add();
                para1.Range.Font.Size = 10;
                para1.Range.Font.Bold = 1;

               // para1.Range.
                //para1.Range.Text = item.Name + " [" + item.Frequency + "]" + "\r\n"; 10/26

                para1.Range.Text = item.Name + " [" + item.Frequency + "]" + "\r\n";  // 10/26

                int columnCount = 0;
                int columnOnePageNumber = 0;
                int columnOneLineNumber = 0;

                mainElements++;

                foreach (Occurrence step in item.PageAndLine)
                {
                    columnCount++; // Actually row count

                    subElements++;

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

            }

            Console.WriteLine("Number of Words : " + mainElements + " and sub items : " + subElements);

            indexDoc.Save();
            indexDoc.Close();
            app.Quit();
        }

        private void processDocument()
        {

            Document document = app.Documents.Open(@"C:\User_Pradeep\Transcript3.doc", ReadOnly: true);
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
