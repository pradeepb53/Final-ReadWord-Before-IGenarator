using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using ReadWord;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace ReadWord
{
    class Program
    {
        private Application app = new Application();



        private static List<TranscriptWord> CustomWordDirectory = new List<TranscriptWord>();

        private static Stopwatch applicationTime = new Stopwatch();

        private string currentWord = "";
        


        static void Main(string[] args)
        {


            Program WordIndex = new Program();
               WordIndex.processDocument();  // Read main document

             WordIndex.printWordIndex(); // Old print method

        }


        private void printWordIndex()
        {
            List<TranscriptWord> SCustomWordDirectory = CustomWordDirectory.OrderBy(o => o.Name).ToList();

            Document indexDoc = app.Documents.Add();
            Range indexRange = indexDoc.Range();

            //Set columns 
            indexDoc.PageSetup.TextColumns.SetCount(5);

            indexDoc.Activate();
            indexRange.Select();

            //Temp vars
            string indexAlphabetLabel = ""; 

            //Logical vars
            bool isNumber = false; 
            bool isLetter = false; 
            bool numberSignPrinted = false; 
            bool firstLetterPrinted = false; 

            foreach (TranscriptWord item in SCustomWordDirectory)
            {

                Console.WriteLine("Assembling index - processing word : " + item.Name );

                //Check first charactor of the word to determine type and identification
                string firstCharOfTheWord = item.Name.Substring(0, 1);
                int valueOne = 0;
                
                //Assert whether first char of the current word is a number or later, then print it under appropriate
                //alphabet label 
                bool parsed = Int32.TryParse(firstCharOfTheWord, out valueOne);

                if (parsed)
                {
                    isNumber = true;
                    isLetter = false;
                   
                    if (indexAlphabetLabel == "#")
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

                        if (indexAlphabetLabel == firstCharOfTheWord)
                        {
                            firstLetterPrinted = true;
                        }else
                        {
                            firstLetterPrinted = false;
                        }
                    }

                }
                  
                if((isNumber) && (!numberSignPrinted)) // First char is a number, number sign is not printed.
                {
                    indexAlphabetLabel = "#";

                    Paragraph labelParagraph = indexDoc.Paragraphs.Add();
                    labelParagraph.Range.Font.Size = 15;
                    labelParagraph.Range.Font.Bold = 1;
                    labelParagraph.Range.Text = " " + indexAlphabetLabel + " " +  "\r\n";
                    numberSignPrinted = true;
                }

                if ((isLetter) && (!firstLetterPrinted)) //First char is a letter, alphabet label is not printed. 
                {
                    indexAlphabetLabel = firstCharOfTheWord;

                    Paragraph labelParagraph = indexDoc.Paragraphs.Add();
                    labelParagraph.Range.ParagraphFormat.SpaceBefore = 0;
                    labelParagraph.Range.Font.Size = 15;
                    labelParagraph.Range.Font.Bold = 1;
                    labelParagraph.Range.Text = "- " + indexAlphabetLabel.ToUpper() + " -" + "\r\n";
                    firstLetterPrinted = true;
                }

                Paragraph currentWordParagraph = indexDoc.Paragraphs.Add();
                currentWordParagraph.Range.Font.Size = 10;
                currentWordParagraph.Range.Font.Bold = 1;
 
                currentWordParagraph.Range.Text = item.Name + " [" + item.Frequency + "]" + "\r\n";  

                int columnsPerRowCount = 0;
                int columnOnePageNumber = 0;
                int columnOneLineNumber = 0;

                foreach (Occurrence step in item.PageAndLine)
                {
                    Console.WriteLine("Checking frequency of : " + item.Name);

                    columnsPerRowCount++; 

                    if (columnsPerRowCount == 2)
                    {
                        Paragraph pageAndLineParagraph = indexDoc.Paragraphs.Add();
                        pageAndLineParagraph.Range.Font.Size = 7;
                        pageAndLineParagraph.Range.Font.Bold = 0;

                        pageAndLineParagraph.Range.Text = "[P" + columnOnePageNumber + ":" + "L" + columnOneLineNumber + "] [P" + step.CustomPageNumber + ":" + "L" + step.CustomLineNumber + "]" + "\r\n";
                        pageAndLineParagraph.Range.ParagraphFormat.SpaceAfter = 0;

                        columnsPerRowCount = 0;
                        columnOnePageNumber = 0;
                        columnOneLineNumber = 0;
                    }
                    else
                    {
                         columnOnePageNumber = step.CustomPageNumber;
                         columnOneLineNumber = step.CustomLineNumber;

                    }
                    
                }

                //If columnsPerRowCount is 1, then print one line and reset counter 
                if (columnsPerRowCount == 1)
                {
                    Paragraph para3 = indexDoc.Paragraphs.Add();
                    para3.Range.Font.Size = 7;
                    para3.Range.Font.Bold = 0;
                    para3.Range.Text = "[P" + columnOnePageNumber + ":" + "L" + columnOneLineNumber + "]" + "\r\n";
                    para3.Range.ParagraphFormat.SpaceAfter = 0;

                    columnsPerRowCount = 0;
                    columnOnePageNumber = 0;
                    columnOneLineNumber = 0;
                }

            }

            applicationTime.Stop();
            TimeSpan elapsedTime = applicationTime.Elapsed;

            Console.WriteLine("Index generation completed!- Duration - {0} hour(s):{1} minute(s):{2} second(s)",
                elapsedTime.Hours, elapsedTime.Minutes, elapsedTime.Seconds);
            
            indexDoc.Save();
            indexDoc.Close();
            app.Quit();
        }

        private void processDocument()
        {

            //  Document document = app.Documents.Open(@"C:\User_Pradeep\Transcript2.doc", ReadOnly: true);
            
           Document document = app.Documents.Open(@"C:\User_Pradeep\Transcript3C.doc", ReadOnly: true); //test version

            //--------------------------------------------------------------------->New changes to be implemented <-------------------------------------------

            //var wordEncode = Microsoft.Office.Core.MsoEncoding.msoEncodingUSASCII;
            ////Split words into an array
            //string str = "";
            //string newStr = "";
            //string finalStr = "";
            //str = document.Content.Text;
            //string[] words = null;
            //string[] newWords = null;
            //words = str.Replace("\r", "").Replace("\tQ", "").Replace("\tA", "").Replace("(", "").Replace(")", "").Replace(":", "").Split(' ');
            //for (int i = 0; i < words.Length; i++)
            //{
            //    if (words[i].Trim().Equals(""))
            //    {
            //        continue;
            //    }
            //    else
            //    {
            //        if (i == words.Length - 1)
            //        {
            //            newStr += words[i];
            //        }
            //        else
            //        {
            //            newStr += words[i] + " ";
            //        }
            //    }
            //}
            //newWords = newStr.Split(' ');

            //for (int i = 0; i < newWords.Length; i++)
            //{
            //    finalStr = newWords[i];
            //}
            //<--------------------------------------------------------------------
            document.Activate();

            applicationTime.Start(); //Start stopwatch

            HashSet<string> processedWordList = new HashSet<string>(); // 10/25
           

            string textOftheWord = "";
            bool validWord = false;
           

            //Console.SetWindowSize();

            for (int i = 1; i < document.Words.Count; i++)
            {

                textOftheWord = document.Words[i].Text.ToString().Trim().ToLower(); // 10/26

                //if ((textOftheWord.Length > 3) && (textOftheWord.Contains("_________") == false) && (textOftheWord != "-")) // 10/26 - 03/28/2017

                //Check length of the word is > 3 and word contains letters, numbers or $,#,* only
                if ((textOftheWord.Length > 2) && (Regex.IsMatch(textOftheWord, @"^[a-zA-Z0-9\$#*]+$"))) // 03/28/2017

                    {
                    validWord = true;      
                    
                    string processedWord = processedWordList.FirstOrDefault(w => w == textOftheWord); // 10/25
    
                        if (processedWord == null) 
                        {

                        //If word is all numbers and equal to number at the begining of the sentence then it is a question number, ignore!
                        if (Regex.IsMatch(textOftheWord, @"^[0-9]")) 
                        {
                            Range sentenceRange = document.Range(Start: document.Words[i].Start, End: document.Content.End); // 03/27/2017
                            string textOfTheCurrentSentence = sentenceRange.Sentences.First.Text; // 03/27/2017

                            string firstWordOfTheCurrentSentence = textOfTheCurrentSentence.Substring(0, textOftheWord.Length);
                            int valueInTheWord = 0;

                            bool isNumber = Int32.TryParse(firstWordOfTheCurrentSentence, out valueInTheWord); // 10/22

                            if ((isNumber) && (textOftheWord == firstWordOfTheCurrentSentence))
                            {
                                validWord = false;
                            }
                        }

                        if (validWord)   //((!parsed) )  03/28/2017
                        {

                            Range currentRange = document.Range(Start: document.Words[i].Start, End: document.Words[i].End); // 10/22

                            int pageNumberOfTheCurrentRange = currentRange.Information[WdInformation.wdActiveEndPageNumber]; // 10/22


                            if ((pageNumberOfTheCurrentRange > 1) ) // 10/26
                            {

                                Console.WriteLine("Scanning transcript- processing word # " + i + " -  In page # " + pageNumberOfTheCurrentRange);

                                var CustomWord = new TranscriptWord();

                                // Not in the array, it is a new word

                                processedWordList.Add(textOftheWord); // 10/24

                                        CustomWord.Name = textOftheWord;

                                    int wordFoundFrequency = 0;

                                    Range searchRange = document.Range(Start: document.Words[i].Start, End: document.Content.End); // 10/23

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

                                    Console.WriteLine("Looking for word : " + currentWord);

                                    // Get currently searched sentence and extract the first word
                                    string textOfTheSearchedRangeSentence = searchRange.Sentences.First.Text;
                                    string firstWordOfTheSearchedSentence = textOfTheSearchedRangeSentence.Substring(0, textOftheWord.Length);

                                    //If begining of the sentence is a number, it usually is a question number, now if it is a number and matches the searched text, it definitely 
                                    // cannot be a regular word, it got to be a question number, so ignore! 
                                    if ((firstWordOfTheSearchedSentence == textOftheWord) && (Regex.IsMatch(firstWordOfTheSearchedSentence, @"^[0-9]")))
                                    {
                                       
                                    }
                                    else
                                    {
                                        wordFoundFrequency++;

                                        currentPageNumber = searchRange.Information[WdInformation.wdActiveEndPageNumber];

                                        currentLineNumber = searchRange.Information[WdInformation.wdFirstCharacterLineNumber];

                                        if (currentPageNumber != 1) // Page # 1 is cover page
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
