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
        private static string[] finalDeDupedWordList = null;

        private static Stopwatch applicationTime = new Stopwatch();

        private string currentWord = "";
        


        static void Main(string[] args)
        {


            Program WordIndex = new Program();
             //  WordIndex.processDocument();  // Read main document - Confirmed working.
            WordIndex.processDocument1();  // New routine - 03/31/2017 (after first prof read on 03/17/2017) 

            WordIndex.printWordIndex(); // Current print method - confirmed working- may need minor modifications

        }

        private void processDocument1()
        {
            applicationTime.Start(); //Start stopwatch
            //  Document document = app.Documents.Open(@"C:\User_Pradeep\Transcript2.doc", ReadOnly: true);
            Document document = app.Documents.Open(@"C:\User_Pradeep\Transcript3C.doc", ReadOnly: true); //test version

            //Split words into an array

            string firstParseString = "";
            string secondParseString = "";
            string thirdParseString = ""; // new changes
            
            firstParseString = document.Content.Text;
        
            string[] firstParseWordList = null;
            string[] secondParseWordList = null;
            string[] thirdParseWordList = null;

            //First parse-split, remove line-carrage, tabs after questions, tabs after answers, paranthesis,tab after ?, tab after period etc..etc

            firstParseWordList = firstParseString.Replace("\r", " ").Replace("Q\t", " ").Replace("A\t", " ").Replace("(", " ")
                      .Replace(")", " ").Replace("?\t", "  ").Replace("—\t", "  ").Replace("?", "  ")
                      .Replace(".\t", "  ").Split(' ');
            
           //Read after first-parse-split, remove tabs with leading digits

            for (int i = 0; i < firstParseWordList.Length; i++)
            {
                if (firstParseWordList[i].Trim().Equals(""))
                {
                    continue;
                }
                else
                {
                    if (i == firstParseWordList.Length - 1)
                    {
                        string tempOutputOne = Regex.Replace(firstParseWordList[i], @"^\d+\t", " "); //Remove "digits and tab" at the begining of the word - OK 
                        string tempOutputTwo = Regex.Replace(tempOutputOne, @"\d+\t", " "); //Remove "digits and tab" from anywhere in the word - OK     
                        secondParseString += tempOutputTwo;    
                    }
                    else
                    {
                        string tempOutputOne= Regex.Replace(firstParseWordList[i], @"^\d+\t", " "); //Remove "digits and tab" at the begining of the word - OK 
                        string tempOutputTwo = Regex.Replace(tempOutputOne, @"\d+\t", " "); //Remove "digits and tab" from anywhere in the word - OK       
                        secondParseString += tempOutputTwo + " ";        
                    }
                }
            }

            //Second parse-split, remove all remaining tabs and under-scores

            secondParseWordList = secondParseString.Replace("\t", " ").Replace("__", " ").Split(' ');

            //Read after second-parse-split, remove leading and trailling hyphens, remove leading and trailling dashes
            //Remove trailling periods, commas, colons and semicolons

            for (int i = 0; i < secondParseWordList.Length; i++)
            {
                if (secondParseWordList[i].Trim().Equals(""))
                {
                    continue;
                }
                else
                {
                    if (i == secondParseWordList.Length - 1)
                    {
                        string tempOutputOne = Regex.Replace(secondParseWordList[i], @"^\-+|\-+$", " "); //Remove leading and trailling hyphens
                        string tempOutputTwo = Regex.Replace(tempOutputOne, @"^\—+|\—+$", " "); //Remove leading and trailling dashes
                        string tempOutputThree = Regex.Replace(tempOutputTwo, @"\.+$|\,+$|\:+$|\;+$", " "); //Remove trailling periods, commas, colons, semicolons
                        
                        thirdParseString += tempOutputThree; 
                    }
                    else
                    {
                        
                        string tempOutputOne = Regex.Replace(secondParseWordList[i], @"^\-+|\-+$", " ");  
                        string tempOutputTwo = Regex.Replace(tempOutputOne, @"^\—+|\—+$", " ");
                        string tempOutputThree = Regex.Replace(tempOutputTwo, @"\.+$|\,+$|\:+$|\;+$", " ");
                        
                        thirdParseString += tempOutputThree + " ";
                    }
                }
                
            }

            thirdParseWordList = thirdParseString.Split(' ');

            finalDeDupedWordList = thirdParseWordList.Distinct().ToArray();

            document.Close();

            this.processDocument2(ref finalDeDupedWordList);
        }

        private void processDocument2(ref string[] finalDeDupedWordList)
        {
            Document document = app.Documents.Open(@"C:\User_Pradeep\Transcript3C.doc", ReadOnly: true); //test version

            document.Activate();

            HashSet<string> processedWordList = new HashSet<string>();

            string finalSearchWord = "";

            //Console.SetWindowSize();

            for (int i = 0; i < finalDeDupedWordList.Length; i++)
            {

                finalSearchWord = finalDeDupedWordList[i].Trim();

                //Words and sentences within double quotes(" ") should be identified, quotes should be removed in order to preserve correct print order (i.e. #'s $'s digits and actual words)  

                if (Regex.IsMatch(finalSearchWord, @"^[a-zA-Z0-9\$#]"))
                {

                }
                else
                {
                    //If the word is not all spaces and does not starts with one of the allowed charactors, then remove first position, could be a starting double quote or single quote 
                    if (finalSearchWord != "")
                    {
                        finalSearchWord = finalSearchWord.Remove(0, 1);
                    }
                   
                }

                //If the last position of the word in not one of the allowed charactors, then remove it, could be a closing double quote or single quote 
                if (finalSearchWord != "")
                {
                    if (Regex.IsMatch(finalSearchWord.Substring(finalSearchWord.Length - 1, 1), @"[a-zA-Z0-9\$#]"))
                    {

                    }
                    else
                    {

                        finalSearchWord = finalSearchWord.Remove(finalSearchWord.Length - 1, 1);

                    }
                }

                //Cleanup any spaces created by above process if any 
                finalSearchWord.Trim();

                if (finalSearchWord.Length > 2)
                {

                    string processedWord = processedWordList.FirstOrDefault(w => w == finalSearchWord); // 10/25

                    if (processedWord == null)
                    {

                        Console.WriteLine("Scanning transcript- processing word # " + i);

                        var CustomWord = new TranscriptWord();

                        // Not in the array, it is a new word so add to processed word list and start processing....

                        processedWordList.Add(finalSearchWord); 

                        CustomWord.Name = finalSearchWord;

                        int wordFoundFrequency = 0;

                        Range searchRange = document.Range(Start: document.Content.Start, End: document.Content.End); //Look for the word from start of the transcript to end

                        searchRange.Find.Forward = true; // 10/26 all range words
                        searchRange.Find.MatchCase = true; // 04/14/2016
                        searchRange.Find.Text = finalSearchWord;

                        currentWord = finalSearchWord;

                        searchRange.Find.Execute(MatchWholeWord: true);
                        int currentLineNumber = 0;
                        int currentPageNumber = 0;
                        int pageNumberOfTheWord = 0;
                        int lineNumberOfTheWord = 0;


                        while (searchRange.Find.Found)
                        {

                            Console.WriteLine("Looking for word : " + currentWord);

                            // Get current sentence being searched and extract the first word

                            string textOfTheSearchedRangeSentence = searchRange.Sentences.First.Text;
                            string firstWordOfTheSearchedSentence = textOfTheSearchedRangeSentence.Substring(0, finalSearchWord.Length);

                            //If sentence starts with a number, it usually is a question number, now if it is a number and matches the searched text, it definitely 
                            // cannot be a regular word, it got to be a question number, so ignore! 

                            if ((firstWordOfTheSearchedSentence == finalSearchWord) && (Regex.IsMatch(firstWordOfTheSearchedSentence, @"^[0-9]")))
                            {

                            }
                            else
                            {

                               // if (currentPageNumber != 1) // Page # 1 is cover page
                              //  {
                                    wordFoundFrequency++;

                                    currentPageNumber = searchRange.Information[WdInformation.wdActiveEndPageNumber]; 

                                    currentLineNumber = searchRange.Information[WdInformation.wdFirstCharacterLineNumber];


                                    //Check whether current word is repeating in the same page and line number, if not, create the "Occurrence" object. 

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
                               // }

                            }

                            searchRange.Find.Execute(MatchWholeWord: true);

                        }

                        CustomWord.Frequency = wordFoundFrequency;
                        CustomWordDirectory.Add(CustomWord);
                    }

                }
            }
            document.Close();
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
            bool isRealNumber = false;
            bool isNumberSign = false;
            bool isLetter = false;
            bool isCurrency = false; //04/14/2017
            bool numberSignPrinted = false;
            bool numericalDigitsPrinted = false;
            bool currencySignPrinted = false; //04/14/2017
            bool firstLetterPrinted = false; 

            foreach (TranscriptWord item in SCustomWordDirectory)
            {

                Console.WriteLine("Assembling index - processing word : " + item.Name );

                //Check first charactor of the word to determine type and identification
                string firstCharOfTheWord = item.Name.Substring(0, 1);
                // int valueOne = 0;

                //Assert whether first char of the current word is a number sign,$ sign,digit or letter 
                //then print it under appropriate label

                //Currency
                if (Regex.IsMatch(firstCharOfTheWord, @"^[$]"))
                {
                    isCurrency = true;
                    isLetter = false;
                    isRealNumber = false;
                    isNumberSign = false;
                }

                //Numerical Digit
                if (Regex.IsMatch(firstCharOfTheWord, @"^[0-9]"))
                {
                    isRealNumber = true;
                    isCurrency = false;
                    isLetter = false;
                    isNumberSign = false;
                }

                //Starts with number sign
                if (Regex.IsMatch(firstCharOfTheWord, @"^[#]"))
                {
                    isNumberSign = true;
                    isCurrency = false;
                    isLetter = false;
                    isRealNumber = false;  
                }

                ////Starts with double quotes
                //if (Regex.IsMatch(firstCharOfTheWord, @"["]"))
                //{
                //    isNumberSign = true;
                //    isCurrency = false;
                //    isLetter = false;
                //    isRealNumber = false;
                //}

                // Alphabetical letter
                if (Regex.IsMatch(firstCharOfTheWord, @"^[a-zA-Z]"))
                {
                    isLetter = true;
                    isCurrency = false;
                    isRealNumber = false;
                    isNumberSign = false;

                    if (indexAlphabetLabel.ToLower() == firstCharOfTheWord.ToLower())
                    {
                        firstLetterPrinted = true;
                    }
                    else
                    {
                        firstLetterPrinted = false;
                    }

                }
               
                // First char is $ sign, dollar sign is not printed.
                if ((isCurrency) && (!currencySignPrinted)) 
                {
                    indexAlphabetLabel = "$";

                    Paragraph labelParagraph = indexDoc.Paragraphs.Add();
                    labelParagraph.Range.Font.Size = 15;
                    labelParagraph.Range.Font.Bold = 1;
                    labelParagraph.Range.Text = " " + indexAlphabetLabel + " " + "\r\n";
                    currencySignPrinted = true;
                }

                // Real digits, label 0-9 not printed.
                if ((isRealNumber) && (!numericalDigitsPrinted))
                {
                    indexAlphabetLabel = "0-9";

                    Paragraph labelParagraph = indexDoc.Paragraphs.Add();
                    labelParagraph.Range.Font.Size = 15;
                    labelParagraph.Range.Font.Bold = 1;
                    labelParagraph.Range.Text = " " + indexAlphabetLabel + " " + "\r\n";
                    numericalDigitsPrinted = true;
                }

                // First char is number sign, number sign is not printed.
                if ((isNumberSign) && (!numberSignPrinted)) 
                {
                    indexAlphabetLabel = "#";

                    Paragraph labelParagraph = indexDoc.Paragraphs.Add();
                    labelParagraph.Range.Font.Size = 15;
                    labelParagraph.Range.Font.Bold = 1;
                    labelParagraph.Range.Text = " " + indexAlphabetLabel + " " +  "\r\n";
                    numberSignPrinted = true;
                }

                //First char is a letter, alphabet label is not printed. 
                if ((isLetter) && (!firstLetterPrinted)) 
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
