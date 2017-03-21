using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace FVRs_IG
{
    class IndexCore
    {
        private Application app = new Application();

        private List<TranscriptWord> WordIndexDictionary = new List<TranscriptWord>();
        public IndexCore()
        {

        }
        public void processTranscript(string fileName, string[] excludedWordList)
        {
            Document document = app.Documents.Open(fileName, ReadOnly: true);

            document.Activate();
            int totalWordCount = document.Words.Count;
            int wordCountArrayLength = 0;
            string[] processedWordList = new string[totalWordCount];

            String textOftheWord = "";

            for (int i = 1; i < document.Words.Count; i++)
            {

                textOftheWord = document.Words[i].Text.ToString().Trim().ToLower();

                if (textOftheWord.Length > 2)
                {

                    //Check whether word is already processed?

                    int wordPosition = Array.IndexOf(processedWordList, textOftheWord);

                    //Check whether word is in the excluded list?
                    
                    int wordInExcludeList = Array.IndexOf(excludedWordList, textOftheWord);

                    var CustomWord = new TranscriptWord();

                    // Not in the array, it is a new word; not in the excluded list so it will be in the index

                    if ((wordPosition < 0 ) && ( wordInExcludeList < 0))
                    {
                        //Add new word to processed list of words

                        if (wordCountArrayLength <= totalWordCount)
                        {
                            processedWordList[wordCountArrayLength] = textOftheWord;
                            wordCountArrayLength++;

                            // Update CustomWord name property
                            CustomWord.Name = textOftheWord;

                        }

                        int wordFoundFrequency = 0;
                        Range fullRange = document.Content;
                        fullRange.Find.Forward = true;
                        fullRange.Find.Text = textOftheWord;
                        fullRange.Find.Execute(MatchWholeWord: true);
                        int currentLineNumber = 0;
                        int currentPageNumber = 0;
                        int pageNumberOfTheWord = 0;
                        int lineNumberOfTheWord = 0;


                        while (fullRange.Find.Found)
                        {
                            wordFoundFrequency++;

                            //Total papagraphs in the documnet

                            int paraCountTotal = document.Paragraphs.Count;
                            //Paragraph count in the selected range = Last paragraph read 
                            int totalParagraphsRead = app.ActiveDocument.Range(0, fullRange.End).Paragraphs.Count;

                            Range rangeFirstSentence = fullRange.Sentences.First;

                            //Text of the current sentence -- Entire Line
                            string textOfTheCurrentSentence = rangeFirstSentence.Text;

                            int lineNumberOfTheCurrentSentence = getLineNumberOfTheWord(textOfTheCurrentSentence);

                            if (lineNumberOfTheCurrentSentence == 0)
                            {
                                var currentParagraph = fullRange.Paragraphs.Last; // If no line number found fir the text, then get the entire paragraph. 
                                Range rangeCurrentParagraph = currentParagraph.Range;

                                //Text of the current paragraph
                                string textOfTheCurrentParagraph = rangeCurrentParagraph.Text;  // Line number found at the begining of the paragraph

                                int lineNumberOfTheCurrentParagraph = getLineNumberOfTheWord(textOfTheCurrentParagraph);

                                if (lineNumberOfTheCurrentParagraph == 0)
                                {
                                    var lastParagraph = document.Paragraphs[totalParagraphsRead - 1];
                                    Range rangeLastParagraph = lastParagraph.Range;

                                    //Text of the last paragraph
                                    string textOfTheLastParagraph = rangeLastParagraph.Text;  // Line number found at the begining of the First paragraph
                                    int lineNumberOfTheLastParagraph = getLineNumberOfTheWord(textOfTheLastParagraph);

                                    //Assume line number of the searced word is found here, if not ?? another logic needed!!!
                                    currentLineNumber = lineNumberOfTheLastParagraph;
                                }
                                else
                                {
                                    currentLineNumber = lineNumberOfTheCurrentParagraph;
                                }
                            }
                            else
                            {
                                currentLineNumber = lineNumberOfTheCurrentSentence;
                            }

                            currentPageNumber = fullRange.Information[WdInformation.wdActiveEndPageNumber];

                            //Check whether current word is repeating in the same page and line number, if not, create "Occurrence" object. 
                            //Only one Occurrence object is needed even if current word is repeated more than once in the same line number 
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


                            fullRange.Find.Execute(MatchWholeWord: true);

                        }

                        CustomWord.Frequency = wordFoundFrequency;
                        WordIndexDictionary.Add(CustomWord);

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

        public void printWordIndex()
        {
            List<TranscriptWord> IndexWordDirectory = WordIndexDictionary.OrderBy(o => o.Name).ToList();

            Document indexDoc = app.Documents.Add();
            Range indexRange = indexDoc.Range();
            indexDoc.Activate();
            indexRange.Select();

            foreach (TranscriptWord item in IndexWordDirectory)
            {
                Paragraph para1 = indexDoc.Paragraphs.Add();
                para1.Range.Font.Size = 10;
                para1.Range.Font.Bold = 1;

                para1.Range.Text = item.Name + " [" + item.Frequency + "]" + "\r\n";

                foreach (Occurrence step in item.PageAndLine)
                {
                    Paragraph para2 = indexDoc.Paragraphs.Add();
                    para2.Range.Font.Size = 7;
                    para2.Range.Font.Bold = 0;

                    para2.Range.Text = "[P" + step.CustomPageNumber + ":" + "L" + step.CustomLineNumber + "]" + "\r\n";
                    para2.Range.ParagraphFormat.SpaceAfter = 0;

                }

            }
             
            
            try
            {
                indexDoc.Save();
                indexDoc.Close();
                app.Quit();
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Console.WriteLine("Error occured, please contact IT bitch!! : " + e);
                
            }
            finally
            {         
                app.Quit();
            }
        }
    }
}
