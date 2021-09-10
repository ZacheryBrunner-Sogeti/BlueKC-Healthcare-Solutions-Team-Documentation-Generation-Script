/*
 * Class: CreateWordDocument.cs
 * Purpose: Creates a Microsoft Word Document (.docx) for each ADOObject
 * Developer: Zachery Q. Brunner
 */
using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace Script2_ADO_to_Word
{
    /// <summary>
    /// Generates Microsoft Word Documents for each ADOObject in ADOObjList
    /// </summary>
    public class CreateWordDocument
    {
        /// <summary>
        /// Global ADOObject that the word document is built from
        /// </summary>
        private List<ADOObject> ADOObjList { get; set; }

        /// <summary>
        /// Public constructor used to populate the ADOObject global variable
        /// </summary>
        /// <param name="ADOObjList">Populated with all test cases</param>
        public CreateWordDocument(List<ADOObject> ADOObjList)
        {
            this.ADOObjList = ADOObjList;
        }

        /// <summary>
        /// Entry point for document creation
        ///         Will generate a word document for each ADO Object
        /// </summary>
        public void GenerateDocuments()
        {
            foreach (ADOObject ADOObj in ADOObjList) { CreateDocument(ADOObj); }
        }

        /// <summary>
        /// Creates Microsoft Word Documents with two help functions
        ///     BuildFirstPage(), BuildTestSteps()
        /// </summary>
        /// <param name="ADOObj">The ADOObject have the Microsoft Word Document built</param>
        private void CreateDocument(ADOObject ADOObj)
        {
            try
            {
                //Create an instance for word app  
                Word.Application winword = new Word.Application();

                //Set animation status for word application  
                winword.ShowAnimation = false;

                //Set status for word application is to be visible or not.  
                winword.Visible = false;

                //Create a missing variable for missing value  
                object missing = System.Reflection.Missing.Value;

                //Create a new document  
                Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                //adding text to document  
                document.Content.SetRange(0, 0);

                //Builds the document using following methods
                BuildFirstPage(document, ADOObj, missing);
                document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                BuildTestSteps(document, ADOObj, missing);

                //Save the document  
                object filename = @"C:\Users\brunnez\OneDrive - Blue Cross Blue Shield of Kansas City\Documents\" + ADOObj.ADO_UNIQ_ID.ToString() +
                                                            " " + ADOObj.Title.Trim() + ".docx";
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Used to assign header #1 properties
        /// </summary>
        /// <param name="WR">Preupdate properties</param>
        /// <returns>Updated properties</returns>
        private Word.Range AssignHeader1(Word.Range WR)
        {
            WR.Font.Bold = 1;
            WR.Font.Size = 40;
            WR.Font.ColorIndex = Word.WdColorIndex.wdBlack;
            return WR;
        }

        /// <summary>
        /// Used to assign header #2 properties
        /// </summary>
        /// <param name="WR">Preupdate properties</param>
        /// <returns>Updated properties</returns>
        private Word.Range AssignHeader2(Word.Range WR)
        {
            WR.Font.Bold = 1;
            WR.Font.Size = 20;
            WR.Font.ColorIndex = Word.WdColorIndex.wdBlack;
            return WR;
        }

        /// <summary>
        /// Used to assign header #3 properties
        /// </summary>
        /// <param name="WR">Preupdate properties</param>
        /// <returns>Updated properties</returns>
        private Word.Range AssignHeader3(Word.Range WR)
        {
            WR.Font.Bold = 0;
            WR.Font.Size = 12;
            WR.Font.ColorIndex = Word.WdColorIndex.wdBlack;
            return WR;
        }

        /// <summary>
        /// Builds the first page (informational page) of the word document
        /// </summary>
        /// <param name="document">The Microsoft Word document being built</param>
        /// <param name="ADOObj">The ADOObject having its Microsoft Word Document built</param>
        /// <param name="missing">System.Reflection.Missing.Value</param>
        private void BuildFirstPage(Word.Document document, ADOObject ADOObj, object missing)
        {
            //Title
            Word.Paragraph titleParagraph = document.Content.Paragraphs.Add(ref missing);
            Word.Range titleRange = titleParagraph.Range;
            titleRange = AssignHeader1(titleRange);
            titleRange.Text = ADOObj.ADO_UNIQ_ID + ": " + ADOObj.Title;
            titleRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            titleRange.InsertParagraphAfter();

            //Testing Link
            Word.Paragraph testingLink = document.Content.Paragraphs.Add(ref missing);
            Word.Range testingLinkRange = testingLink.Range;
            testingLinkRange = AssignHeader2(testingLinkRange);
            testingLinkRange.Text = "Testing Link:";
            testingLinkRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            testingLinkRange.InsertParagraphAfter();

            //Testing Link Bullet Point
            Word.Paragraph testingLinkBullet = document.Content.Paragraphs.Add(ref missing);
            testingLinkBullet.Range.ListFormat.ApplyBulletDefault();
            testingLinkBullet.Range.InsertBefore("REQUIRES UPDATE");
            testingLinkBullet.Range.InsertParagraphAfter();

            //Tester
            Word.Paragraph tester = document.Content.Paragraphs.Add(ref missing);
            Word.Range testerRange = tester.Range;
            testerRange = AssignHeader2(testerRange);
            testerRange.Text = "Tester:";
            testerRange.InsertParagraphAfter();

            //Tester Bullet Point
            Word.Paragraph testerLinkBullet = document.Content.Paragraphs.Add(ref missing);
            testerLinkBullet.Range.ListFormat.ApplyBulletDefault();
            testerLinkBullet.Range.InsertBefore(ADOObj.Tester);
            testerLinkBullet.Range.InsertParagraphAfter();

            //Additional Information
            Word.Paragraph AdditionalInfo = document.Content.Paragraphs.Add(ref missing);
            Word.Range AdditionalInfoRange = AdditionalInfo.Range;
            AdditionalInfoRange = AssignHeader2(AdditionalInfoRange);
            AdditionalInfoRange.Text = "Additional Information:";
            AdditionalInfoRange.InsertParagraphAfter();

            //Additional Information Bullet Point
            Word.Paragraph additionalInfoBullet = document.Content.Paragraphs.Add(ref missing);
            additionalInfoBullet.Range.ListFormat.ApplyBulletDefault();
            additionalInfoBullet.Range.InsertBefore("TEST STATUS: REQUIRES UPDATE & HIGHLIGHT");
            additionalInfoBullet.Range.InsertParagraphAfter();

            //Test Case Links
            Word.Paragraph TestCaseLinks = document.Content.Paragraphs.Add(ref missing);
            Word.Range TestCaseLinksRange = TestCaseLinks.Range;
            TestCaseLinksRange = AssignHeader2(TestCaseLinksRange);
            TestCaseLinksRange.Text = "Test Case Link(s):";
            TestCaseLinksRange.InsertParagraphAfter();

            //Test Case Links Bullet Point
            object Address = "https://bluekc.visualstudio.com/BlueKC/_workitems/edit/" + ADOObj.ADO_UNIQ_ID.ToString();
            Word.Paragraph testCaseLinksBullet = document.Content.Paragraphs.Add(ref missing);
            testCaseLinksBullet.Range.ListFormat.ApplyBulletDefault();
            testCaseLinksBullet.Range.Hyperlinks.Add(testCaseLinksBullet.Range, ref Address, ref missing, ref missing, ref missing);
            testCaseLinksBullet.Range.InsertParagraphAfter();
        }

        /// <summary>
        /// Builds the test steps of the word document
        /// </summary>
        /// <param name="document">The Microsoft Word document being built</param>
        /// <param name="ADOObj">The ADOObject having its Microsoft Word Document built</param>
        /// <param name="missing">System.Reflection.Missing.Value</param>
        private void BuildTestSteps(Word.Document document, ADOObject ADOObj, object missing)
        {
            Word.Paragraph ProcessingSteps = document.Content.Paragraphs.Add(ref missing);
            Word.Range ProcessingStepsRange = ProcessingSteps.Range;
            ProcessingStepsRange = AssignHeader1(ProcessingStepsRange);
            ProcessingStepsRange.Text = "Processing Steps";
            ProcessingStepsRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            ProcessingStepsRange.InsertParagraphAfter();

            int stepNumber = 1;
            foreach (string s in ADOObj.StepAction)
            {
                //Step N:
                Word.Paragraph stepParagraph = document.Content.Paragraphs.Add(ref missing);
                Word.Range stepParagraphRange = stepParagraph.Range;
                stepParagraphRange = AssignHeader2(stepParagraphRange);
                stepParagraphRange.Text = "Step " + stepNumber + ":";
                stepParagraphRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                stepParagraphRange.InsertParagraphAfter();

                //Step Action: Some string
                //Step Expected: Some string
                Word.Paragraph paragraph = document.Content.Paragraphs.Add(ref missing);
                Word.Range paragraphRange = paragraph.Range;
                paragraphRange = AssignHeader3(paragraphRange);
                paragraphRange.Text = "Step Action: " + s + "\nStep Expected: " + ADOObj.StepExpected[stepNumber - 1];
                paragraphRange.InsertParagraphAfter();
                paragraphRange.InsertParagraphAfter();
                stepNumber++;
            }
        }
    }
}