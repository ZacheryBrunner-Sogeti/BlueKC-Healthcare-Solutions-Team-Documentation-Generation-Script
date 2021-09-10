/*
 * Class: ReadExcelDocument.cs
 * Purpose: Handles the conversion of the excel document to ADOObjects
 * Developer: Zachery Q. Brunner
 */
using System;
using System.IO;
using System.Collections.Generic;

namespace Script2_ADO_to_Word
{

    /* Excel File Layout
     * ID | Work Item | Title | Test Step | Step Action | Step Expected | Area Path | Assigned To | State
     *
     * int "ID"                 A   0
     * string "Work Item"       B   1
     * string "Title"           C   2
     * string "Test Step"       D   3
     * string "Step Action"     E   4
     * string "Step Expected"   F   5
     * string "Area Path"       G   6
     * string "Assigned To"     H   7
     * string "State"           I   8
     */

    public class ReadExcelDocument
    {
        /// <summary>
        /// Performs the transformation between Excel document and ADOObjects that will be turned into Word Documents
        /// </summary>
        /// <returns>List of ADOObjects ready for Excel document transformation</returns>
        public static List<ADOObject> GetExcelFile()
        {
            /* Open up file reader. It is important to have the document named the proper way
             *          Inside the natural "Downloads" folder - "Tester_Excel_File"
             */
            try
            {
                int ADO_UNIQ_ID = 0;
                string currentLine = null;
                string ADO_Title = "";
                string ADO_Tester = "";
                List<string> stepAction = new List<string>();
                List<string> stepExpected = new List<string>();
                List<ADOObject> ADOObjList = new List<ADOObject>();

                /* Initialize file read
                 *      Will read the file until .Close()
                 */
                using (StreamReader sr = new StreamReader(@"\Users\brunnez\Downloads\Tester_Excel_File.csv"))
                {
                    /* Read the first line of the file
                     *          Get the Header out of the way
                     */
                    if ((currentLine = sr.ReadLine()) == null)
                    {
                        throw new OutOfBoundsException("First Row Blank - Please check Excel file for errors - Shutting down");
                    }

                    /* Iterates the rest of the file 
                     *      Builds ADOObjects for later use
                     */
                    while ((currentLine = sr.ReadLine()) != null)
                    {
                        string charRemovedLine = currentLine.Replace("\"", "");
                        string[] information = charRemovedLine.Split(',');

                        if (information.Length != 4)
                        {
                            ///IF information[0] is NOT empty
                            if (!information[0].Equals(""))
                            {
                                ADOObjList.Add(new ADOObject(ADO_UNIQ_ID, ADO_Title, ADO_Tester, stepAction, stepExpected));
                                stepAction = new List<string>();
                                stepExpected = new List<string>();
                                ADO_UNIQ_ID = int.Parse(information[0].ToString().Trim());
                                ADO_Title = information[2];
                                ADO_Tester = information[7];
                            }
                            else
                            {
                                stepAction.Add(information[4]);

                                /* Not all steps will have an expected result
                                 *      This statement ensure that the program does not break because of it
                                 */
                                if (information.Length == 5)
                                {
                                    currentLine = sr.ReadLine();
                                    charRemovedLine = currentLine.Replace("\"", "");
                                    information = charRemovedLine.Split(',');
                                    stepExpected.Add(information[1]);
                                    sr.ReadLine();
                                }
                                else
                                {
                                    if (information[5].Equals(""))
                                        stepExpected.Add("No Expected Result Entered in ADO");
                                    else
                                        stepExpected.Add(information[5]);
                                }
                            }
                        }
                    }
                    ADOObjList.Add(new ADOObject(ADO_UNIQ_ID, ADO_Title, ADO_Tester, stepAction, stepExpected));
                    ADOObjList.RemoveAt(0);
                    sr.Close();
                    return ADOObjList;
                }
                throw new Exception("Unknown Exit");
            } 
            catch (FileNotFoundException FNFE)
            {
                Console.WriteLine("Please ensure file is named correctly - 'Tester_Excel_File'\n\t\tFile Type: Microsoft Excel Comma Separated Values File\n\n\n\n" + FNFE.ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            throw new Exception("Unknown Error - Please consult Developer");
        }
    }
}