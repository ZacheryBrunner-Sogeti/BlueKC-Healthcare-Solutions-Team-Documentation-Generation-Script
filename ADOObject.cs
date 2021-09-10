/*
 * Class: ADOObject.cs
 * Purpose: Blueprint for ADO Test Case from Azure DevOps
 * Developer: Zachery Q. Brunner
 */
using System.Collections.Generic;

namespace Script2_ADO_to_Word
{
    /// <summary>
    /// Object blueprint resembling the ADO Test Case from Azure DevOps
    /// </summary>
    public class ADOObject
    {
        /// <summary>
        /// Holds the unique identifier of the ADO test case
        /// </summary>
        public int ADO_UNIQ_ID { get; private set; }

        /// <summary>
        /// Title of the ADO Test Case
        /// </summary>
        public string Title { get; private set; }

        /// <summary>
        /// Hold the name of the tester assigned to the test case
        /// </summary>
        public string Tester { get; private set; }

        /// <summary>
        /// The steps of the ADO Test Case
        /// </summary>
        public List<string> StepAction { get; private set; }

        /// <summary>
        /// The expected result of each step of the ADO Test Case
        /// </summary>
        public List<string> StepExpected { get; private set; }

        /// <summary>
        /// Public constructor for the ADOObject Class
        /// </summary>
        public ADOObject(int ADO_UNIQ_ID, string Title, string Tester, List<string> StepAction, List<string> StepExpected)
        {
            this.ADO_UNIQ_ID = ADO_UNIQ_ID;
            this.Title = Title;
            this.Tester = Tester;
            this.StepAction = StepAction;
            this.StepExpected = StepExpected;
        }
    }
}