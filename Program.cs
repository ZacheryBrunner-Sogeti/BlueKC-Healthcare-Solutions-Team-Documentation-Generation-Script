/*
 * Class: Program.cs
 * Purpose: Manages the program - Main method found here
 * Developer: Zachery Q. Brunner
 */
using System;

namespace Script2_ADO_to_Word
{
    /// <summary>
    /// Program entry point
    /// </summary>
    class Program
    {
        /// <summary>
        /// Main method - Program entry point
        /// </summary>
        /// <param name="args">No arguments inputed</param>
        static void Main(string[] args)
        {
            Console.WriteLine("Script #1 Starting");

            CreateWordDocument CWD = new CreateWordDocument(ReadExcelDocument.GetExcelFile());
            CWD.GenerateDocuments();

            Console.WriteLine("Script #1 Ending");
        }
    }
}