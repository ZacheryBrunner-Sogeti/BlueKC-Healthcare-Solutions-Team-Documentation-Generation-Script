/*
 * Class: OutOfBoundsException.cs
 * Purpose: Custom exception for handling the Excel document
 * Developer: Zachery Q. Brunner
 */
using System;

namespace Script2_ADO_to_Word
{
    /// <summary>
    /// Custom exception class - Used when reading Excel documents
    /// </summary>
    public class OutOfBoundsException : Exception
    {
        /// <summary>
        /// Public constructor
        /// </summary>
        /// <param name="message">"First Row Blank - Please check Excel file for errors - Shutting down"</param>
        public OutOfBoundsException(string message)
                                        : base(message)
        { }
    }
}