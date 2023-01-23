using System;
using System.Collections.Generic;
using System.Text;
using System.Web;

namespace ExcelDataReader.Core.OpenXmlFormat.Records
{
    internal sealed class CommentRecord : Record
    {
        
        public CommentRecord(string cellRef, string comment, string author) 
        {
            CellRef = cellRef;
            Comment = comment;   
            Author = author;
        }

        public string CellRef { get; }
        public string Comment { get; }
        public string Author { get; }


    }
}
