using System.Collections.Generic;
using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal sealed class XmlCommentsReader : XmlRecordReader
    {
        private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string ElementComments = "comments";
        private const string ElementCommentList = "commentList";
        private const string ElementComment = "comment";
        private const string ElementText = "text";
        private const string ElementAuthors = "authors";
        private const string ElementAuthor = "author";
        private const string AttributeRef = "ref";


        public XmlCommentsReader(XmlReader reader)
            : base(reader)
        {
        }

        protected override IEnumerable<Record> ReadOverride()
        {if (!Reader.IsStartElement(ElementComments, NsSpreadsheetMl))
            {
                yield break;
            }
            

            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                //var authors = new List<string>();
                //var comments = new Dictionary<string,CommentRecord>();
                //if (Reader.IsStartElement(ElementAuthors, NsSpreadsheetMl))
                //{
                //    var value = StringHelper.ReadStringItem(Reader);
                //    yield return new CommentRecord(value);
                //}
                //else
                if (Reader.IsStartElement(ElementCommentList, NsSpreadsheetMl))
                {
                    foreach (var comment in ReadCellComments())
                    {
                        yield return comment;
                    }
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }
        }

        private IEnumerable<CommentRecord> ReadCellComments()
        {
            if (!XmlReaderHelper.ReadFirstContent(Reader))
            {
                yield break;
            }

            while (!Reader.EOF)
            {
                if (Reader.IsStartElement(ElementComment, NsSpreadsheetMl))
                {
                    var cellRef = Reader.GetAttribute(AttributeRef);
                    var comment = "";

                    if (!XmlReaderHelper.ReadFirstContent(Reader))
                    {
                        yield break;
                    }

                    while (!Reader.EOF)
                    {
                        if (Reader.IsStartElement(ElementText, NsSpreadsheetMl))
                        {
                            comment = StringHelper.ReadStringItem(Reader);
                           // reader.Skip();
                        }
                        else if (!XmlReaderHelper.SkipContent(Reader))
                        {
                            break;
                        }
                    }

                    yield return new CommentRecord(cellRef, comment, "");
                }
                else if (!XmlReaderHelper.SkipContent(Reader))
                {
                    break;
                }
            }
        }
    }
}
