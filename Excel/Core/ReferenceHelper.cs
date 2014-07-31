
using System.Text.RegularExpressions;

namespace Excel.Core
{
    public static class ReferenceHelper
    {
        /// <summary>
        /// Converts references of form A1, B1, C3, DD99 etc to row and col
        /// </summary>
        /// <param name="reference"></param>
        /// <returns>array of two elements 0 index is row num, 1 index is col. Note that the result is 1-based</returns>
        public static int[] ReferenceToColumnAndRow(string reference)
        {
            //split the string into row and column parts
            

            Regex matchLettersNumbers = new Regex("([a-zA-Z]*)([0-9]*)");
            string column = matchLettersNumbers.Match(reference).Groups[1].Value.ToUpper();
            string rowString = matchLettersNumbers.Match(reference).Groups[2].Value;

            //.net 3.5 or 4.5 we could do this awesomeness
            //return reference.Aggregate(0, (s,c)=>{s*26+c-'A'+1});
            //but we are trying to retain 2.0 support so do it a longer way
            //this is basically base 26 arithmetic
            int columnValue = 0;
            int pow = 1;

            //reverse through the string
            for (int i = column.Length - 1; i >= 0; i--)
            {
                int pos = column[i] - 'A' + 1;
                columnValue += pow * pos;
                pow *= 26;
            }

            return new int[2] { int.Parse(rowString), columnValue };
        }
    }
}
