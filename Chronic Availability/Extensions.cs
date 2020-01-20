using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Globalization;
using System.ComponentModel;

namespace Chronic_Availability
{
    static class Extensions
    {


       
        /// Get substring of specified number of characters on the right.
        public static string Right(this string value, int length)
        {
            return value.Substring(value.Length - length);
        }
        /// Get substring of specified number of characters on the Left.
        public static string Left(this string value, int length, int z=0)
        {
            return value.Substring(0,length);
        }
        /// Get substring of specified number of characters on the Mid.
        public static string Mid(this string value, int length)
        {
            return value.Substring(length);
        }
        public static string Mid(this string value, int length, int Index)
        {
            return value.Substring(Index, length);
        }

        /// Get number from string 
        public static string GetNumbers(this string text)
        {
            var result = new List<int>();

            string numberStr = string.Empty;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                // if code of char is between code of '0' and '9'
                if (c >= '0' && c <= '9')
                {
                    numberStr += c;

                    // if is the last char of string, add the last number
                    if (i == text.Length - 1)
                    {
                        result.Add(int.Parse(numberStr));
                    }
                }
                // if char is not a number and numberStr is not empty
                else if (!string.IsNullOrWhiteSpace(numberStr))
                {
                    result.Add(int.Parse(numberStr)); // add the new number

                    numberStr = string.Empty; // clean
                }
            }

            return string.Join("", result.ToArray());
        }
        ///Convert  list of obj  To DataTable
        public static DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties =
               TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

        // This presumes that weeks start with Monday.
        // Week 1 is the 1st week of the year with a Thursday in it.
        public static int GetIso8601WeekOfYear(DateTime time)
        {
            // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll 
            // be the same week# as whatever Thursday, Friday or Saturday are,
            // and we always get those right
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(time);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                time = time.AddDays(3);
            }

            // Return the week of our adjusted day
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(time, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }



    }
}
