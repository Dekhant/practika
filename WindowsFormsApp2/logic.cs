using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace WindowsFormsApp2
{
    public class logic
    {
        public void CreateDates(string start, string end, ref List<string> dates)
        {
            int i;
            IDictionary<int, int> dict = new Dictionary<int, int>() {
                { 1, 31 },
                { 2, 28 },
                { 3, 31 },
                { 4, 30 },
                { 5, 31 },
                { 6, 30 },
                { 7, 31 },
                { 8, 31 },
                { 9, 30 },
                { 10, 31 },
                { 11, 30 },
                { 12, 31 },
            };
            var dateInMonths = new ReadOnlyDictionary<int, int>(dict);

            int dayStart = int.Parse(start.Split('.')[0]);
            int monthStart = int.Parse(start.Split('.')[1]);

            int dayEnd = int.Parse(end.Split('.')[0]);
            int monthEnd = int.Parse(end.Split('.')[1]);

            while(dayStart != dayEnd || monthStart != monthEnd)
            {
                dateInMonths.TryGetValue(monthStart, out i);
                while(dayStart != i && dayStart != dayEnd)
                {
                    dates.Add(dayStart.ToString() + "." + monthStart.ToString());
                    dayStart++;
                }
                dates.Add(dayStart.ToString() + "." + monthStart.ToString());
                if (monthStart == monthEnd)
                {
                    break;
                }
                monthStart++;
                dayStart = 1;
            }
        }

        public void initializeDates(DateTime start, DateTime end, ref List<string> dates)
        {
            string formatted = start.ToString("dd-MM-yyyy");
            DateTime mayDay;
            while (start.ToString("dd-MM-yyyy") != end.ToString("dd-MM-yyyy"))
            {
                formatted = start.ToString("dd-MM-yyyy");
                dates.Add(formatted);

                mayDay = start.AddDays(1);
                start = mayDay;
            }
            formatted = start.ToString("dd-MM-yyyy");
            dates.Add(formatted);
        }

        public int tableMarkup(string numOfCategories, string[] numOfCategory, string[] roomCapacity)
        {
            int numOfSeats = 0;
            var category = new List<int>();
            for (var i = 0; i < numOfCategory.Length; i++)
            {
                category.Add(int.Parse(numOfCategory[i]));
                numOfSeats += category[i] * int.Parse(roomCapacity[i]);
            }

            return numOfSeats;
        }
    }
}
