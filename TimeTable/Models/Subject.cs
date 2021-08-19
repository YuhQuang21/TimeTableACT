using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1
{
    public class Subject
    {
        public WeekDays weekDays { get; set; }
        public string Name { get; set; }
        public string Time { get; set; }
        public string Class { get; set; }
        public string DateAt { get; set; }

        public enum WeekDays : byte
        {
            Monday = 2,
            Tuesday = 3,
            Wednesday = 4,
            Thursday = 5,
            Friday = 6,
            Saturday = 7,
            Sunday = 8
        }
    }
}
