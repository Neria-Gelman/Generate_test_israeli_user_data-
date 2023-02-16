using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;


namespace MyNamespace
{
	class MyClass
	{
		static void Main(string[] args)
		{
            GeneratePeopleExcelFile();
            Console.WriteLine("success!");
        }

        public static void GeneratePeopleExcelFile()
        {
            // Generate 100 people
            List<Person> people = new List<Person>();
            for (int i = 0; i < 100; i++)
            {
                string idNumber = GenerateIsraeliIdNumber();
                string firstName = GenerateIsraeliFirstName();
                string lastName = GenerateIsraeliLastName();
                string phoneNumber = GenerateIsraeliPhoneNumber();
                people.Add(new Person(idNumber, firstName, lastName, phoneNumber));
            }

            // Export people data to an Excel file
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // Create a new worksheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("People");

                // Write column headers
                worksheet.Cells[1, 1].Value = "ID Number";
                worksheet.Cells[1, 2].Value = "First Name";
                worksheet.Cells[1, 3].Value = "Last Name";
                worksheet.Cells[1, 4].Value = "Phone Number";

                // Write data for each person
                for (int i = 0; i < people.Count; i++)
                {
                    int row = i + 2;
                    worksheet.Cells[row, 1].Value = people[i].IdNumber;
                    worksheet.Cells[row, 2].Value = people[i].FirstName;
                    worksheet.Cells[row, 3].Value = people[i].LastName;
                    worksheet.Cells[row, 4].Value = people[i].phoneNumber;
                }

                // Save the Excel file to disk
                string filePath = "C:\\Users\\neriag\\Desktop\\people.xlsx";
                FileInfo file = new FileInfo(filePath);
                excelPackage.SaveAs(file);
            }
        }

        public static string GenerateIsraeliPhoneNumber()
        {
            Random random = new Random();
            string areaCode = "05" + random.Next(0, 5).ToString();
            string subscriberNumber = random.Next(0, 10000000).ToString("D7");
            return areaCode + "-" + subscriberNumber;
        }

        public static string GenerateIsraeliIdNumber()
        {
            Random random = new Random();

            // Generate gender digit (1 for male, 2 for female)
            int gender = random.Next(1, 3);

            // Generate birth year in the range of 00-99
            int birthYear = random.Next(0, 100);

            // Generate birth month in the range of 01-12
            int birthMonth = random.Next(1, 13);

            // Generate a unique number with 3 digits
            int uniqueNumber = random.Next(1, 1000);

            // Combine all parts to form the ID number
            string idNumber = string.Format("{0}{1:00}{2:00}{3:000}", gender, birthYear, birthMonth, uniqueNumber);

            // Calculate the last digit of the ID number to pass the validation
            int sum = 0;
            for (int i = 0; i < idNumber.Length; i++)
            {
                int digit = int.Parse(idNumber[i].ToString());
                int step = digit * ((i % 2) + 1);
                sum += step > 9 ? step - 9 : step;
            }
            int lastDigit = (10 - (sum % 10)) % 10;

            // Append the last digit to the ID number
            idNumber += lastDigit;

            return idNumber;
        }

        // Function to generate a random first Israeli name
        public static string GenerateIsraeliFirstName()
        {
            Random random = new Random();

            string[] firstNames = { "Yosef", "Moshe", "David", "Avraham", "Meir", "Shlomo", "Yitzhak", "Yehuda", "Haim", "Yaakov", "Eitan", "Oren", "Gideon", "Eli", "Asaf", "Matan", "Itai", "Yair", "Yaniv", "Tal" };

            int randomIndex = random.Next(0, firstNames.Length);

            return firstNames[randomIndex];
        }

        // Function to generate a random last Israeli name
        public static string GenerateIsraeliLastName()
        {
            Random random = new Random();

            string[] lastNames = { "Cohen", "Levi", "Ben David", "Mizrahi", "Avrahami", "Yosef", "Katz", "Peretz", "Zamir", "Ezra", "Shmueli", "Sharon", "Tal", "Baruch", "Goldberg", "Golan", "Mizrachi", "Shaked" };

            int randomIndex = random.Next(0, lastNames.Length);

            return lastNames[randomIndex];
        }
    }

    public class Person
    {
        public string IdNumber { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string phoneNumber { get; set; }

        public Person(string idNumber, string firstName, string lastName, string phoneNumber)
        {
            this.IdNumber = idNumber;
            this.FirstName = firstName;
            this.LastName = lastName;
            this.phoneNumber = phoneNumber;
        }
    }
}