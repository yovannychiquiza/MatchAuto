using ExcelDataReader;
using MatchAuto.Model;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;


namespace MatchAuto
{
    public class Process
    {
        static string PATH_FROM = "";
        static string PATH_TO = "";
        int HIGH_VALUE = 3; 
        int MEDIUM_VALUE = 2; 
        int LOW_VALUE = 1;
        
        Dictionary<string, Person> dicMatch = new Dictionary<string, Person>();

        public void Match()
        {
            IConfiguration configuration = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .Build();

            PATH_FROM = configuration.GetValue<string>("MySettings:PathFrom");
            PATH_TO = configuration.GetValue<string>("MySettings:PathTo");

            //MatchAutoContext matchAutoContext = new MatchAutoContext();
            //var menteeList = matchAutoContext.Person.Where(t => t.Type == "Mentee Registration Deposit").OrderByDescending(t => t.IndustryExperience.Length);
            //var mentorList = matchAutoContext.Person.Where(t => t.Type == "Mentor").OrderByDescending(t => t.IndustryExperience.Length);
            var listPerson = ReadFile();
            var menteeList = listPerson.Where(t => t.Type == "Mentee Registration Deposit")
                //.OrderByDescending(t => t.LegallyWorkCanada)
                //.OrderByDescending(t => t.Function.Length)
                //.OrderByDescending(t => t.Subfunction.Length)
                //.OrderByDescending(t => t.Industry.Length)
                .ToList();
            var mentorList = listPerson.Where(t => t.Type == "Mentor")
                //.OrderByDescending(t => t.Industry.Length)
                .ToList();

            FindBestOption(menteeList, mentorList);
            //third loop
            //foreach (var mentee in menteeList)
            //{
            //    foreach (var mentor in mentorList)
            //    {
            //        if (mentee.Assigned == null && mentor.Assigned == null)//if the mentee is not assigned
            //        {
            //            mentee.Assigned = mentor.FirsName + " " + mentor.LastName;
            //            mentor.Assigned = mentee.FirsName + " " + mentee.LastName;
            //            break;
            //        }
            //    }
            //}

            CreateExcel(menteeList, mentorList.OrderBy(t => t.FirsName));
        }

        public void FindBestOption(List<Person> menteeList, List<Person> mentorList)
        {
            ///first loop criteria (MentorshipBefore = No, IndustryExperience)
            foreach (var mentee in menteeList)
            {
                int coincidences = 0;

                string mentorOrderNo = null;
                foreach (var mentor in mentorList)
                {
                    int coincidencesNew = 0;
                    coincidencesNew = coincidencesNew + FindCoincidences(mentee.Function, mentor.Function, HIGH_VALUE);
                    coincidencesNew = coincidencesNew + FindCoincidences(mentee.Subfunction, mentor.Subfunction, MEDIUM_VALUE);
                    coincidencesNew = coincidencesNew + FindCoincidences(mentee.Industry, mentor.Industry, HIGH_VALUE);
                    //coincidencesNew = coincidencesNew + FindCoincidences(mentee.AgeGroup, mentor.AgeGroup, LOW_VALUE);

                    if (coincidencesNew > coincidences) ///choose the best
                    {
                        mentorOrderNo = mentor.OrderNo;
                        coincidences = coincidencesNew;
                    }
                }

                foreach (var mentor in mentorList)///assing mentor and mentee
                {
                    if (mentor.OrderNo == mentorOrderNo)
                    {
                        if(UnassignCoincidences(menteeList, mentorOrderNo, coincidences))
                        {
                            mentee.Coincidences = coincidences;
                            mentee.OrderNoAssigned = mentorOrderNo;
                            break;
                        }
                    }

                }
            }

        }

        public bool UnassignCoincidences(List<Person> menteeList, string order, int coincidences)
        {
            bool isReassign = true;
            foreach (var item in menteeList)
            {
                if (item.OrderNoAssigned == order)
                {
                    if (coincidences > item.Coincidences)
                    {
                        item.OrderNoAssigned = null;
                        item.Coincidences = 0;
                        isReassign = true;
                    }
                    else
                    {
                        isReassign = false;
                    }
                    break;
                }
            }
            return isReassign;
        }

        //find coincidences
        public int FindCoincidences(string menteeList, string mentorList, int increment)
        {
            var menteeArray = menteeList.Split('|');
            var mentorArray = mentorList.Split('|');
            int coincidences = 0;
            foreach (var item1 in menteeArray)
            {
                foreach (var item2 in mentorArray)
                {
                    if (item1.Trim() == item2.Trim())
                    {
                        coincidences += increment;
                    }
                }
            }
            return coincidences;
        }

        public void CreateExcel(IEnumerable<Person> menteeList, IEnumerable<Person> mentorList) {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");

                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                int row = 1;
                int col = 1;
                worksheet.Cells[row, 2].Value = "Mentee";
                worksheet.Cells[row, 2].Style.Font.Bold = true;
                row++;
                worksheet.Cells[row, col++].Value = "No";
                worksheet.Cells[row, col++].Value = "Firs Name";
                worksheet.Cells[row, col++].Value = "Last Name";
                worksheet.Cells[row, col++].Value = "Assigned";
                worksheet.Cells[row, col++].Value = "Coincidences";
                worksheet.Cells[row, col++].Value = "Function";
                worksheet.Cells[row, col++].Value = "Subfunction";
                worksheet.Cells[row, col++].Value = "Industry Experience";

                worksheet.Cells[ExcelRange.GetAddress(row, 1, row, col)].Style.Font.Bold = true;
                int cont = 1;
                var dicMentorList = mentorList.ToDictionary(t => t.OrderNo);

                foreach (var item in menteeList)
                {
                    row++;
                    col = 1;
                    string assingTo = "";

                    if (item.OrderNoAssigned != null && dicMentorList.ContainsKey(item.OrderNoAssigned))
                    {
                        var person = dicMentorList.GetValueOrDefault(item.OrderNoAssigned);
                        assingTo = person.FirsName + " " + person.LastName;
                    }

                    worksheet.Cells[row, col++].Value = item.OrderNo;
                    worksheet.Cells[row, col++].Value = item.FirsName;
                    worksheet.Cells[row, col++].Value = item.LastName;
                    worksheet.Cells[row, col++].Value = assingTo;
                    worksheet.Cells[row, col++].Value = item.Coincidences;
                    worksheet.Cells[row, col++].Value = item.Function;
                    worksheet.Cells[row, col++].Value = item.Subfunction;
                    worksheet.Cells[row, col++].Value = item.Industry;
                }

                ///////////////////////////////////////////////////////////////////////////
                row = row + 2;

                worksheet.Cells[row, 2].Value = "Mentor";
                worksheet.Cells[row, 2].Style.Font.Bold = true;
                row++;
                col = 1;
                worksheet.Cells[row, col++].Value = "No";
                worksheet.Cells[row, col++].Value = "FirsName";
                worksheet.Cells[row, col++].Value = "LastName";
                worksheet.Cells[row, col++].Value = "Assigned";
                worksheet.Cells[row, col++].Value = "Coincidences";
                worksheet.Cells[row, col++].Value = "Function";
                worksheet.Cells[row, col++].Value = "Subfunction";
                worksheet.Cells[row, col++].Value = "Industry Experience";

                worksheet.Cells[ExcelRange.GetAddress(row, 1, row, col)].Style.Font.Bold = true;
                cont = 1;
                var dicMenteeList = menteeList.Where(t => t.OrderNoAssigned != null).ToDictionary(t => t.OrderNoAssigned);

                foreach (var item in mentorList)
                {
                    row++;
                    col = 1;
                    string assingTo = "";
                    if (item.OrderNo != null && dicMenteeList.ContainsKey(item.OrderNo))
                    {
                        var person = dicMenteeList.GetValueOrDefault(item.OrderNo);
                        assingTo = person.FirsName + " " + person.LastName;
                    }

                    worksheet.Cells[row, col++].Value = item.OrderNo;
                    worksheet.Cells[row, col++].Value = item.FirsName;
                    worksheet.Cells[row, col++].Value = item.LastName;
                    worksheet.Cells[row, col++].Value = assingTo;
                    worksheet.Cells[row, col++].Value = item.Coincidences;
                    worksheet.Cells[row, col++].Value = item.Function;
                    worksheet.Cells[row, col++].Value = item.Subfunction;
                    worksheet.Cells[row, col++].Value = item.Industry;
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                FileInfo excelFile = new FileInfo(@PATH_TO);
                excel.SaveAs(excelFile);
                Console.WriteLine("Match Process Finished");

            }
        }


        public List<Person> ReadFile()
        {

            List<Person> personList = new List<Person>();

            int row = 0;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(@PATH_FROM, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read()) //Each ROW
                    {
                        if (row > 0)
                        {
                            Person person = new Person();
                            for (int column = 0; column < reader.FieldCount; column++)
                            {
                                 
                                //CreateLog(column + "-"+ getValue(reader, column));
                                if (column == 0) person.OrderNo = row.ToString();
                                if (column == 2) person.FirsName = getValue(reader, column);
                                if (column == 3) person.LastName = getValue(reader, column);
                                if (column == 7) person.Type = getValue(reader, column);
                                //if (column == 11) person.AttendeeStatus = getValue(reader, column);
                                if (column == 26) person.YearsExperienceCanada = getValue(reader, column);
                                //if (column == 16) person.ApplyingTo = getValue(reader, column);
                                if (column == 27) person.OrganizationMember = getValue(reader, column);
                                if (column == 30) person.LegallyWorkCanada = getValue(reader, column);
                                if (column == 33) person.AgeGroup = getValue(reader, column);
                                if (column == 39) person.MentorshipBefore = getValue(reader, column);
                                if (column == 49) person.JobCurrently = getValue(reader, column);
                                if (column == 50) person.Function = getValue(reader, column);
                                if (column >=51 && column <= 68 ) person.Subfunction = Subfunction(person.Subfunction, getValue(reader, column));
                                if (column == 69) person.Industry = getValue(reader, column);
                            }
                            personList.Add(person);

                        }

                        row++;
                    }
                }
            }
            return personList;
        }
        public string Subfunction(string subfunctionSave, string subfunctionNew)
        {
            if(subfunctionNew != null)
            {
                if(subfunctionSave == null)
                    subfunctionSave = subfunctionNew;
                else
                    subfunctionSave = subfunctionSave + "|" + subfunctionNew;
            }

            return subfunctionSave;
        }

        public static string getValue(IExcelDataReader reader, int column)
        {
            string val = null;
            if (reader.GetValue(column) != null)
            {
                val = reader.GetValue(column).ToString();
            }
            return val;
        }
        public static void CreateLog(string text)
        {
            string path = @"C:\IIS\";
            string filename = "Log_" + DateTime.Now.ToString("yyyy_MM") + ".txt";

            if (!Directory.Exists(path))
            {
                DirectoryInfo di = Directory.CreateDirectory(path);
            }
            using (StreamWriter writer = new StreamWriter(path + filename, true))
            {
                writer.WriteLine(text);
                writer.Close();
            }
        }



        //public void ReadFileSave()
        //{

        //    MatchAutoContext matchAutoContext = new MatchAutoContext();
        //    List<Person> personList = new List<Person>();

        //    int row = 0;
        //    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        //    using (var stream = File.Open(@PATH_FROM, FileMode.Open, FileAccess.Read))
        //    {
        //        using (var reader = ExcelReaderFactory.CreateReader(stream))
        //        {
        //            while (reader.Read()) //Each ROW
        //            {
        //                if(row > 0){
        //                    Person person = new Person();
        //                    for (int column = 0; column < reader.FieldCount; column++)
        //                    {
        //                        //CreateLog(column + "-"+ reader.GetValue(column).ToString());
        //                        if (column == 2) person.FirsName = getValue(reader, column);
        //                        if (column == 3) person.LastName = getValue(reader, column);
        //                        if (column == 6) person.Type = getValue(reader, column);
        //                        if (column == 11) person.AttendeeStatus = getValue(reader, column);
        //                        if (column == 14) person.YearsExperienceCanada = getValue(reader, column);
        //                        if (column == 16) person.ApplyingTo = getValue(reader, column);
        //                        if (column == 17) person.OrganizationMember = getValue(reader, column);
        //                        if (column == 18) person.LegallyWorkCanada = getValue(reader, column);
        //                        if (column == 20) person.AgeGroup = getValue(reader, column);
        //                        if (column == 29) person.MentorshipBefore = getValue(reader, column);
        //                        if (column == 31) person.IndustryExperience = getValue(reader, column);
        //                        if (column == 34) person.ProfessionalInterest = getValue(reader, column);
        //                    }
        //                    personList.Add(person);
        //                    matchAutoContext.Add(person);

        //                }

        //                row++;
        //            }
        //        }
        //    }

        //    matchAutoContext.SaveChanges();
        //}
    }
}
