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
        int MAX_VALUE = 13;
        string YES = "Yes";
        string NO = "No";


        string JobCurrentlyValidate = "JobCurrentlyValidate";
        string FirstName = "FirstName";
        string LastName = "LastName";
        string Mentee = "Mentee";
        string Mentor = "Mentor";
        string Type = "Type";
        string YearsExperienceCanada = "YearsExperienceCanada"; 
        string LegallyWorkCanada = "LegallyWorkCanada";
        string AgeGroup = "AgeGroup";
        string MentorshipBefore = "MentorshipBefore"; 
        string JobCurrently = "JobCurrently";
        string Function = "Function";
        string Subfunction = "Subfunction";
        string Industry = "Industry";
        string Experience = "Experience_";
        

        int FirstNameCol = 0;
        int LastNameCol = 0;
        int TypeCol = 0;
        int YearsExperienceCanadaCol = 0;
        int LegallyWorkCanadaCol = 0;
        int AgeGroupCol = 0;
        int MentorshipBeforeCol = 0;
        int JobCurrentlyCol = 0;
        int FunctionCol = 0;
        int IndustryCol = 0;
        Dictionary<string, Person> mentorDic = new Dictionary<string, Person>();

        Dictionary<string, string> matchDic = new Dictionary<string, string>();
        IConfiguration configuration = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .Build();

        public void Match()
        {

            PATH_FROM = configuration.GetValue<string>("MySettings:PathFrom");
            PATH_TO = configuration.GetValue<string>("MySettings:PathTo");

            //MatchAutoContext matchAutoContext = new MatchAutoContext();
            //var menteeList = matchAutoContext.Person.Where(t => t.Type == "Mentee Registration Deposit").OrderByDescending(t => t.IndustryExperience.Length);
            //var mentorList = matchAutoContext.Person.Where(t => t.Type == "Mentor").OrderByDescending(t => t.IndustryExperience.Length);
            var listPerson = ReadFile();
            var menteeList = listPerson.Where(t => t.Type == GetName(Mentee)
            && t.LegallyWorkCanada.Contains(YES) ).ToList();

            var mentorList = listPerson.Where(t => t.Type == GetName(Mentor)).ToList();
            mentorDic = mentorList.ToDictionary(t => t.OrderNo);

            if (GetName(JobCurrentlyValidate) == "true")
            {
                StartMatch(menteeList.Where(t => t.JobCurrently == "No").ToList(), mentorList);
                StartMatch(menteeList.Where(t => t.JobCurrently == "Yes").ToList(), mentorList);
            }
            else
            {
                StartMatch(menteeList, mentorList);
            }

            
            CreateExcel(menteeList, mentorList.OrderBy(t => t.FirsName));
        }


        public void StartMatch(List<Person> menteeList, List<Person> mentorList)
        {
            bool isSearching = true;
            int assignedCountTotal = 0;

            while (isSearching)
            {
                FindBestOption(menteeList, mentorList);
                var assignedCount = menteeList.Where(t => t.OrderNoAssigned != null).Count();
                if (assignedCount == assignedCountTotal)
                {
                    isSearching = false;
                }
                assignedCountTotal = assignedCount;
            }
        }

        public void FindBestOption(List<Person> menteeList, List<Person> mentorList)
        {
            foreach (var mentee in menteeList)
            {
                if (!matchDic.ContainsKey(mentee.OrderNo))
                {
                    int coincidences = 0;

                    string mentorOrderNo = null;
                    foreach (var mentor in mentorList)
                    {
                        if (!matchDic.ContainsValue(mentor.OrderNo))
                        { 
                            int coincidencesNew = 0;
                            coincidencesNew = coincidencesNew + FindCoincidences(mentee.Function, mentor.Function, HIGH_VALUE);
                            coincidencesNew = coincidencesNew + FindCoincidences(mentee.Subfunction, mentor.Subfunction, MEDIUM_VALUE);
                            coincidencesNew = coincidencesNew + FindCoincidences(mentee.Industry, mentor.Industry, HIGH_VALUE);
                            coincidencesNew = coincidencesNew + FindCoincidences(mentee.AgeGroup, mentor.AgeGroup, LOW_VALUE);
                            bool isChanged = false;
                            if (coincidencesNew > coincidences) ///choose the best temporal
                                isChanged = true;

                            if (coincidencesNew == coincidences && coincidencesNew != 0 && coincidences != 0) //if two mentors has the same coincidences
                            {
                                var mentorNew = mentorDic.GetValueOrDefault(mentor.OrderNo);
                                var mentorOld = mentorDic.GetValueOrDefault(mentorOrderNo);
                                if (mentorNew.MentorshipBefore == YES && mentorOld.MentorshipBefore == NO)
                                    isChanged = true;
                                
                                if (mentorNew.MentorshipBefore == YES && mentorOld.MentorshipBefore == YES)///if both have been mentor before
                                    if(FindExperience(mentorNew.YearsExperienceCanada) > FindExperience(mentorOld.YearsExperienceCanada))
                                        isChanged = true;
                            }
                            if(isChanged)
                            {
                                mentorOrderNo = mentor.OrderNo;
                                coincidences = coincidencesNew;
                            }
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

            foreach (var item in menteeList)///assing mentor and mentee final
            {
                if (item.OrderNoAssigned != null && !matchDic.ContainsValue(item.OrderNoAssigned))
                    matchDic.Add(item.OrderNo, item.OrderNoAssigned);
            }

        }


        public int FindExperience(string experience )
        {
            int experienceMentor = 0;
            for (int i = 1; i <= 4; i++)
            {
                var exp = GetName(Experience + i);
                if (exp == experience) 
                    experienceMentor = i;
            }
            return experienceMentor;
        }

        /// <summary>
        /// Compare the previous coincidences saved and Unassign if is necessary
        /// </summary>
        /// <param name="menteeList"></param>
        /// <param name="order"></param>
        /// <param name="coincidences"></param>
        /// <returns></returns>
        public bool UnassignCoincidences(List<Person> menteeList, string order, int coincidences)
        {
            bool isReassign = true;
            foreach (var item in menteeList)
            {
                if (item.OrderNoAssigned == order) ///if mentor have been assinged before 
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
            int coincidences = 0;
            if(menteeList != null && mentorList != null)
            {
                var menteeArray = menteeList.Split('|');
                var mentorArray = mentorList.Split('|');
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
                worksheet.Cells[row, col++].Value = "Mentor Assigned";
                worksheet.Cells[row, col++].Value = "Coincidences";
                worksheet.Cells[row, col++].Value = "Match %";
                worksheet.Cells[row, col++].Value = "Function";
                worksheet.Cells[row, col++].Value = "Subfunction";
                worksheet.Cells[row, col++].Value = "Industry";
                worksheet.Cells[row, col++].Value = "Age Group";
                worksheet.Cells[row, col++].Value = "Job Currently";

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
                    worksheet.Cells[row, col++].Value = (item.Coincidences * 100) / MAX_VALUE;
                    worksheet.Cells[row, col++].Value = item.Function;
                    worksheet.Cells[row, col++].Value = item.Subfunction;
                    worksheet.Cells[row, col++].Value = item.Industry;
                    worksheet.Cells[row, col++].Value = item.AgeGroup;
                    worksheet.Cells[row, col++].Value = item.JobCurrently;
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
                worksheet.Cells[row, col++].Value = "Mentee Assigned";
                worksheet.Cells[row, col++].Value = "Coincidences";
                worksheet.Cells[row, col++].Value = "Match %";
                worksheet.Cells[row, col++].Value = "Function";
                worksheet.Cells[row, col++].Value = "Subfunction";
                worksheet.Cells[row, col++].Value = "Industry";
                worksheet.Cells[row, col++].Value = "Age Group";
                worksheet.Cells[row, col++].Value = "Mentorship Before";
                worksheet.Cells[row, col++].Value = "Years Experience Canada";

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
                    worksheet.Cells[row, col++].Value = (item.Coincidences * 100) / MAX_VALUE;
                    worksheet.Cells[row, col++].Value = item.Function;
                    worksheet.Cells[row, col++].Value = item.Subfunction;
                    worksheet.Cells[row, col++].Value = item.Industry;
                    worksheet.Cells[row, col++].Value = item.AgeGroup;
                    worksheet.Cells[row, col++].Value = item.MentorshipBefore;
                    worksheet.Cells[row, col++].Value = item.YearsExperienceCanada;
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
            Dictionary<int, int> subfunctionDic = new Dictionary<int, int>();

            int row = 0;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(@PATH_FROM, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read()) //Each ROW
                    {
                        if (row == 0)//read column names
                        {
                            for (int column = 0; column < reader.FieldCount; column++)
                            {
                                if (getValue(reader, column) == GetName(FirstName)) FirstNameCol = column;
                                if (getValue(reader, column) == GetName(LastName)) LastNameCol = column;
                                if (getValue(reader, column) == GetName(Type)) TypeCol = column;
                                if (getValue(reader, column) == GetName(YearsExperienceCanada)) YearsExperienceCanadaCol = column;
                                if (getValue(reader, column) == GetName(LegallyWorkCanada)) LegallyWorkCanadaCol = column;
                                if (getValue(reader, column) == GetName(AgeGroup)) AgeGroupCol = column;
                                if (getValue(reader, column) == GetName(MentorshipBefore)) MentorshipBeforeCol = column;
                                if (getValue(reader, column) == GetName(JobCurrently)) JobCurrentlyCol = column;
                                if (getValue(reader, column) == GetName(Function)) FunctionCol = column;
                                if (getValue(reader, column).Contains(GetName(Subfunction))) subfunctionDic.Add(column, column);
                                if (getValue(reader, column) == GetName(Industry)) IndustryCol = column;

                            }
                        }
                        else
                        {


                            Person person = new Person();
                            for (int column = 0; column < reader.FieldCount; column++)
                            {
                                //CreateLog(column + "-"+ getValue(reader, column));
                                if (column == 0) person.OrderNo = row.ToString();
                                if (column == FirstNameCol) person.FirsName = getValue(reader, column);
                                if (column == LastNameCol) person.LastName = getValue(reader, column);
                                if (column == TypeCol) person.Type = getValue(reader, column);
                                if (column == YearsExperienceCanadaCol) person.YearsExperienceCanada = getValue(reader, column);
                                if (column == LegallyWorkCanadaCol) person.LegallyWorkCanada = getValue(reader, column);
                                if (column == AgeGroupCol) person.AgeGroup = getValue(reader, column);
                                if (column == MentorshipBeforeCol) person.MentorshipBefore = getValue(reader, column);
                                if (column == JobCurrentlyCol) person.JobCurrently = getValue(reader, column);
                                if (column == FunctionCol) person.Function = getValue(reader, column);
                                if (subfunctionDic.ContainsKey(column)) person.Subfunction = SubfunctionAdd(person.Subfunction, getValue(reader, column));
                                if (column == IndustryCol) person.Industry = getValue(reader, column);
                            }
                            personList.Add(person);

                        }

                        row++;
                    }
                }
            }
            return personList;
        }
        public string SubfunctionAdd(string subfunctionSave, string subfunctionNew)
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

        public string GetName(string value)
        {
            return configuration.GetValue<string>("MySettings:" + value);
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
