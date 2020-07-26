using ExcelDataReader;
using MatchAuto.Model;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;


namespace MatchAuto
{
    public class Process
    {
        static string PATH_FROM = "";
        static string PATH_TO = "";
        static string PATH_FROM_MEMBER_TYPE = "";
        static string PATH_FROM_MENTOR_MENTEE = "";
        int HIGH_VALUE = 3;
        int MEDIUM_VALUE = 2;
        int LOW_VALUE = 1;
        int MAX_VALUE = 13;
        string YES = "Yes";
        string NO = "No";

        string FirstName = "FirstName";
        string LastName = "LastName";
        string Email = "Email";
        string MemberType = "MemberType";
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
        string IndustryExperience = "IndustryExperience";
        string Experience = "Experience_";
        string AgeGroup_ = "AgeGroup_";
        string LinkedIn = "LinkedIn";


        int FirstNameCol = 0;
        int LastNameCol = 0;
        int EmailCol = 0;
        int TypeCol = 0;
        int MemberTypeCol = 0;
        int YearsExperienceCanadaCol = 0;
        int LegallyWorkCanadaCol = 0;
        int AgeGroupCol = 0;
        int MentorshipBeforeCol = 0;
        int JobCurrentlyCol = 0;
        int FunctionCol = 0;
        int IndustryCol = 0;
        int IndustryExperienceCol = 0;
        int LinkedInCol = 0;

        Dictionary<string, Person> mentorDic = new Dictionary<string, Person>();

        Dictionary<string, string> matchDic = new Dictionary<string, string>();
        IConfiguration configuration = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .Build();

        public void Match()
        {

            PATH_FROM_MEMBER_TYPE = configuration.GetValue<string>("MySettings:PathFromMemberType");
            PATH_FROM_MENTOR_MENTEE = configuration.GetValue<string>("MySettings:PathFromMentorMentee");
            PATH_FROM = configuration.GetValue<string>("MySettings:PathFrom");
            PATH_TO = configuration.GetValue<string>("MySettings:PathTo");

            var listPerson = ReadFile();
            var menteeList = listPerson.Where(t => t.Type == GetName(Mentee)
            && t.LegallyWorkCanada.Contains(YES)).ToList();

            var mentorList = listPerson.Where(t => t.Type == GetName(Mentor)).ToList();
            mentorDic = mentorList.ToDictionary(t => t.OrderNo);

            StartMatch(menteeList.Where(t => t.JobCurrently == "No").ToList(), mentorList);
            StartMatch(menteeList.Where(t => t.JobCurrently == "Yes").ToList(), mentorList);
         
            CreateExcel(menteeList.OrderByDescending(t => t.Coincidences), mentorList.OrderBy(t => t.FirsName), listPerson);
        }

        /// <summary>
        /// Start the loop until all match possible is done 
        /// </summary>
        /// <param name="menteeList"></param>
        /// <param name="mentorList"></param>
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

        /// <summary>
        /// Find the best option for mentor mentee
        /// </summary>
        /// <param name="menteeList"></param>
        /// <param name="mentorList"></param>
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
                            if (FindAgeGroup(mentor.AgeGroup) >= FindAgeGroup(mentee.AgeGroup)) //Compare if mentor age group is greater or equal than mentee age group
                            { 
                                coincidencesNew = coincidencesNew + FindCoincidences(mentee.Function, mentor.Function, HIGH_VALUE);
                                coincidencesNew = coincidencesNew + FindCoincidences(mentee.Subfunction, mentor.Subfunction, LOW_VALUE);
                                coincidencesNew = coincidencesNew + FindCoincidences(mentee.Industry, mentor.Industry, HIGH_VALUE);
                                coincidencesNew = coincidencesNew + FindCoincidences(mentee.MemberType, mentor.MemberType, HIGH_VALUE);
                                coincidencesNew = coincidencesNew + LOW_VALUE;
                            }
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
                                    if (FindExperience(mentorNew.YearsExperienceCanada) > FindExperience(mentorOld.YearsExperienceCanada)) //Compare the best choice for mentor experience
                                         isChanged = true;
                            }
                            if (isChanged)
                            {
                                mentorOrderNo = mentor.OrderNo;
                                coincidences = coincidencesNew;
                            }
                        }
                    }

                    foreach (var mentor in mentorList)///assing mentor and mentee temporal
                    {
                        if (mentor.OrderNo == mentorOrderNo)
                        {
                            if (UnassignCoincidences(menteeList, mentorOrderNo, coincidences))
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


        public int FindAgeGroup(string age)
        {
            int value = 0, i = 1;
            while (GetName(AgeGroup_ + i) != null)
            {
                if (GetName(AgeGroup_ + i) == age)
                    value = i;
                i++;
            }
            return value;
        }

        /// <summary>
        /// Return the number for mentor experience 
        /// </summary>
        /// <param name="experience"></param>
        /// <returns></returns>
        public int FindExperience(string experience)
        {
            int experienceMentor = 0, i = 1;
            experience = experience.Replace(' ', ' ');//it is required if contains special character 
            while (GetName(Experience + i) != null)
            {
                if (GetName(Experience + i) == experience)
                    experienceMentor = i;
                i++;
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

        /// <summary>
        /// Find coincidences and increase score
        /// </summary>
        /// <param name="menteeList"></param>
        /// <param name="mentorList"></param>
        /// <param name="increment"></param>
        /// <returns></returns>
        public int FindCoincidences(string menteeList, string mentorList, int increment)
        {
            int coincidences = 0;
            if (menteeList != null && mentorList != null)
            {
                var menteeArray = menteeList.Split('|');
                var mentorArray = mentorList.Split('|');
                foreach (var item1 in menteeArray)
                {
                    foreach (var item2 in mentorArray)
                    {
                        if (item1.Trim() == item2.Trim()) //if match
                        {
                            coincidences += increment;// then count 
                        }
                    }
                }
            }
            return coincidences;
        }

        /// <summary>
        /// Create excel file with match generated
        /// </summary>
        /// <param name="menteeList"></param>
        /// <param name="mentorList"></param>
        public void CreateExcel(IEnumerable<Person> menteeList, IEnumerable<Person> mentorList, IEnumerable<Person> listPerson)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Match");

                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Match"];

                int row = 1;
                int col = 1;
    
                worksheet.Cells[row, col++].Value = "Code";
                worksheet.Cells[row, col++].Value = "Type";
                worksheet.Cells[row, col++].Value = "Firs Name";
                worksheet.Cells[row, col++].Value = "Last Name";
                worksheet.Cells[row, col++].Value = "Email";
                worksheet.Cells[row, col++].Value = "Score";
                worksheet.Cells[row, col++].Value = "Match %";
                worksheet.Cells[row, col++].Value = "Function";
                worksheet.Cells[row, col++].Value = "Subfunction";
                worksheet.Cells[row, col++].Value = "Industry";
                worksheet.Cells[row, col++].Value = "Age Group";
                worksheet.Cells[row, col++].Value = "Job Currently";
                worksheet.Cells[row, col++].Value = "Member Type";
                worksheet.Cells[row, col++].Value = "Mentorship Before";
                worksheet.Cells[row, col++].Value = "Years Experience Canada";
                worksheet.Cells[row, col++].Value = "LinkedIn";
                worksheet.Cells[ExcelRange.GetAddress(row, 1, row, col)].Style.Font.Bold = true;
                
                var dicMentorList = mentorList.ToDictionary(t => t.OrderNo);
                int cod = 1;
                bool colorChange = false;
                foreach (var item in menteeList.Where(t => t.OrderNoAssigned != null))
                {
                    row++;
                    col = 1;
         
                    Person person = null;
                    if (item.OrderNoAssigned != null && dicMentorList.ContainsKey(item.OrderNoAssigned))
                    {
                        person = dicMentorList.GetValueOrDefault(item.OrderNoAssigned);
                    }

                    worksheet.Cells[row, col++].Value = person != null ? cod: 0;
                    worksheet.Cells[row, col++].Value = "Mentee";
                    worksheet.Cells[row, col++].Value = item.FirsName;
                    worksheet.Cells[row, col++].Value = item.LastName;
                    worksheet.Cells[row, col++].Value = item.Email;
                    worksheet.Cells[row, col++].Value = item.Coincidences;
                    worksheet.Cells[row, col++].Value = (item.Coincidences * 100) / MAX_VALUE;
                    worksheet.Cells[row, col++].Value = item.Function;
                    worksheet.Cells[row, col++].Value = item.Subfunction;
                    worksheet.Cells[row, col++].Value = item.Industry;
                    worksheet.Cells[row, col++].Value = item.AgeGroup;
                    worksheet.Cells[row, col++].Value = item.JobCurrently;
                    worksheet.Cells[row, col++].Value = item.MemberType;
                    col += 2;
                    worksheet.Cells[row, col++].Value = item.LinkedIn;

                    if (colorChange)
                    {
                        worksheet.Cells[ExcelRange.GetAddress(row, 1, row, 16)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[ExcelRange.GetAddress(row, 1, row, 16)].Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
                    }
                    else
                    {
                        worksheet.Cells[ExcelRange.GetAddress(row, 1, row, 16)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[ExcelRange.GetAddress(row, 1, row, 16)].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                    }

                    if (person != null)
                    {
                        row++;
                        col = 1;
                        worksheet.Cells[row, col++].Value = cod;
                        worksheet.Cells[row, col++].Value = "Mentor";
                        worksheet.Cells[row, col++].Value = person.FirsName;
                        worksheet.Cells[row, col++].Value = person.LastName;
                        worksheet.Cells[row, col++].Value = person.Email;
                        worksheet.Cells[row, col++].Value = "";//person.Coincidences;
                        worksheet.Cells[row, col++].Value = "";//(item.Coincidences * 100) / MAX_VALUE;
                        worksheet.Cells[row, col++].Value = person.Function;
                        worksheet.Cells[row, col++].Value = person.Subfunction;
                        worksheet.Cells[row, col++].Value = person.Industry;
                        worksheet.Cells[row, col++].Value = person.AgeGroup;
                        col++;
                        worksheet.Cells[row, col++].Value = person.MemberType;
                        worksheet.Cells[row, col++].Value = person.MentorshipBefore;
                        worksheet.Cells[row, col++].Value = person.YearsExperienceCanada;
                        worksheet.Cells[row, col++].Value = person.LinkedIn;
                        cod++;
                        if (colorChange)
                        {
                            worksheet.Cells[ExcelRange.GetAddress(row, 1, row, 16)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[ExcelRange.GetAddress(row, 1, row, 16)].Style.Fill.BackgroundColor.SetColor(Color.LightSkyBlue);
                        }
                        else
                        {
                            worksheet.Cells[ExcelRange.GetAddress(row, 1, row, 16)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[ExcelRange.GetAddress(row, 1, row, 16)].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                        }

                    }
                    colorChange = colorChange ? false : true;
                }
                ////////////////////////////////////////////////////////////////////////////


                excel.Workbook.Worksheets.Add("Unmatch");

                // Target a worksheet
                var unmatch = excel.Workbook.Worksheets["Unmatch"];

                row = 1;
                col = 1;

                unmatch.Cells[row, col++].Value = "Code";
                unmatch.Cells[row, col++].Value = "Type";
                unmatch.Cells[row, col++].Value = "Firs Name";
                unmatch.Cells[row, col++].Value = "Last Name";
                unmatch.Cells[row, col++].Value = "Email";
                unmatch.Cells[row, col++].Value = "Function";
                unmatch.Cells[row, col++].Value = "Subfunction";
                unmatch.Cells[row, col++].Value = "Industry";
                unmatch.Cells[row, col++].Value = "Age Group";
                unmatch.Cells[row, col++].Value = "Member Type";
                unmatch.Cells[row, col++].Value = "LinkedIn";

                unmatch.Cells[ExcelRange.GetAddress(row, 1, row, col)].Style.Font.Bold = true;

                foreach (var item in menteeList.Where(t => t.OrderNoAssigned == null).OrderBy(t => t.Function))
                {

                    col = 1;
                    row++;
                    unmatch.Cells[row, col++].Value = row - 1;
                    unmatch.Cells[row, col++].Value = "Mentee";
                    unmatch.Cells[row, col++].Value = item.FirsName;
                    unmatch.Cells[row, col++].Value = item.LastName;
                    unmatch.Cells[row, col++].Value = item.Email;
                    unmatch.Cells[row, col++].Value = item.Function;
                    unmatch.Cells[row, col++].Value = item.Subfunction;
                    unmatch.Cells[row, col++].Value = item.Industry;
                    unmatch.Cells[row, col++].Value = item.AgeGroup;
                    unmatch.Cells[row, col++].Value = item.MemberType;
                    unmatch.Cells[row, col++].Value = item.LinkedIn;

                }


                var dicMenteeList = menteeList.Where(t => t.OrderNoAssigned != null).ToDictionary(t => t.OrderNoAssigned);

                foreach (var item in mentorList)
                {
                    if (!dicMenteeList.ContainsKey(item.OrderNo))
                    {
                        col = 1;
                        row++;
                        unmatch.Cells[row, col++].Value = row - 1;
                        unmatch.Cells[row, col++].Value = item.Type;
                        unmatch.Cells[row, col++].Value = item.FirsName;
                        unmatch.Cells[row, col++].Value = item.LastName;
                        unmatch.Cells[row, col++].Value = item.Email;
                        unmatch.Cells[row, col++].Value = item.Function;
                        unmatch.Cells[row, col++].Value = item.Subfunction;
                        unmatch.Cells[row, col++].Value = item.Industry;
                        unmatch.Cells[row, col++].Value = item.AgeGroup;
                        unmatch.Cells[row, col++].Value = item.MemberType;
                        unmatch.Cells[row, col++].Value = item.LinkedIn;
                    }
                   
                }

                ////////////////////////////////////////////////////////////////////////////

                excel.Workbook.Worksheets.Add("Excluded");

                // Target a worksheet
                var excluded = excel.Workbook.Worksheets["Excluded"];

                row = 1;
                col = 1;

                excluded.Cells[row, col++].Value = "Code";
                excluded.Cells[row, col++].Value = "Type";
                excluded.Cells[row, col++].Value = "Firs Name";
                excluded.Cells[row, col++].Value = "Last Name";
                excluded.Cells[row, col++].Value = "Email";
                excluded.Cells[row, col++].Value = "Function";
                excluded.Cells[row, col++].Value = "Subfunction";
                excluded.Cells[row, col++].Value = "Industry";
                excluded.Cells[row, col++].Value = "Age Group";
                excluded.Cells[row, col++].Value = "Work permit";

                excluded.Cells[ExcelRange.GetAddress(row, 1, row, col)].Style.Font.Bold = true;

                foreach (var item in listPerson.Where(t => t.Type.Equals("Super Mentor") ||
                                                           t.Type.Equals("Volunteer") 
                                                           || (t.LegallyWorkCanada != null && t.LegallyWorkCanada.Equals("No"))
                                                           ).OrderBy(t => t.Type))
                                                           
                {
                    
                    col = 1;
                    row++;
                    excluded.Cells[row, col++].Value = row - 1;
                    excluded.Cells[row, col++].Value = item.Type;
                    excluded.Cells[row, col++].Value = item.FirsName;
                    excluded.Cells[row, col++].Value = item.LastName;
                    excluded.Cells[row, col++].Value = item.Email;
                    excluded.Cells[row, col++].Value = item.Function;
                    excluded.Cells[row, col++].Value = item.Subfunction;
                    excluded.Cells[row, col++].Value = item.Industry;
                    excluded.Cells[row, col++].Value = item.AgeGroup;
                    excluded.Cells[row, col++].Value = item.LegallyWorkCanada;

                }
                
                ////////////////////////////////////////////////////////////////////////////

                //worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                FileInfo excelFile = new FileInfo(@PATH_TO);
                excel.SaveAs(excelFile);
                Console.WriteLine("Match Process Finished");

            }
        }

        /// <summary>
        /// Read file generated from eventbrite
        /// </summary>
        /// <returns></returns>
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
                                if (getValue(reader, column) == GetName(Email)) EmailCol = column;
                                if (getValue(reader, column) == GetName(Type)) TypeCol = column;
                                if (getValue(reader, column) == GetName(YearsExperienceCanada)) YearsExperienceCanadaCol = column;
                                if (getValue(reader, column) == GetName(LegallyWorkCanada)) LegallyWorkCanadaCol = column;
                                if (getValue(reader, column) == GetName(AgeGroup)) AgeGroupCol = column;
                                if (getValue(reader, column) == GetName(MentorshipBefore)) MentorshipBeforeCol = column;
                                if (getValue(reader, column) == GetName(JobCurrently)) JobCurrentlyCol = column;
                                if (getValue(reader, column) == GetName(Function)) FunctionCol = column;
                                if (getValue(reader, column).Contains(GetName(Subfunction))) subfunctionDic.Add(column, column);
                                if (getValue(reader, column) == GetName(Industry)) IndustryCol = column;
                                if (getValue(reader, column) == GetName(IndustryExperience)) IndustryExperienceCol = column;
                                if (getValue(reader, column) == GetName(LinkedIn)) LinkedInCol = column;

                            }
                        }
                        else
                        {


                            Person person = new Person();
                            for (int column = 0; column < reader.FieldCount; column++)
                            {
                                if (column == 0) person.OrderNo = row.ToString();
                                if (column == FirstNameCol) person.FirsName = getValue(reader, column);
                                if (column == LastNameCol) person.LastName = getValue(reader, column);
                                if (column == EmailCol) person.Email = getValue(reader, column);
                                if (column == TypeCol) person.Type = getValue(reader, column);
                                if (column == YearsExperienceCanadaCol) person.YearsExperienceCanada = getValue(reader, column);
                                if (column == LegallyWorkCanadaCol) person.LegallyWorkCanada = getValue(reader, column);
                                if (column == AgeGroupCol) person.AgeGroup = getValue(reader, column);
                                if (column == MentorshipBeforeCol) person.MentorshipBefore = getValue(reader, column);
                                if (column == JobCurrentlyCol) person.JobCurrently = getValue(reader, column);
                                if (column == FunctionCol) person.Function = getValue(reader, column);
                                if (subfunctionDic.ContainsKey(column)) person.Subfunction = SubfunctionAdd(person.Subfunction, getValue(reader, column));
                                if (column == IndustryCol) person.Industry = getValue(reader, column) != null ? getValue(reader, column) : person.Industry;
                                if (column == IndustryExperienceCol) person.Industry = getValue(reader, column) != null ? getValue(reader, column): person.Industry;
                                if (column == LinkedInCol) person.LinkedIn = getValue(reader, column);
                            }
                            personList.Add(person);

                        }

                        row++;
                    }
                }
            }

            personList = ConfigureMemberType(personList, ReadFileMemberType());
            personList = ConfigureExistingMentorMentee(personList, ReadFileMentorMentee());
            return personList;
        }

        /// <summary>
        /// Assign member type to person list
        /// </summary>
        /// <param name="personList"></param>
        /// <param name="personMemberType"></param>
        /// <returns></returns>
        public List<Person> ConfigureMemberType(List<Person> personList, List<Person> personMemberType)
        {
            var dicMemberTypeList = personMemberType.ToDictionary(t => t.Email);
            foreach (var item in personList)
            {
                if (dicMemberTypeList.ContainsKey(item.Email))
                {
                    var person = dicMemberTypeList.GetValueOrDefault(item.Email);
                    item.MemberType = person.MemberType;
                }
            }
            return personList;
        }

        /// <summary>
        /// Assign mentor mentee pair to person list  
        /// </summary>
        /// <param name="personList"></param>
        /// <param name="mentorMenteeDic"></param>
        /// <returns></returns>
        public List<Person> ConfigureExistingMentorMentee(List<Person> personList, Dictionary<string, string> mentorMenteeDic)
        {
            var menteeDic = personList.ToDictionary(t => t.Email);

            foreach (var item in personList)
            {
                if (mentorMenteeDic.ContainsKey(item.Email))
                {
                    var person = menteeDic.GetValueOrDefault(mentorMenteeDic.GetValueOrDefault(item.Email));//get mentor id
                    item.OrderNoAssigned = person.OrderNo;
                    matchDic.Add(item.OrderNo, item.OrderNoAssigned);
                }
            }
            return personList;
        }

        /// <summary>
        /// Read file for Member Type
        /// </summary>
        /// <returns></returns>
        public List<Person> ReadFileMemberType()
        {
            List<Person> personList = new List<Person>();

            int row = 0;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(@PATH_FROM_MEMBER_TYPE, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read()) //Each ROW
                    {
                        if (row == 0)//read column names
                        {
                            for (int column = 0; column < reader.FieldCount; column++)
                            {
                                if (getValue(reader, column) == GetName(Email)) EmailCol = column;
                                if (getValue(reader, column) == GetName(MemberType)) MemberTypeCol = column;
                            }
                        }
                        else
                        {
                            Person person = new Person();
                            for (int column = 0; column < reader.FieldCount; column++)
                            {
                                if (column == EmailCol) person.Email = getValue(reader, column);
                                if (column == MemberTypeCol) person.MemberType = getValue(reader, column);  
                            }
                            personList.Add(person);
                        }
                        row++;
                    }
                }
            }
            return personList;
        }

        /// <summary>
        /// Read file with existing pairs of mentor and mentee
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, string> ReadFileMentorMentee()
        {
            List<Person> personList = new List<Person>();
            Dictionary<string, string> mentorMenteeDic = new Dictionary<string, string>();

            int row = 0;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(@PATH_FROM_MENTOR_MENTEE, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read()) //Each ROW
                    {
                        if (row != 0)//don't read the title
                        {    
                            mentorMenteeDic.Add(getValue(reader, 0), getValue(reader, 1));   
                        }
                        row++;
                    }
                }
            }
            return mentorMenteeDic;
        }

        /// <summary>
        /// Concat subfunction values 
        /// </summary>
        /// <param name="subfunctionSave"></param>
        /// <param name="subfunctionNew"></param>
        /// <returns></returns>
        public string SubfunctionAdd(string subfunctionSave, string subfunctionNew)
        {
            string SEPARATOR = "|";
            if (subfunctionNew != null)
            {
                if (subfunctionSave == null)
                    subfunctionSave = subfunctionNew;
                else
                    subfunctionSave = subfunctionSave + SEPARATOR + subfunctionNew;
           
                string tmp = null;
                //order by name
                var subfunctionList = subfunctionSave.Split(SEPARATOR).ToList().OrderBy(t => t.ToString());
                foreach (var item in subfunctionList)
                {
                    if (tmp == null)
                        tmp = item;
                    else
                        tmp = tmp + SEPARATOR + item;
                }
                subfunctionSave = tmp;
            }

            return subfunctionSave;
        }

        /// <summary>
        /// Get value from excel file
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public static string getValue(IExcelDataReader reader, int column)
        {
            string val = null;
            if (reader.GetValue(column) != null)
            {
                val = reader.GetValue(column).ToString();
            }
            return val;
        }

        /// <summary>
        /// Create file with log
        /// </summary>
        /// <param name="text"></param>
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

        /// <summary>
        /// Get settings from configuration file
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public string GetName(string value)
        {
            return configuration.GetValue<string>("MySettings:" + value);
        }
    }
}
