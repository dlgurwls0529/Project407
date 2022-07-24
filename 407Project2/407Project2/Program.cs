using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.Linq;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace _407Project2
{
    interface Student
    {

    }

    class OriginStudent : Student
    {
        private string station;
        private string classForm;
        private string level;
        private string teacher;
        private string name;
        private string schoolAndGrade;
        private string phoneNumber;
        private string cardNumber;

        public string Level { get => level; set => level = value; }
        public string Teacher { get => teacher; set => teacher = value; }
        public string Name { get => name; set => name = value; }
        public string SchoolAndGrade { get => schoolAndGrade; set => schoolAndGrade = value; }
        public string PhoneNumber { get => phoneNumber; set => phoneNumber = value; }
        public string CardNumber { get => cardNumber; set => cardNumber = value; }
        public string Station { get => station; set => station = value; }
        public string ClassForm { get => classForm; set => classForm = value; }
    }

    class BasicStudent : Student
    {
        private string station;
        private string classForm;
        private string level;
        private string teacher;
        private string name;
        private string studentNumber;
        private string password;

        public string Station { get => station; set => station = value; }
        public string ClassForm { get => classForm; set => classForm = value; }
        public string Level { get => level; set => level = value; }
        public string Teacher { get => teacher; set => teacher = value; }
        public string Name { get => name; set => name = value; }
        public string StudentNumber { get => studentNumber; set => studentNumber = value; }
        public string Password { get => password; set => password = value; }
    }

    interface Preprocessor
    {
        public void fit(Student student);

        public Student transform();
    }

    class PreprocessorOriginToBasic : Preprocessor
    {
        private OriginStudent OriginStudent;

        public void fit(Student student)
        {
            if(student is OriginStudent)
            {
                OriginStudent = (OriginStudent) student;
            }
            else
            {
                throw new InvalidCastException();
            }
            
        }
        // (관명, 반형태, 레벨명, 담임, 이름, 학교/학년, 핸드폰번호, 카드번호)
        // (관명, 반형태, 레벨명, 담임, 이름, 카드번호, 핸드폰번호)
        // (관명, 반형태, 레벨명, 담임, 이름, 학생번호, 비밀번호)
        public Student transform()
        {
            if(OriginStudent == null)
            {
                throw new NullReferenceException();
            }
            else
            {
                BasicStudent basicStudent = new BasicStudent();

                basicStudent.Station = getPreprocessedStation(
                    OriginStudent.Station);
                basicStudent.ClassForm = getPreprocessedClassForm(
                    OriginStudent.ClassForm);
                basicStudent.Level = getPreprocessedLevel(
                    OriginStudent.Level);
                basicStudent.Teacher = getPreprocessedTeacher(
                    OriginStudent.Teacher);
                basicStudent.Name = getPreprocessedName(
                    OriginStudent.Name);
                basicStudent.StudentNumber = getPreprocessedStudentNumber(
                    OriginStudent.CardNumber);
                basicStudent.Password = getPreprocessedPassword(
                    OriginStudent.PhoneNumber);

                return basicStudent;
            }
        }

        private string getPreprocessedStation(string station)
        {
            return station;
        }

        private string getPreprocessedClassForm(string classForm)
        {
            return classForm;
        }

        private string getPreprocessedLevel(string level)
        {
            return level;
        }

        private string getPreprocessedTeacher(string teacher)
        {
            return teacher;
        }

        private string getPreprocessedName(string name)
        {
            return name;
        }

        private string getPreprocessedStudentNumber(string cardNumber)
        {
            return cardNumber;
        }

        private string getPreprocessedPassword(string phoneNumber)
        {
            return phoneNumber;
        }
    }

    class Program
    {

        static void Main(string[] args)
        {
            // https://stackoverflow.com/questions/10704582/read-an-xslx-file-and-convert-to-list

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Console.WriteLine("파일 읽기 시작.");

            var ep_mokdong = new ExcelPackage(new FileInfo(@"C:\Users\dlgur\mokdongData.xlsx"));
            var ws_mokdong = ep_mokdong.Workbook.Worksheets["Sheet1"];
            List<OriginStudent> domains_mokdong = WsToList(ref ws_mokdong, "목동관", "국제관 초등 0");

            var ep_premium = new ExcelPackage(new FileInfo(@"C:\Users\dlgur\premiumData.xlsx"));
            var ws_premium = ep_premium.Workbook.Worksheets["Sheet1"];
            List<OriginStudent> domains_preminum = WsToList(ref ws_premium, "프리미엄관", "(프) 기타 등등");

            var ep_main = new ExcelPackage(new FileInfo(@"C:\Users\dlgur\mainData.xlsx"));
            var ws_main = ep_main.Workbook.Worksheets["Sheet1"];
            List<OriginStudent> domains_main = WsToList(ref ws_main, "본관", "(특) 특목 테스트");

            Console.WriteLine("파일 읽기 완료.");

            List<OriginStudent> domains_total = new List<OriginStudent>();
            domains_total.AddRange(domains_main);
            domains_total.AddRange(domains_preminum);
            domains_total.AddRange(domains_mokdong);

            WriteStudentData<OriginStudent>(ref domains_total);

            /*for (int i = 0; i < domains_total.Count; i++)
            {
                OriginStudent originStudent = new OriginStudent();
                originStudent = domains_total[i];
                PreprocessorOriginToBasic preprocessor = new PreprocessorOriginToBasic();
                preprocessor.fit(originStudent);
                BasicStudent basicStudent = preprocessor.transform() as BasicStudent;

                if (basicStudent != null)
                {
                    *//*Console.WriteLine("관명 : " + originStudent.Station);
                    Console.WriteLine("반 형태 : " + originStudent.ClassForm);
                    Console.WriteLine("레벨 : " + originStudent.Level);
                    Console.WriteLine("선생님 : " + originStudent.Teacher);
                    Console.WriteLine("이름 : " + originStudent.Name);
                    Console.WriteLine("학년/학교 : " + originStudent.SchoolAndGrade);
                    Console.WriteLine("핸드폰 번호 : " + originStudent.PhoneNumber);
                    Console.WriteLine("카드 번호 : " + originStudent.CardNumber);

                    Console.WriteLine(" ");*//*

                    Console.WriteLine("관명 : " + basicStudent.Station);
                    Console.WriteLine("반 형태 : " + basicStudent.ClassForm);
                    Console.WriteLine("레벨 : " + basicStudent.Level);
                    Console.WriteLine("선생님 : " + basicStudent.Teacher);
                    Console.WriteLine("이름 : " + basicStudent.Name);
                    Console.WriteLine("학생 번호 : " + basicStudent.StudentNumber);
                    Console.WriteLine("비밀 번호 : " + basicStudent.Password);

                    Console.WriteLine(" ");
                }
            }
*/

        }

        static List<OriginStudent> WsToList(ref OfficeOpenXml.ExcelWorksheet ws, string station, string classForm)
        {
            var domains = new List<OriginStudent>();

            string currentLevel = "";
            string currentTeacher = "";
            int currentRowEnd = 0;

            int col_len = ws.Dimension.End.Column;
            int row_len = ws.Dimension.End.Row;

            for (int i = 1; i <= col_len; i++)
            {
                // 해당 행이 빈 데이터인지 확인하고 점프
                if (IsCellEmpty(ws, i))
                {
                    continue;
                }

                // 1, 5, 9 ..번째 컬럼마다 레벨과 담임을 갱신
                if (i % 4 == 1)
                {
                    currentLevel = ws.Cells[1, i].Value.ToString();
                    currentTeacher = ws.Cells[2, i].Value.ToString();
                }

                // 반복자 i가 컬럼 '이름' 을 가리킬 경우
                // 학생을 전부 이름만 설정하여 domain에 추가
                if (i % 4 == 1)
                {
                    currentRowEnd = ws.Dimension.End.Row;

                    // 학생 정보가 시작하는 4행부터 정보를 저장한다.
                    for (int j = 4; j <= currentRowEnd; j++)
                    {
                        if (ws.Cells[j, i].Value == null)
                        {
                            currentRowEnd = j - 1;
                            break;
                        }
                        else
                        {
                            OriginStudent originStudent = new OriginStudent();
                            originStudent.Station = station;
                            originStudent.ClassForm = classForm;
                            originStudent.Level = currentLevel;
                            originStudent.Teacher = currentTeacher;
                            originStudent.Name = ws.Cells[j, i].Value.ToString();
                            domains.Add(originStudent);
                        }

                    }

                }
                // 반복자 i가 컬럼 '학교/학년' 을 가리킬 경우
                // domain 의 해당 필드를 초기화
                else if (i % 4 == 2)
                {
                    int k = currentRowEnd;

                    for (int j = domains.Count - 1; j >= domains.Count - (currentRowEnd - 3); j--)
                    {
                        if (ws.Cells[k, i].Value != null)
                        {
                            domains[j].SchoolAndGrade = ws.Cells[k, i].Value.ToString();
                        }
                        else
                        {
                            domains[j].SchoolAndGrade = "";
                        }

                        k--;
                    }
                }
                // 반복자 i가 컬럼 '부모핸펀' 을 가리킬 경우
                else if (i % 4 == 3)
                {
                    int k = currentRowEnd;

                    for (int j = domains.Count - 1; j >= domains.Count - (currentRowEnd - 3); j--)
                    {
                        if (ws.Cells[k, i].Value != null)
                        {
                            domains[j].PhoneNumber = ws.Cells[k, i].Value.ToString();
                        }
                        else
                        {
                            domains[j].PhoneNumber = "";
                        }

                        k--;
                    }
                }
                // 반복자 i가 컬럼 '카드번호' 을 가리킬 경우
                else if (i % 4 == 0)
                {
                    int k = currentRowEnd;

                    for (int j = domains.Count - 1; j >= domains.Count - (currentRowEnd - 3); j--)
                    {
                        if (ws.Cells[k, i].Value != null)
                        {
                            domains[j].CardNumber = ws.Cells[k, i].Value.ToString();
                        }
                        else
                        {
                            domains[j].CardNumber = "";
                        }

                        k--;
                    }
                }

            }

            return domains;
        }

        static bool IsCellEmpty(OfficeOpenXml.ExcelWorksheet ws, int col)
        {
            // 하나라도 채워져 있으면 false
            // 전부 다 비워져 있으면 true

            bool result = true;
            int row_len = ws.Dimension.End.Row;

            for (int i = 4; i <= row_len; i++)
            {
                result = result && (ws.Cells[i, col].Value == null);
            }

            return result;
        }

        static void WriteStudentData<T>(ref List<T> domains)
        {
            Console.WriteLine("파일 작성 시작...");

            string fileName = "studentDataTable_chita_" +
                DateTime.Now.ToString("yyyyMMdd") +
                DateTime.Now.ToString("HHmmss") +
                ".xlsx";

            ExcelPackage pck = new ExcelPackage();
            var wsheet = pck.Workbook.Worksheets.Add("Sheet1");

            TypeInfo typeInfo = typeof(T).GetTypeInfo();
            IEnumerable<PropertyInfo> pList = typeInfo.DeclaredProperties;
            
            // row
            for(int i = 1; i <= domains.Count; i++)
            {
                int j = 1;
                // col
                foreach(PropertyInfo p in pList)
                {
                    if (i == 1)
                    {
                        wsheet.Cells[i, j].Value = p.Name;
                    }
                    else
                    {
                        wsheet.Cells[i, j].Value = p.GetValue(domains[i-1]);
                    }

                    j++;
                }
            }

            FileStream fileStream = File.Create(@"C:\Users\dlgur\" + fileName);
            fileStream.Close();

            File.WriteAllBytes(@"C:\Users\dlgur\" + fileName, pck.GetAsByteArray());

            pck.Dispose();

            Console.WriteLine("파일 작성 완료...");
        }
    }
}
