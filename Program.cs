using ClosedXML.Excel;

namespace ReadMarks
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelReader reader = new ExcelReader();
            List<Student> students = reader.ReadStudents("C:\\temp\\tmp.xlsm");
            reader.ExportStudentsToExcel(students, "C:\\temp\\output.xlsm");
        }


        public class ExcelReader
        {
            public List<Student> ReadStudents(string filePath)
            {
                var students = new List<Student>();

                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    var rows = worksheet.RangeUsed().RowsUsed().Skip(11);

                    foreach (var row in rows)
                    {
                        if (!row.Cell(2).IsEmpty())
                        {
                            var student = new Student
                            {
                                Class = row.Cell(1).GetValue<string>(),
                                StudentNumber = row.Cell(2).GetValue<int>(),
                                Achternaam = row.Cell(3).GetValue<string>(),
                                Roepnaam = row.Cell(4).GetValue<string>(),
                                Mark = row.Cell(5).GetValue<double>()
                            };
                            students.Add(student);
                        }

                    }
                }

                var orderedStudents = students.OrderBy(a => a.Achternaam).ToList();
                return orderedStudents;
            }

            public void ExportStudentsToExcel(List<Student> students, string filePath)
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Students");

                    worksheet.Cell(1, 1).InsertTable(students, "StudentTable", true);

                    worksheet.Columns().AdjustToContents();

                    workbook.SaveAs(filePath);
                }
            }
        }
    }
}