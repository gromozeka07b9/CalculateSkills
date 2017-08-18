using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCalc
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() == 6)
            {
                string pathToSource = args[0];
                string pageName = args[1];
                ColumnsParameters columnsParameters = new ColumnsParameters();
                columnsParameters.columnNameBasicStart = args[2];
                columnsParameters.columnNameBasicEnd = args[3];
                columnsParameters.columnNameSpecialStart = args[4];
                columnsParameters.columnNameSpecialEnd = args[5];

                DataSet ds = fillTableFromFile(pathToSource, pageName);
                List<Employee> listEmpl = fillEmployeeList(ds, pageName, columnsParameters);
                foreach (var item in listEmpl)
                {
                    item.SocialPoints = item.SocialPoints / item.CountOfLead;
                    item.TechPoints = item.TechPoints / item.CountOfLead;
                }

                fillAveragePoints(listEmpl);

                string pathToAvgResult = Path.GetFileNameWithoutExtension(pathToSource) + "_avg_" + DateTime.Now.ToShortDateString() + ".txt";
                exportAverageToText(listEmpl, pathToAvgResult);
                Console.WriteLine("Усредненные результаты выгружены в " + pathToAvgResult);

                string pathToResult = Path.GetFileNameWithoutExtension(pathToSource) + "_" + DateTime.Now.ToShortDateString() + ".txt";
                exportToText(listEmpl, pathToResult);
                Console.WriteLine("Результаты выгружены в " + pathToResult);

                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Некорректное количество параметров, должен быть формат: <путь к исходному файлу> <название листа> <имя колонки базовых начало> <имя колонки базовых конец> <имя колонки спец начало> <имя колонки спец конец>");
                Console.ReadLine();
            }
        }

        private static void exportToText(List<Employee> employeeList, string pathToTextFile)
        {
            StringBuilder sb = new StringBuilder();
            using (StreamWriter writer = new StreamWriter(pathToTextFile))
            {
                foreach (var employeeItem in employeeList)
                {
                    writer.WriteLine("{0}, базовые навыки:{1}, специальные навыки:{2}, количество оценивающих:{3}", employeeItem.Name, employeeItem.SocialPoints, employeeItem.TechPoints, employeeItem.CountOfLead);
                }
            }
        }

        private static void exportAverageToText(List<Employee> employeeList, string pathToTextFile)
        {
            StringBuilder sb = new StringBuilder();
            using (StreamWriter writer = new StreamWriter(pathToTextFile))
            {
                foreach (var employeeItem in employeeList)
                {
                    writer.WriteLine("Сотрудник: {0}", employeeItem.Name);
                    float sum = 0;
                    foreach (var avg in employeeItem.ListAveragePoints)
                    {
                        writer.WriteLine("Навык: [{0}], усредненный показатель: [{1}]", avg.Item1, avg.Item2);
                        sum += avg.Item2;
                    }
                    writer.WriteLine("Сумма: {0}", sum);
                }
            }
        }

        private static void fillAveragePoints(List<Employee> listEmployee)
        {
 
            foreach(Employee employee in listEmployee)
            {
                employee.ListSummaryPoints = new List<Tuple<string, int>>();
                foreach (var listPoints in employee.ListPoints)
                {
                    foreach (var item in listPoints)
                    {
                        Tuple<string, int> currentPoint;
                        if (!employee.ListSummaryPoints.Exists(e => e.Item1 == item.Item1))
                        {
                            currentPoint = new Tuple<string, int>(item.Item1, item.Item2);
                            employee.ListSummaryPoints.Add(currentPoint);
                        }
                        else
                        {
                            currentPoint = employee.ListSummaryPoints.First(e => e.Item1 == item.Item1);
                            int point = currentPoint.Item2 + item.Item2;
                            employee.ListSummaryPoints.Remove(currentPoint);
                            currentPoint = new Tuple<string, int>(currentPoint.Item1, point);
                            employee.ListSummaryPoints.Add(currentPoint);
                        }

                    }
                }
            }

            foreach(Employee employee in listEmployee)
            {
                employee.ListAveragePoints = new List<Tuple<string, float>>();
                foreach(var summary in employee.ListSummaryPoints)
                {
                    employee.ListAveragePoints.Add(new Tuple<string, float>(summary.Item1, summary.Item2/employee.CountOfLead));                    
                }
            }
        }

        private static List<Employee> fillEmployeeList(DataSet ds, string pageName, ColumnsParameters columnsParameters)
        {

            List<Employee> emplList = new List<Employee>();
            List<string> basicColumns = getColumnNames(ds.Tables[pageName].Columns, ds.Tables[pageName].Columns[columnsParameters.columnNameBasicStart].Ordinal, ds.Tables[pageName].Columns[columnsParameters.columnNameBasicEnd].Ordinal);
            List<string> specialColumns = getColumnNames(ds.Tables[pageName].Columns, ds.Tables[pageName].Columns[columnsParameters.columnNameSpecialStart].Ordinal, ds.Tables[pageName].Columns[columnsParameters.columnNameSpecialEnd].Ordinal);

            foreach (DataRow row in ds.Tables[pageName].Rows)
            {
                string currentEmployeeName = row["Оцениваемый сотрудник"].ToString();
                Console.WriteLine("basic points:");
                PointsByPerson personBasicPoints = summarizePoints(row, basicColumns, currentEmployeeName);
                Console.WriteLine("special points:");
                PointsByPerson personSpecialPoints = summarizePoints(row, specialColumns, currentEmployeeName);
                Employee currentEmployee;
                if (!emplList.Exists(e => e.Name == currentEmployeeName))
                {
                    currentEmployee = new Employee();
                    currentEmployee.CountOfLead = 1;
                    currentEmployee.Name = currentEmployeeName;
                    currentEmployee.ListPoints = new List<List<Tuple<string, int>>>();
                    emplList.Add(currentEmployee);
                }
                else
                {
                    currentEmployee = emplList.First(e => e.Name == currentEmployeeName);
                    currentEmployee.CountOfLead++;
                }
                currentEmployee.SocialPoints += personBasicPoints.Count;
                currentEmployee.TechPoints += personSpecialPoints.Count;
                List<Tuple<string, int>> listPoints = new List<Tuple<string, int>>();
                foreach (PointsBySkill basicItem in personBasicPoints.PointsList)
                {
                    listPoints.Add(new Tuple<string, int>(basicItem.column, basicItem.count));
                }
                foreach (PointsBySkill specialItem in personSpecialPoints.PointsList)
                {
                    listPoints.Add(new Tuple<string, int>(specialItem.column, specialItem.count));
                }
                currentEmployee.ListPoints.Add(listPoints);
            }
            return emplList;
        }

        private static PointsByPerson summarizePoints(DataRow row, List<string> columns, string employeeName)
        {
            //int points = 0;
            PointsByPerson person = new PointsByPerson();
            person.PointsList = new List<PointsBySkill>();
            //List<PointsBySkill> listPoints = new List<PointsBySkill>();
            foreach (var column in columns)
            {
                int count = getPointFromString(row[column].ToString());
                //points += count;
                person.Count += count;
                person.PointsList.Add(new PointsBySkill(employeeName, column, count));
            }
            return person;
        }

        private static int getPointFromString(string cellValue)
        {
            int result = 0;
            if(cellValue.Length>3)
            {
                string pointsDraft = cellValue.Substring(cellValue.Length - 4);
                string points = pointsDraft.Replace("(", "").Replace(")", "").Trim();
                result = Convert.ToInt32(points);
            }
            return result;
        }

        private static List<string> getColumnNames(DataColumnCollection columns, int columnFrom, int columnTo)
        {
            List<string> resultColumns = new List<string>();
            for (int i = columnFrom; i <= columnTo; i++)
            {
                resultColumns.Add(columns[i].ColumnName);
            }
            return resultColumns;                
        }

        private static DataSet fillTableFromFile(string path, string tableName)
        {
            DataSet ds = new DataSet();
            Dictionary<string, string> props = new Dictionary<string, string>();
            props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            props["Data Source"] = path;
            props["Extended Properties"] = "Excel 8.0";

            StringBuilder sb = new StringBuilder();
            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }
            string properties = sb.ToString();

            using (OleDbConnection conn = new OleDbConnection(properties))
            {
                conn.Open();
                using (OleDbDataAdapter da = new OleDbDataAdapter(
                    "SELECT * FROM [" + tableName + "$]", conn))
                {
                    DataTable dt = new DataTable(tableName);
                    da.Fill(dt);
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }

        struct ColumnsParameters
        {
            public string columnNameBasicStart { get; internal set; }
            public string columnNameBasicEnd { get; internal set; }
            public string columnNameSpecialStart { get; internal set; }
            public string columnNameSpecialEnd { get; internal set; }
        }
    }
}
