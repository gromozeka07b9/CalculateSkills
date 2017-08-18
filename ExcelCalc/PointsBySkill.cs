namespace ExcelCalc
{
    internal class PointsBySkill
    {
        public string column;
        public string employeeName;
        public int count;

        public PointsBySkill(string employeeName, string column, int count)
        {
            this.employeeName = employeeName;
            this.column = column;
            this.count = count;
        }
    }
}