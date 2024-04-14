using System.Collections.Generic;


namespace XlsbUpload.models
{
    internal class DepartmentTaskReportRow
    {
        public string DepartmentName { get; set; }
        public IEnumerable<EmployeeTask> EmployeeTasks { get; set; }
    }
}
