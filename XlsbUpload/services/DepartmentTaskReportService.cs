using System.Collections.Generic;
using System.Linq;
using XlsbUpload.models;
using XlsbUpload.services.reports_common;

namespace XlsbUpload.services
{
    internal class DepartmentTaskReportService : ReaderBase
    {
        string[] _docsPathArgs;
        EmployeeReader _employeeReader;
        DepartmentReader _departmentReader;
        TaskReader _taskReader;
        DocumentFileWriter _writer;

        public DepartmentTaskReportService(string[] args)
        {
            _employeeReader = new EmployeeReader();
            _departmentReader = new DepartmentReader();
            _taskReader = new TaskReader();
            _writer = new DocumentFileWriter();

            _docsPathArgs = args;
        }

        public IEnumerable<bool> MakeReport()
        {
            var reports = BuildReports();

            foreach (var item in reports)
            {
                yield return _writer.Write(item);
            }
        }

        private IEnumerable<IEnumerable<DepartmentTaskReportRow>> BuildReports()
        {
            var reports = new List<List<DepartmentTaskReportRow>>();

            var docsPath = GetDocumentPath(_docsPathArgs);

            foreach (var docPath in docsPath)
            {
                var employees = _employeeReader.Read(docPath);
                var departments = _departmentReader.Read(docPath);
                var employeeTasks = _taskReader.Read(docPath);

                var employeeTasksGrouped = employeeTasks.GroupBy(task => task.TIN);

                var departmentBuildResult = departments.Select(department =>
                {
                    var departmentTasks = employees
                        .Where(emp => emp.DepartmentId == department.IdDepartment)
                        .SelectMany(emp => employeeTasksGrouped.FirstOrDefault(group => group.Key == emp.TIN))
                        .ToList();

                    // Обновляем задачи с именем и фамилией сотрудника
                    departmentTasks.ForEach(task =>
                    {
                        var matchingEmployee = employees.FirstOrDefault(emp => emp.TIN == task.TIN);
                        if (matchingEmployee != null)
                        {
                            task.FirstName = matchingEmployee.FirstName;
                            task.LastName = matchingEmployee.LastName;
                        }
                    });

                    return new DepartmentTaskReportRow
                    {
                        DepartmentName = department.DepartmentName,
                        EmployeeTasks = departmentTasks
                    };
                }).ToList();

                reports.Add(departmentBuildResult);
            }

            return reports;
        }

    }
}
