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

                // Группируем задачи по сотрудникам
                var employeeTasksGrouped = employeeTasks.GroupBy(task => task.TIN);

                var departmentBuildResult = new List<DepartmentTaskReportRow>();
                // Для каждого отдела создаем строку отчета
                foreach (var department in departments)
                {
                    var departmentRow = new DepartmentTaskReportRow
                    {
                        DepartmentName = department.DepartmentName,
                        EmployeeTasks = new List<EmployeeTask>()
                    };

                    // Добавляем задачи для сотрудников в отделе
                    foreach (var employee in employees.Where(emp => emp.DepartmentId == department.IdDepartment))
                    {
                        if (employeeTasksGrouped.Any(group => group.Key == employee.TIN))
                        {
                            departmentRow.EmployeeTasks = departmentRow.EmployeeTasks.Concat(employeeTasksGrouped.First(group => group.Key == employee.TIN));
                        }
                    }

                    reports.Add(departmentBuildResult);
                }
                reports.Add(departmentBuildResult);
            }


            return reports;
        }

    }
}
