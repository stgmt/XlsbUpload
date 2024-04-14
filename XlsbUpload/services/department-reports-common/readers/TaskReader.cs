using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using XlsbUpload.models;
using XlsbUpload.services.department_reports_common;

namespace XlsbUpload.services.reports_common
{
    internal class TaskReader
    {

        TaskParserService _parserService;

        internal TaskReader()
        {
            _parserService = new TaskParserService();
        }

        public IEnumerable<EmployeeTask> Read(string filePath)
        {
            var employeeTasks = new List<EmployeeTask>();

            var pageName = "Задачи"; // Укажите имя страницы в вашем Excel файле


            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(filePath, ReadOnly: true);

            Worksheet worksheet = null;
            foreach (Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == pageName)
                {
                    worksheet = sheet;
                    break;
                }
            }

            if (worksheet != null)
            {
                var tasks = _parserService.ParseTasktPage(worksheet);
                if (!tasks.Any())
                {
                    throw new Exception($"ошибка. не спарсилась страница {pageName}");
                }
                employeeTasks.AddRange(tasks);
            }
            else
            {
                throw new Exception("Ошибка: не удалось открыть файл.");
            }

            workbook.Close(false);
            excelApp.Quit();
            return employeeTasks;
        }
    }
}
