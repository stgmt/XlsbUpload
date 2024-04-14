using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
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

            var pageName = "Название_страницы"; // Укажите имя страницы в вашем Excel файле

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                if (workbook != null)
                {
                    var worksheet = workbook.Worksheets.FirstOrDefault(x => x.Name == pageName);
                    if (worksheet != null)
                    {
                        return _parserService.ParseTasktPage(worksheet);
                    }
                    else
                    {
                        throw new Exception($"Ошибка: страница '{pageName}' не найдена в файле.");
                    }
                }
                else
                {
                    throw new Exception("Ошибка: не удалось открыть файл.");
                }
            }

        }
    }
}
