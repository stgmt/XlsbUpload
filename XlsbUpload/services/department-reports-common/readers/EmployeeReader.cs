using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using XlsbUpload.models;
using XlsbUpload.services;
using XlsbUpload.services.department_reports_common;

namespace XlsbUpload
{
    internal class EmployeeReader
    {

        EmployeeParserService _parserService;

        internal EmployeeReader()
        {
            _parserService = new EmployeeParserService();
        }

        public IEnumerable<Employee> Read(string filePath)
        {
            var result = new List<Employee>();

            var pageName = "Сотрудники";

            List<Employee> employees = new List<Employee>();

            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(filePath, ReadOnly: true);

            // Предполагаем, что данные начинаются с первой строки в первом листе
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
                var empls = _parserService.ParseEmployeePage(worksheet);
           
                if (!empls.Any())
                {
                    throw new Exception($"ошибка. не спарсилась страница {pageName}");
                }
                result.AddRange(empls);
            }
            else
            {
                throw new Exception("Ошибка: не удалось открыть файл.");
            }

            workbook.Close(false);
            excelApp.Quit();

            return result;
        }
    }
}
