using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using XlsbUpload.models;
using XlsbUpload.services.department_reports_common;

namespace XlsbUpload.services
{
    internal class DepartmentReader
    {

        DepartmentParserService _departmentParserService;
        public DepartmentReader()
        {
            _departmentParserService = new DepartmentParserService();
        }

        public IEnumerable<Department> Read(string filePath)
        {
            var result = new List<Department>();

            var pageName = "Отделы";

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
                var departs = _departmentParserService.ParseDepartmentPage(worksheet);
        
                if (!departs.Any())
                {
                    throw new Exception($"ошибка. не спарсилась страница {pageName}");
                }
                result.AddRange(departs);
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
