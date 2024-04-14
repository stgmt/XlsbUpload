using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XlsbUpload.models;
using XlsbUpload.services;

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

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                if (workbook != null)
                {

                    var worksheet = workbook.Worksheets.FirstOrDefault(x => x.Name == pageName);
                    var employee = _parserService.ParseEmployeePage(worksheet);
                    if (!employee.Any())
                    {
                        throw new Exception($"ошибка. не спарсилась страница {pageName}");
                    }
                }
            }

            return employees;
        }
    }
}
