using OfficeOpenXml;
using System;
using System.Collections.Generic;
using XlsbUpload.models;

namespace XlsbUpload.services
{
    internal class EmployeeParserService
    {

        internal IEnumerable<Employee> ParseEmployeePage(ExcelWorksheet worksheet)
        {
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++) // Начинаем считывать с 2 строки, так как первая строка содержит заголовки
            {
                Employee employee = default;
                try
                {
                    employee = new Employee
                    {
                        TIN = worksheet.Cells[row, 1].Text,
                        LastName = worksheet.Cells[row, 2].Text,
                        FirstName = worksheet.Cells[row, 3].Text,
                        MiddleName = worksheet.Cells[row, 4].Text,
                        DateOfBirth = DateTime.Parse(worksheet.Cells[row, 5].Text),
                        DepartmentId = worksheet.Cells[row, 6].Text
                    };
                }
                catch (Exception ex)
                {

                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Ошибка парсинга сотрудника. В строке {row} {ex.Message}");
                    Console.ResetColor();
                }

                yield return employee;
            }
        }
    }
}
