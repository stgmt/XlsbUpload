using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using XlsbUpload.models;

namespace XlsbUpload.services
{
    internal class EmployeeParserService
    {

        internal IEnumerable<Employee> ParseEmployeePage(Worksheet worksheet)
        {
            Range range = worksheet.UsedRange;
            int rowCount = range.Rows.Count;

            for (int row = 2; row <= rowCount; row++) // Начинаем считывать с 2 строки, так как первая строка содержит заголовки
            {
                Employee employee = default;
                try
                {
                    string tin = (range.Cells[row, 1] as Range).Value2.ToString();
                    string lastName = (range.Cells[row, 2] as Range).Value2.ToString();
                    string firstName = (range.Cells[row, 3] as Range).Value2.ToString();
                    string middleName = (range.Cells[row, 4] as Range).Value2.ToString();
                    DateTime birthDate = DateTime.FromOADate((double)(range.Cells[row, 5] as Range).Value2);
                    string departmentId = (string)(range.Cells[row, 6] as Range).Value2.ToString();

                    employee = new Employee
                    {
                        TIN = tin,
                        LastName = lastName,
                        FirstName = firstName,
                        MiddleName = middleName,
                        DateOfBirth = birthDate,
                        DepartmentId = departmentId
                    };
                }
                catch (Exception ex)
                {

                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Ошибка парсинга сотрудника. В строке {row} {ex.Message}");
                    Console.ResetColor();
                    continue;
                }

                 yield return employee;
            }

        }
    }
}
