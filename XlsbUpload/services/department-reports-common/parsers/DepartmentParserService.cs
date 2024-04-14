using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using XlsbUpload.models;

namespace XlsbUpload.services.department_reports_common
{
    internal class DepartmentParserService
    {
        internal IEnumerable<Department> ParseDepartmentPage(Worksheet worksheet)
        {
            Range range = worksheet.UsedRange;

            int rowCount = range.Rows.Count;


            for (int row = 2; row <= rowCount; row++) // Начинаем считывать с 2 строки, предполагая, что первая строка - заголовок
            {
                Department dep = default;
                try
                {

                    string departmentId = (string)(range.Cells[row, 1] as Range).Value2.ToString();
                    string departmentName = (range.Cells[row, 2] as Range).Value2.ToString();

                    dep = new Department
                    {
                        IdDepartment = departmentId,
                        DepartmentName = departmentName
                    };
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Ошибка парсинга отделов. В строке {row} {ex.Message}");
                    Console.ResetColor();
                    continue;
                }

                yield return dep;
            }

        }
    }
}
