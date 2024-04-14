using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using XlsbUpload.models;

namespace XlsbUpload.services.department_reports_common
{
    internal class TaskParserService
    {
        internal IEnumerable<EmployeeTask> ParseTasktPage(Worksheet worksheet)
        {
            Range range = worksheet.UsedRange;

            int rowCount = range.Rows.Count;

            for (int row = 2; row <= rowCount; row++) // Начинаем считывать с 2 строки, так как первая строка содержит заголовки
            {
                EmployeeTask emplTask = default;
                try
                {
                    string taskId = (string)(range.Cells[row, 1] as Range).Value2.ToString();
                    string tin = (range.Cells[row, 2] as Range).Value2.ToString();

                    emplTask = new EmployeeTask
                    {
                        TIN = tin,
                        IdTask = taskId
                    };
                }
                catch (Exception ex)
                {

                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Ошибка парсинга заданий. В строке {row} {ex.Message}");
                    Console.ResetColor();
                    continue;
                }

                yield return emplTask;
            }
        }
    }
}
