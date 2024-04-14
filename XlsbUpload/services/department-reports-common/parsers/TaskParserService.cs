using OfficeOpenXml;
using System.Collections.Generic;
using XlsbUpload.models;

namespace XlsbUpload.services.department_reports_common
{
    internal class TaskParserService
    {
        internal IEnumerable<EmployeeTask> ParseTasktPage(ExcelWorksheet worksheet)
        {
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++) // Начинаем считывать с 2 строки, предполагая, что первая строка - заголовок
            {
                var taskId = worksheet.Cells[row, 1].GetValue<string>(); // Первый столбец - идентификатор задачи
                var tin = worksheet.Cells[row, 2].GetValue<string>(); // Второй столбец - табельный номер

                yield return new EmployeeTask { IdTask = taskId, TIN = tin };
            }
        }
    }
}
