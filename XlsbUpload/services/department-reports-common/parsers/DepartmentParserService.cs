using OfficeOpenXml;
using System.Collections.Generic;
using XlsbUpload.models;

namespace XlsbUpload.services.department_reports_common
{
    internal class DepartmentParserService
    {
        internal IEnumerable<Department> ParseDepartmentPage(ExcelWorksheet worksheet)
        {
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++) // Начинаем считывать с 2 строки, предполагая, что первая строка - заголовок
            {
                string departmentId = worksheet.Cells[row, 1].GetValue<string>(); // Первый столбец - идентификатор отдела
                string departmentName = worksheet.Cells[row, 2].GetValue<string>(); // Второй столбец - наименование отдела

                yield return new Department { IdDepartment = departmentId, DepartmentName = departmentName };
            }

        }
    }
}
