using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XlsbUpload.models;


namespace XlsbUpload.services.reports_common
{
    internal class DocumentFileWriter
    {
        internal bool Write(IEnumerable<DepartmentTaskReportRow> docModel)
        {
            try
            {
                // Создаем приложение Word
                Application wordApp = new Application();
                // Создаем новый документ Word
                Document doc = wordApp.Documents.Add();

                // Добавляем заголовок в документ
                Paragraph title = doc.Paragraphs.Add();
                title.Range.Text = "Отчет по загрузке";
                title.Range.Font.Bold = 1;
                title.Range.Font.Size = 14;
                title.Format.SpaceAfter = 10; // Пробел после заголовка

                // Добавляем таблицу
                Table table = doc.Tables.Add(title.Range, docModel.Count() + 1, 2); // Количество строк = количество отделов + 1 для заголовка, количество столбцов = 2
                table.Cell(1, 1).Range.Text = "Отдел";
                table.Cell(1, 2).Range.Text = "Количество задач";

                int row = 2; // Начинаем с второй строки, первая строка - заголовок
                foreach (var departmentRow in docModel)
                {
                    table.Cell(row, 1).Range.Text = departmentRow.DepartmentName;
                    table.Cell(row, 2).Range.Text = departmentRow.EmployeeTasks.Count().ToString();
                    row++;
                }

                // Сохраняем документ
                object fileName = Directory.GetCurrentDirectory() + $"reports/report-{Guid.NewGuid()}.docx"; // Замените путь на свой
                doc.SaveAs2(ref fileName);

                // Закрываем приложение Word
                wordApp.Quit();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ошибка записи doc файла {ex.Message}");
                return false;
            }
        }
    }
}
