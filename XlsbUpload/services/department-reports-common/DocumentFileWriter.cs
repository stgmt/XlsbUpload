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
                var wordApp = new Application();
                var doc = CreateDocument(wordApp);

                AddTitle(doc);
                CreateTable(doc, docModel);

                SaveDocument(doc);
                CloseWordApplication(wordApp);

                return true;
            }
            catch (Exception ex)
            {
                HandleError(ex);
                return false;
            }
        }

        private Document CreateDocument(Application wordApp)
        {
            return wordApp.Documents.Add();
        }

        private void AddTitle(Document doc)
        {
            var title = doc.Paragraphs.Add();
            title.Range.Text = "Отчет по загрузке";
            title.Range.Font.Bold = 1;
            title.Range.Font.Size = 16;
            title.Format.SpaceAfter = 10;
        }

        private void CreateTable(Document doc, IEnumerable<DepartmentTaskReportRow> docModel)
        {
            var table = doc.Tables.Add(doc.Range(), CalculateTableRows(docModel), 2);
            FormatTable(table);
            AddTableHeader(table);
            PopulateTableData(table, docModel);
        }

        private int CalculateTableRows(IEnumerable<DepartmentTaskReportRow> docModel)
        {
            return docModel.Sum(d => d.EmployeeTasks.Count()) + docModel.Count() + 1;
        }

        private void FormatTable(Table table)
        {
            table.Borders.Enable = 1;
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth050pt;
            table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineWidth = WdLineWidth.wdLineWidth050pt;
        }

        private void AddTableHeader(Table table)
        {
            var headerCells = table.Rows[1].Cells;
            foreach (Cell cell in headerCells)
            {
                FormatHeaderCell(cell);
            }
            headerCells[1].Range.Text = "Отдел";
            headerCells[2].Range.Text = "Количество задач";
        }

        private void FormatHeaderCell(Cell cell)
        {
            cell.Shading.BackgroundPatternColor = WdColor.wdColorGray50;
            cell.Range.Font.Color = WdColor.wdColorWhite;
            cell.Range.Font.Bold = 1;
            cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        }

        private void PopulateTableData(Table table, IEnumerable<DepartmentTaskReportRow> docModel)
        {
            int row = 2;
            foreach (var departmentRow in docModel)
            {
                AddDepartmentRow(table, departmentRow, ref row);
                AddEmployeeRows(table, departmentRow, ref row);
            }
        }

        private void AddDepartmentRow(Table table, DepartmentTaskReportRow departmentRow, ref int row)
        {
            table.Cell(row, 1).Range.Text = departmentRow.DepartmentName;
            table.Cell(row, 1).Range.Font.Bold = 1;
            table.Cell(row, 1).Shading.BackgroundPatternColor = WdColor.wdColorGray25;
            table.Cell(row, 2).Range.Text = departmentRow.EmployeeTasks.Count().ToString();
            table.Cell(row, 2).Shading.BackgroundPatternColor = WdColor.wdColorGray25;
            table.Cell(row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            row++;
        }

        private void AddEmployeeRows(Table table, DepartmentTaskReportRow departmentRow, ref int row)
        {
            var employeeTaskCount = departmentRow.EmployeeTasks
                .GroupBy(task => $"{task.FirstName} {task.LastName}")
                .Select(group => new
                {
                    FullName = group.Key,
                    TaskCount = group.Count()
                });

            foreach (var employee in employeeTaskCount)
            {
                AddEmployeeRow(table, employee, ref row);
            }
        }

        private void AddEmployeeRow(Table table, dynamic employee, ref int row)
        {
            table.Cell(row, 1).Range.Text = employee.FullName;
            table.Cell(row, 2).Range.Text = employee.TaskCount.ToString();
            table.Cell(row, 2).Range.Font.Bold = 0; //
            table.Cell(row, 1).Range.Font.Bold = 0; //
            table.Cell(row, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            row++;
        }

        private void SaveDocument(Document doc)
        {
            string fileName = Path.Combine(Directory.GetCurrentDirectory(), $"report-{Guid.NewGuid()}.docx");
            doc.SaveAs2(fileName);
        }

        private void CloseWordApplication(Application wordApp)
        {
            wordApp.Quit();
        }

        private void HandleError(Exception ex)
        {
            Console.WriteLine($"Ошибка при создании отчета: {ex.Message}");
        }
    }
}

