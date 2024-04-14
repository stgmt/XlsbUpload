using System;
using System.Linq;
using XlsbUpload.services;

namespace XlsbUpload
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var reportBuilder = new DepartmentTaskReportService(args);

            var result = reportBuilder.MakeReport().ToList();

            foreach (var report in result)
            {
                Console.WriteLine(report ? $"Успешная выгрузка. очтета : {result.IndexOf(report)}." : "Неведомая ошибка");
            }

            Console.ReadLine();
        }
    }
}
