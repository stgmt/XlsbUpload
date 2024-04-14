using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace XlsbUpload.services
{

    internal class ReaderBase
    {
        internal IEnumerable<string> GetDocumentPath(string[] filesPath)
        {
            // Получаем список всех файлов с расширением ".xlsb" в текущей директории
            string[] xlsbFilesFromCurrentDir = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xlsb");

            var xlsbFiles = filesPath.Where(filePath => filePath.EndsWith(".xlsb")).Concat(xlsbFilesFromCurrentDir);

            if (!xlsbFiles.Any())
            {
                throw new ArgumentException("не найден файл .xlsb");
            }

            foreach (var xlsbFile in xlsbFiles)
            {
                // Проверяем, существует ли файл по указанному пути
                if (!File.Exists(xlsbFile))
                {
                    throw new FileNotFoundException("файл не существует");
                }
                yield return xlsbFile;
            }
        }

    }
}
