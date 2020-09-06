using System;
using System.IO;
using Extensions;

namespace BehaviorSelectorConverter
{
    public sealed class DataWriter
    {
        //----- params -----

        //----- field -----

        //----- property -----

        //----- method -----

        public static void WriteSheetIndex(string excelFilePath, string[] sheetNames, Settings settings)
        {
            var filePath = Path.ChangeExtension(excelFilePath, Constants.IndexFileExtension);

            var indexData = new IndexData()
            {
                sheetNames = sheetNames
            };

            FileSystem.WriteFile(filePath, indexData, settings.FileFormat);
        }

        public static void WriteAllSheetData(string excelFilePath, SheetData[] sheetData, Settings settings)
        {
            var excelFileName = Path.GetFileNameWithoutExtension(excelFilePath);

            var directory = Path.GetDirectoryName(excelFilePath);

            var dataFileDirectory = PathUtility.Combine(directory, excelFileName);
            
            DirectoryUtility.Clean(dataFileDirectory);
            
            var extension = settings.GetFileExtension();

            if (sheetData.IsEmpty()){ return; }

            ConsoleUtility.Progress("------ WriteData ------");

            foreach (var data in sheetData)
            {
                if (string.IsNullOrEmpty(data.sheetName)) { continue; }

                var records = data.records;

                if (records == null || records.IsEmpty()) { continue; }

                // シート情報書き出し.

                if (!string.IsNullOrEmpty(data.fileName))
                {
                    var fileName = data.fileName + extension;

                    var filePath = PathUtility.Combine(dataFileDirectory, fileName);

                    FileSystem.WriteFile(filePath, data, settings.FileFormat);
                }

                ConsoleUtility.Task("- {0}", data.sheetName);
            }
        }
    }
}
