
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Extensions;

namespace BehaviorSelectorConverter
{
    public sealed class DataLoader
    {
        //----- params -----

        //----- field -----

        //----- property -----

        //----- method -----

        public static IndexData LoadSheetIndex(string indexFilePath, Settings settings)
        {
            return FileSystem.LoadFile<IndexData>(indexFilePath, settings.FileFormat);
        }

        /// <summary> エクセル情報読み込み </summary>
        public static SheetData[] LoadAllSheetData(string folderPath, Settings settings)
        {
            // シート情報読み込み.

            var extension = settings.GetFileExtension();

            var dataFiles = Directory.EnumerateFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly)
                .Where(x => Path.GetExtension(x) == extension)
                .ToArray();

            var sheets = new List<SheetData>();

            if (dataFiles.IsEmpty()){ return new SheetData[0]; }

            ConsoleUtility.Progress("------ LoadSheetData ------");

            foreach (var sheetFile in dataFiles)
            {
                var sheet = LoadSheetData(sheetFile, settings);

                if (sheet != null)
                {
                    ConsoleUtility.Task("- {0}", sheet.sheetName);

                    sheets.Add(sheet);
                }
            }

            return sheets.ToArray();
        }

        private static SheetData LoadSheetData(string filePath, Settings settings)
        {
            return FileSystem.LoadFile<SheetData>(filePath, settings.FileFormat);
        }
    }
}
