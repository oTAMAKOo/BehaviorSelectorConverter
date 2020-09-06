
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CommandLine;
using OfficeOpenXml;
using Extensions;

namespace BehaviorSelectorConverter
{
    class Program
    {
        private class CommandLineOptions
        {
            [Option("workspace", Required = false)]
            public string Workspace { get; set; } = "";
            [Option("mode", Required = false, HelpText = "Convert mode. (import or export)")]
            public string Mode { get; set; } = "import";
        }

        [STAThread]
        static void Main(string[] args)
        {
            // コマンドライン引数.

            var options = Parser.Default.ParseArguments<CommandLineOptions>(args) as Parsed<CommandLineOptions>;

            if (options == null)
            {
                Exit(1, "Arguments parse failed.");
            }

            // 設定ファイル.

            var settings = new Settings();

            if (!settings.Load())
            {
                Exit(1, "Settings load failed.");
            }
            
            var workspace = options.Value.Workspace;
            var mode = options.Value.Mode;

            //=== 開発用 ========================================

            #if DEBUG

            workspace = @"G:\project\BehaviorSelectorConverter\bin\Debug";

            Directory.SetCurrentDirectory(workspace);

            mode = "export";

            #endif

            //==================================================*/

            if (string.IsNullOrEmpty(workspace))
            {
                Exit(0);
            }

            // EPPlus License setup.
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 実行.

            switch (mode)
            {
                case "import":
                    Import(workspace, settings);
                    break;

                case "export":
                    Export(workspace, settings);
                    break;
            }

            Exit(0);
        }

        private static string OpenSelectFileDialog(string workspace, string title, string filter)
        {
            var filePath = string.Empty;

            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = title;
                openFileDialog.InitialDirectory = workspace;
                openFileDialog.Filter = filter;
                openFileDialog.FilterIndex = 2;
                openFileDialog.Multiselect = false;
                openFileDialog.CheckFileExists = true;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                }
            }

            return filePath;
        }

        private static void Import(string workspace, Settings settings)
        {
            var title = "Select Index File";
            var filter = string.Format("Index File (*{0})|*{0}", Constants.IndexFileExtension);

            var indexFilePath = OpenSelectFileDialog(workspace, title, filter);

            if (string.IsNullOrEmpty(indexFilePath))
            {
                Exit(0);

                return;
            }

            if (IsEditExcelFileLocked(indexFilePath, settings)) { return; }

            var indexData = DataLoader.LoadSheetIndex(indexFilePath, settings);

            var directory = Path.GetDirectoryName(indexFilePath);

            var fileName = Path.GetFileNameWithoutExtension(indexFilePath);

            var folderPath = PathUtility.Combine(directory, fileName);

            var sheetData = DataLoader.LoadAllSheetData(folderPath, settings);

            EditExcelBuilder.Build(indexFilePath, indexData, sheetData, settings);
        }

        private static void Export(string workspace, Settings settings)
        {
            var title = "Select Excel File";
            var filter = string.Format("Excel file (*{0})|*{0}", Constants.ExcelExtension);

            var excelFilePath = OpenSelectFileDialog(workspace, title, filter);

            if (string.IsNullOrEmpty(excelFilePath))
            {
                Exit(0);

                return;
            }

            var builder = new StringBuilder();
            
            var sheetData = ExcelDataLoader.LoadExcelData(excelFilePath, settings);

            var emptyFileNameSheetData = sheetData.Where(x => string.IsNullOrEmpty(x.fileName)).ToArray();

            if (emptyFileNameSheetData.Any())
            {
                builder.AppendLine();

                foreach (var item in emptyFileNameSheetData)
                {
                    builder.AppendFormat("Empty fileName sheet exists. SheetName = {0}", item.sheetName).AppendLine();
                }
            }

            var duplicates = sheetData.GroupBy(x => x.fileName).Where(x => 1 < x.Count()).ToArray();

            if (duplicates.Any())
            {
                builder.AppendLine();

                foreach (var group in duplicates)
                {
                    foreach (var item in group)
                    {
                        builder.AppendFormat("Duplicate file name exists. Sheet = {0} FileName = {1}", item.sheetName, item.fileName).AppendLine();
                    }
                }
            }

            var errorMessage = builder.ToString();

            if (!string.IsNullOrEmpty(errorMessage))
            {
                Exit(1, errorMessage);
                return;
            }

            DataWriter.WriteAllSheetData(excelFilePath, sheetData, settings);

            var sheetNames = ExcelDataLoader.LoadSheetNames(excelFilePath, settings);

            DataWriter.WriteSheetIndex(excelFilePath, sheetNames, settings);
        }

        private static bool IsEditExcelFileLocked(string indexFilePath, Settings settings)
        {
            var excelFilePath = Path.ChangeExtension(indexFilePath, Constants.ExcelExtension);
            
            // ファイルが存在＋ロック時はエラー.
            if (File.Exists(excelFilePath))
            {
                if (FileUtility.IsFileLocked(excelFilePath))
                {
                    Exit(1, string.Format("File locked. {0}", excelFilePath));
                    return true;
                }
            }

            return false;
        }

        private static void Exit(int exitCode, string message = "")
        {
            if (!string.IsNullOrEmpty(message))
            {
                ConsoleUtility.Error(message);
            }

            // 正常終了以外ならコンソールを閉じない.
            if (exitCode != 0)
            {
                Console.ReadLine();
            }

            Environment.Exit(exitCode);
        }
    }
}
