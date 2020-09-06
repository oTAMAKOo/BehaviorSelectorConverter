
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Extensions;
using OfficeOpenXml;

namespace BehaviorSelectorConverter
{
    public static class ExcelDataLoader
    {
        /// <summary> シート名一覧読み込み(.xlsx) </summary>
        public static string[] LoadSheetNames(string excelFilePath, Settings settings)
        {
            if (!File.Exists(excelFilePath)) { return null; }

            var sheetNames = new List<string>();

            using (var excel = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                foreach (var worksheet in excel.Workbook.Worksheets)
                {
                    sheetNames.Add(worksheet.Name);
                }
            }

            return sheetNames.ToArray();
        }

        /// <summary> レコード情報読み込み(.xlsx) </summary>
        public static SheetData[] LoadExcelData(string excelFilePath, Settings settings)
        {
            if (!File.Exists(excelFilePath)) { return null; }

            ConsoleUtility.Progress("------ LoadExcelData ------");

            var sheets = new List<SheetData>();

            using (var excel = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                foreach (var worksheet in excel.Workbook.Worksheets)
                {
                    if (worksheet.Name == settings.TemplateSheetName) { continue; }

                    if (settings.IgnoreSheetNames.Contains(worksheet.Name)) { continue; }

                    var fileNameValue = worksheet.GetValue(Constants.FileNameAddress.Y, Constants.FileNameAddress.X);

                    var fileName = ExcelUtility.ConvertValue<string>(fileNameValue);

                    var descriptionValue = worksheet.GetValue(Constants.DescriptionAddress.Y, Constants.DescriptionAddress.X);

                    var description = ExcelUtility.ConvertValue<string>(descriptionValue);

                    var sheetData = new SheetData()
                    {
                        fileName = fileName,
                        description = description,
                        sheetName = worksheet.Name,
                    };

                    // タイトル文字列の頭が「-」の場合はデータに含めない.

                    var titleValues = ExcelUtility.GetRowValues(worksheet, Constants.TitleRow).ToArray();

                    var ignoreColumnList = new List<int>();

                    for (var i = 0; i < titleValues.Length; i++)
                    {
                        var title = ExcelUtility.ConvertValue<string>(titleValues, i);

                        if (string.IsNullOrEmpty(title)){ continue; }

                        if (title.StartsWith("-"))
                        {
                            ignoreColumnList.Add(i + 1);
                        }
                    }

                    var records = new List<RecordData>();

                    for (var r = Constants.RecordStartRow; r <= worksheet.Dimension.End.Row; r++)
                    {
                        var rowValues = ExcelUtility.GetRowValues(worksheet, r).ToArray();

                        // 開始.

                        var column = Constants.DataStartColumn - 1;

                        var behaviorData = new RecordData.Behavior();

                        if(rowValues.All(x => x == null)){ continue; }

                        // 行動データ.

                        var successRateValue = ExcelUtility.ConvertValue<float>(rowValues, column).ToString("F4");
                        behaviorData.successRate = (float)Math.Truncate(Convert.ToSingle(successRateValue) * 1000.0f) / 1000.0f;
                        column = GetNextColumn(column, ignoreColumnList);

                        behaviorData.actionType = ExcelUtility.ConvertValue<string>(rowValues, column);
                        column = GetNextColumn(column, ignoreColumnList);

                        behaviorData.actionParameters = ExcelUtility.ConvertValue<string>(rowValues, column);
                        column = GetNextColumn(column, ignoreColumnList);

                        behaviorData.targetType = ExcelUtility.ConvertValue<string>(rowValues, column);
                        column = GetNextColumn(column, ignoreColumnList);

                        behaviorData.targetParameters = ExcelUtility.ConvertValue<string>(rowValues, column);
                        column = GetNextColumn(column, ignoreColumnList);

                        var conditions = new List<RecordData.Condition>();

                        while (true)
                        {
                            var cellValue = ExcelUtility.ConvertValue<string>(rowValues, column);

                            if (cellValue == null){ break; }

                            var condition = new RecordData.Condition();

                            try
                            {
                                if (conditions.Any())
                                {
                                    var connecter = ExcelUtility.ConvertValue<string>(rowValues, column);
                                    column = GetNextColumn(column, ignoreColumnList);

                                    if (string.IsNullOrEmpty(connecter)){ break; }

                                    connecter = connecter.Trim();

                                    if (connecter != "|" && connecter != "&")
                                    {
                                        throw new InvalidDataException("connecter support | or &.");
                                    }

                                    condition.connecter = connecter;
                                }

                                condition.type = ExcelUtility.ConvertValue<string>(rowValues, column);
                                column = GetNextColumn(column, ignoreColumnList);

                                condition.parameters = ExcelUtility.ConvertValue<string>(rowValues, column);
                                column = GetNextColumn(column, ignoreColumnList);
                            }
                            catch (Exception e)
                            {
                                var errorMessage = string.Format("[{0},{1}] condition data error.\n{2}", r, column, e.Message);

                                ConsoleUtility.Error(errorMessage);

                                break;
                            }

                            conditions.Add(condition);
                        }

                        behaviorData.conditions = conditions.ToArray();

                        // セル情報.

                        var cells = new List<ExcelCell>();

                        for (var c = Constants.DataStartColumn; c < column; c++)
                        {
                            var cellData = ExcelCellUtility.Get<ExcelCell>(worksheet, r, c);

                            if (cellData == null){ continue; }

                            cellData.address = string.Format("{0},{1}", r, c);

                            cells.Add(cellData);
                        }

                        var recordData = new RecordData()
                        {
                            behavior = behaviorData,
                            cells = cells.Any() ? cells.ToArray() : null,
                        };

                        records.Add(recordData);
                    }

                    sheetData.records = records.ToArray();

                    sheets.Add(sheetData);

                    ConsoleUtility.Task("- {0}", sheetData.sheetName);
                }
            }

            return sheets.ToArray();
        }

        private static int GetNextColumn(int current, List<int> ignoreColumns)
        {
            var next = current + 1;

            while (true)
            {
                if (!ignoreColumns.Contains(next)) { break; }

                next++;
            }

            return next;
        }
    }
}
