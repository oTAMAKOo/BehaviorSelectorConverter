
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BehaviorSelectorConverter
{
    public sealed class EditExcelBuilder
    {
        //----- params -----

        //----- field -----

        //----- property -----

        //----- method -----

        public static void Build(string indexFilePath, IndexData indexData, SheetData[] sheetData, Settings settings)
        {
            var originExcelPath = Path.GetFullPath(settings.ExcelPath);

            var editExcelPath = Path.ChangeExtension(indexFilePath, Constants.ExcelExtension);

            ConsoleUtility.Progress("------ Build edit excel file ------");

            //------ エディット用にエクセルファイルを複製 ------

            if (!File.Exists(originExcelPath))
            {
                throw new FileNotFoundException(string.Format("{0} is not exists.", originExcelPath));
            }

            var originXlsxFile = new FileInfo(originExcelPath);

            var editXlsxFile = originXlsxFile.CopyTo(editExcelPath, true);

            //------ レコード情報を書き込み ------

            if (sheetData == null) { return; }

            using (var excel = new ExcelPackage(editXlsxFile))
            {
                var worksheets = excel.Workbook.Worksheets;

                // テンプレートシート.

                var templateSheet = worksheets.FirstOrDefault(x => x.Name.ToLower() == settings.TemplateSheetName);

                if (templateSheet == null)
                {
                    throw new Exception(string.Format("Template worksheet {0} not found.", settings.TemplateSheetName));
                }

                // シート作成.

                foreach (var data in sheetData)
                {
                    if (string.IsNullOrEmpty(data.sheetName)) { continue; }

                    if (worksheets.Any(x => x.Name == data.sheetName))
                    {
                        throw new Exception(string.Format("Worksheet create failed. Worksheet {0} already exists", data.sheetName));
                    }

                    // テンプレートシートを複製.                    
                    var newWorksheet = worksheets.Add(data.sheetName, templateSheet);

                    // 保護解除.
                    newWorksheet.Protection.IsProtected = false;
                    // タブ選択状態解除.
                    newWorksheet.View.TabSelected = false;
                    // セルサイズ調整.
                    newWorksheet.Cells.AutoFitColumns();

                    // エラー無視.
                    var excelIgnoredError = newWorksheet.IgnoredErrors.Add(newWorksheet.Dimension);

                    excelIgnoredError.NumberStoredAsText = true;
                }

                // シート順番入れ替え.

                if (worksheets.Any() && indexData != null)
                {
                    for (var i = indexData.sheetNames.Length - 1; 0 <= i; i--)
                    {
                        var sheetName = indexData.sheetNames[i];

                        if (worksheets.All(x => x.Name != sheetName)) { continue; }

                        worksheets.MoveToStart(sheetName);
                    }
                }

                // 先頭のシートをアクティブ化.

                var firstWorksheet = worksheets.FirstOrDefault();

                if (firstWorksheet != null)
                {
                    firstWorksheet.View.TabSelected = true;
                }

                // レコード情報設定.

                foreach (var data in sheetData)
                {
                    var worksheet = worksheets.FirstOrDefault(x => x.Name == data.sheetName);

                    if (worksheet == null)
                    {
                        ConsoleUtility.Error("Worksheet:{0} not found.", data.sheetName);
                        continue;
                    }

                    var dimension = worksheet.Dimension;

                    var records = data.records;

                    if (records == null) { continue; }

                    worksheet.SetValue(Constants.FileNameAddress.Y, Constants.FileNameAddress.X, data.fileName);

                    // レコード投入用セルを用意.

                    for (var i = 0; i < records.Length; i++)
                    {
                        var recordRow = Constants.RecordStartRow + i;

                        // 行追加.
                        if (worksheet.Cells.End.Row < recordRow)
                        {
                            worksheet.InsertRow(recordRow, 1);
                        }

                        // セル情報コピー.
                        for (var column = 1; column < dimension.End.Column; column++)
                        {
                            CloneCellFormat(worksheet, Constants.RecordStartRow, recordRow, column);
                        }
                    }

                    // タイトル文字列の頭が「-」の場合は値設定をスキップ.

                    var titleValues = ExcelUtility.GetRowValues(worksheet, Constants.TitleRow).ToArray();

                    var ignoreIndexList = new List<int>();

                    for (var i = 0; i < titleValues.Length; i++)
                    {
                        var title = ExcelUtility.ConvertValue<string>(titleValues, i);

                        if (string.IsNullOrEmpty(title)) { continue; }

                        if (title.StartsWith("-"))
                        {
                            ignoreIndexList.Add(i + 1);
                        }
                    }

                    // 値設定.

                    for (var i = 0; i < records.Length; i++)
                    {
                        var row = Constants.RecordStartRow + i;

                        var column = Constants.DataStartColumn;

                        var record = records[i];

                        var successRate = (float)Math.Truncate(Convert.ToSingle(record.behavior.successRate) * 1000.0f) / 1000.0f;
                        
                        worksheet.SetValue(row, column, successRate.ToString());
                        column = GetNextColumn(column, ignoreIndexList);

                        worksheet.SetValue(row, column, record.behavior.actionType);
                        column = GetNextColumn(column, ignoreIndexList);

                        worksheet.SetValue(row, column, record.behavior.actionParameters);
                        column = GetNextColumn(column, ignoreIndexList);

                        worksheet.SetValue(row, column, record.behavior.targetType);
                        column = GetNextColumn(column, ignoreIndexList);

                        worksheet.SetValue(row, column, record.behavior.targetParameters);
                        column = GetNextColumn(column, ignoreIndexList);

                        for (var j = 0; j < record.behavior.conditions.Length; j++)
                        {
                            var condition = record.behavior.conditions[j];

                            if (j != 0)
                            {
                                worksheet.SetValue(row, column, condition.connecter);
                                column = GetNextColumn(column, ignoreIndexList);
                            }

                            worksheet.SetValue(row, column, condition.type);
                            column = GetNextColumn(column, ignoreIndexList);

                            worksheet.SetValue(row, column, condition.parameters);
                            column = GetNextColumn(column, ignoreIndexList);
                        }

                        // セル情報.
                        if (record.cells != null)
                        {
                            foreach (var cellData in record.cells)
                            {
                                var address = cellData.address.Split(',');

                                var rowStr = address.ElementAtOrDefault(0);
                                var columnStr = address.ElementAtOrDefault(1);

                                if (string.IsNullOrEmpty(rowStr) || string.IsNullOrEmpty(columnStr)) { continue; }

                                var r = Convert.ToInt32(rowStr);
                                var c = Convert.ToInt32(columnStr);

                                ExcelCellUtility.Set(worksheet, r, c, cellData);
                            }
                        }
                    }

                    // セルサイズを調整.
                    
                    var maxRow = Constants.RecordStartRow + records.Length + 1;

                    var celFitRange = worksheet.Cells[1, 1, maxRow, dimension.End.Column];

                    ExcelUtility.FitColumnSize(worksheet, celFitRange);

                    ExcelUtility.FitRowSize(worksheet, celFitRange);

                    // 説明.
                    if (!string.IsNullOrEmpty(data.description))
                    {
                        var r = Constants.DescriptionAddress.Y;
                        var c = Constants.DescriptionAddress.X;

                        var cell = worksheet.Cells[r, c];

                        cell.Style.WrapText = false;

                        worksheet.SetValue(r, c, data.description);
                    }

                    ConsoleUtility.Task("- {0}", data.sheetName);
                }

                // 保存.
                excel.Save();
            }
        }

        private static int GetNextColumn(int current, List<int> ignoreColumns)
        {
            var next = current + 1;

            while (true)
            {
                if (!ignoreColumns.Contains(next)){ break; }

                next++;
            }

            return next;
        }

        private static void CloneCellFormat(ExcelWorksheet worksheet, int recordStartRow, int row, int column)
        {
            var srcCell = worksheet.Cells[recordStartRow, column];
            var destCell = worksheet.Cells[row, column];

            srcCell.Copy(destCell);
        }
    }
}
