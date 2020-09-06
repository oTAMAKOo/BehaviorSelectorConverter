
using System;
using System.Drawing;

namespace BehaviorSelectorConverter
{
    public static class Constants
    {
        /// <summary>インデックスファイル拡張子 </summary>
        public const string IndexFileExtension = ".index";
        
        /// <summary> Json拡張子 </summary>
        public const string JsonFileExtension = ".json";

        /// <summary> Yaml拡張子 </summary>
        public const string YamlFileExtension = ".yaml";

        /// <summary> Excel拡張子 </summary>
        public const string ExcelExtension = ".xlsx";

        /// <summary> ファイル名名定義アドレス </summary>
        public static readonly Point FileNameAddress = new Point(2, 1);

        /// <summary> シート説明定義アドレス </summary>
        public static readonly Point DescriptionAddress = new Point(2, 2);

        /// <summary> データ名定義行 </summary>
        public const int TitleRow = 3;

        /// <summary> データ開始行 </summary>
        public const int RecordStartRow = 4;

        /// <summary> データ開始列 </summary>
        public const int DataStartColumn = 1;
    }
}
