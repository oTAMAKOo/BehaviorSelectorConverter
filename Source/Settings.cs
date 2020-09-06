using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Extensions;

namespace BehaviorSelectorConverter
{
    public class Settings
    {
        //----- params -----

        public class FileSettings
        {
            /// <summary> フォーマット </summary>
            public string format = string.Empty;
        }

        private class ExcelSettings
        {
            /// <summary> パス </summary>
            public string origin = string.Empty;
            /// <summary> テンプレートシート名 </summary>
            public string templateSheetName = null;
            /// <summary> 除外シート名 </summary>
            public string ignoreSheetNames = null;
        }

        //----- field -----

        private FileSettings fileSettings = null;
        private ExcelSettings excelSettings = null;

        //----- property -----

        /// <summary> ファイルフォーマット </summary>
        public FileSystem.Format FileFormat
        {
            get
            {
                var format = fileSettings.format.ToLower();

                switch (format)
                {
                    case "yaml": return FileSystem.Format.Yaml;
                    case "json": return FileSystem.Format.Json;
                }

                return FileSystem.Format.Yaml;
            }
        }

        /// <summary> エクセルファイルパス </summary>
        public string ExcelPath { get { return excelSettings.origin; } }

        /// <summary> テンプレートシート名 </summary>
        public string TemplateSheetName { get { return excelSettings.templateSheetName; } }

        /// <summary> 除外シート名 </summary>
        public IReadOnlyList<string> IgnoreSheetNames { get; private set; }

        //----- method -----        

        public bool Load()
        {
            var iniFilePath = "./settings.ini";

            if (!File.Exists(iniFilePath)) { return false; }

            fileSettings = IniFile.Read<FileSettings>("File", iniFilePath);

            if (fileSettings == null) { return false; }

            excelSettings = IniFile.Read<ExcelSettings>("Excel", iniFilePath);

            if (excelSettings == null) { return false; }

            IgnoreSheetNames = excelSettings.ignoreSheetNames.Split(',').Select(x => x.Trim()).ToArray();

            return true;
        }

        public string GetFileExtension()
        {
            var extension = string.Empty;

            switch (FileFormat)
            {
                case FileSystem.Format.Json:
                    extension = Constants.JsonFileExtension;
                    break;
                case FileSystem.Format.Yaml:
                    extension = Constants.YamlFileExtension;
                    break;
            }

            return extension;
        }
    }
}
