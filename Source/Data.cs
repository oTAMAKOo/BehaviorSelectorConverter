
using System;
using Extensions;

namespace BehaviorSelectorConverter
{
    public sealed class IndexData
    {
        public string[] sheetNames = null;
    }

    public sealed class SheetData
    {
        public string fileName = null;

        public string description = null;

        public string sheetName = null;

        public RecordData[] records = null;
    }

    public sealed class RecordData
    {
        [Serializable]
        public sealed class Behavior
        {
            public float successRate = 0f;

            public string actionType = null;

            public string actionParameters = null;

            public string targetType = null;

            public string targetParameters = null;

            public Condition[] conditions = null;
        }

        [Serializable]
        public sealed class Condition
        {
            public string type = null;

            public string parameters = null;

            public string connecter = null;
        }

        public Behavior behavior = null;

        public ExcelCell[] cells = null;
    }

    public sealed class ExcelCell : Extensions.ExcelCell
    {
        public string address = null;
    }
}
