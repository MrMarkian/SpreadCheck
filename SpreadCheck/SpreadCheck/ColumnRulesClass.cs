using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpreadCheck
{
    class ColumnRulesClass
    {
    }

    public class ColumnRules
    {   public const int MaxAllowedValues = 20;
		public static int LastDetectedRow { get; set; }

        public ColumnRules()
        {  List<string> AllowedValuesArray = new List<string>();
			ColumnName = "";
			MustBeginWith = "";
			MustEndWith = "";
			MoreThanValue = 0;
			LessThanValue = 0;

        }

        static int ReturnMaxAllowedValue() { return MaxAllowedValues; }
        public static int ReturnLastDetectedRow() { return LastDetectedRow; }
        public static void SetLastDetectedRow(int count) { LastDetectedRow = count; }

        public bool IsEmpty { get; set; }
        public bool ContainsSpaces { get; set; }
        public bool ContainsNumber { get; set; }
        public bool ContainsNonAlpha { get; set; }
        public bool ContainsLetters { get; set; }

        public bool CheckDateTime { get; set; }
        public DateTime LowerDate { get; set; }
        public DateTime UpperDate { get; set; }

        public bool Enabled { get; set; }
        public bool AllowValuesEnabled { get; set; }
        public List<string> AllowedValuesArray;
        public string ColumnName { get; set; }

        public bool CheckMustBeginWith { get; set; }
        public string MustBeginWith { get; set; }
        public bool CheckMustEndWith {get;set;}
        public string MustEndWith { get; set; }

        public bool CheckMoreThan { get; set; }
        public bool CheckLessThan { get; set; }
        public double MoreThanValue { get; set; }
        public double LessThanValue { get; set; }

        public bool CheckLength { get; set; }
        public int Length { get; set; }
        public int HeaderColumnPosition { get; set; }
        public int HeaderRowPosition { get; set; }

        public bool TrimCell { get; set; }
        public bool ReverseData { get; set; }

        public int ChangeCaseType = 0;
        public bool ChangeCaseEnabled { get; set; }
		public bool ChangeTextAlignment { get; set; }
		public int TextAlignmentType = 0;
		
        public bool BackColorEnabled { get; set; }
    }
}
