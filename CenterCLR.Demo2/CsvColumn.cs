using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace CenterCLR.Demo2
{
	internal sealed class CsvColumn
	{
		public readonly int Number;
		public readonly string FieldName;
		public readonly string FieldDescription;
		public readonly string FieldType;

		public CsvColumn(int number, string fieldName, string fieldDescription, string fieldType)
		{
			this.Number = number;
			this.FieldName = fieldName;
			this.FieldDescription = fieldDescription;
			this.FieldType = fieldType;
		}

		public string FieldClrType
		{
			get 
			{
				switch (this.FieldType)
				{
					case "32ビット整数":
						return "int";
					case "文字列":
						return "string";
					case "真偽値":
						return "bool";
					case "8ビット整数":
						return this.FieldName + "Values";
					default:
						return "string";
				}
			}
		}

		public IReadOnlyDictionary<string, int> FieldClrEnumDefinition
		{
			get
			{
				if (this.FieldType != "8ビット整数")
				{
					return null;
				}

				return this.FieldDescription.
					Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).
					Select(word => word.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries)).
					Where(kv => kv.Length == 2).
					ToDictionary(kv => kv[1], kv => int.Parse(kv[0]));
			}
		}
	}
}
