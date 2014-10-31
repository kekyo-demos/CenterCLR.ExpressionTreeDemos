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
	}
}
