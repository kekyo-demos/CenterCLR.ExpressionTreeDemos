using System.ComponentModel.Design.Serialization;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic.FileIO;

namespace CenterCLR.Demo2
{
	class Program
	{
		private static int ParseInt32(string stringValue, int defaultValue)
		{
			int value;
			if (int.TryParse(stringValue, out value) == true)
			{
				return value;
			}
			else
			{
				return defaultValue;
			}
		}

		private static IReadOnlyDictionary<string, CsvColumn> LoadCsvDefinitions(string csvDefinitionsPath) 
		{
			XLWorkbook workbook;
			using (var fs = new FileStream(csvDefinitionsPath, FileMode.Open, FileAccess.Read, FileShare.Read))
			{
				workbook = new XLWorkbook(fs);
			}

			return
				(from worksheet in workbook.Worksheets
					from rowIndex in Enumerable.Range(1, worksheet.RowCount()).AsParallel()
					let number = ParseInt32(worksheet.Cell(rowIndex, 1).Value.ToString().Trim(), -1)
					let fieldName = worksheet.Cell(rowIndex, 2).Value.ToString().Trim()
					let fieldDescription = worksheet.Cell(rowIndex, 3).Value.ToString().Trim()
					let fieldType = worksheet.Cell(rowIndex, 4).Value.ToString().Trim()
					where (number >= 0) && (fieldName.Length >= 1) && (fieldType.Length >= 1)
					select new CsvColumn(number, fieldName, fieldDescription, fieldType)).
				ToDictionary(entry => entry.FieldName, entry => entry);
		}

		static void Main(string[] args)
		{
			var csvDefinitions = LoadCsvDefinitions("csvdefinition.xlsx");

			using (var fs = new FileStream("sample.csv", FileMode.Open, FileAccess.Read, FileShare.Read))
			{
				using (var parser = new TextFieldParser(fs, Encoding.UTF8, true))
				{
					parser.Delimiters = new[] { "," };
					parser.HasFieldsEnclosedInQuotes = true;
					parser.TextFieldType = FieldType.Delimited;
					parser.TrimWhiteSpace = true;

					// レコードが存在する限り続ける
					while (parser.EndOfData == false)
					{
						// 一行読み取る
						string[] fields = parser.ReadFields();

						// TODO: なにか
					}
				}
			}
		}
	}
}
