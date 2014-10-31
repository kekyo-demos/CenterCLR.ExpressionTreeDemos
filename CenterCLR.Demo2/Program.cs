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

		private static void WriteModelSource(
			string modelSourcePath,
			string modelNamespace,
			string modelName,
			IReadOnlyDictionary<string, CsvColumn> csvDefinitions)
		{
			using (var fs = new FileStream(modelSourcePath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
			{
				var tw = new StreamWriter(fs, Encoding.UTF8);

				tw.WriteLine("namespace {0}", modelNamespace);
				tw.WriteLine("{");

				// Enumを出力
				foreach (var entry in
					from column in csvDefinitions.Values
					let definition = column.FieldClrEnumDefinition
					where definition != null
					select new
					{
						Type = column.FieldClrType,
						Definition = definition
					})
				{
					tw.WriteLine("	public enum {0}", entry.Type);
					tw.WriteLine("	{");

					foreach (var definition in entry.Definition.OrderBy(definition => definition.Value))
					{
						tw.WriteLine("		{0} = {1},", definition.Key, definition.Value);
					}

					tw.WriteLine("	}");
				}

				tw.WriteLine("	public sealed class {0}", modelName);
				tw.WriteLine("	{");

				foreach (var column in csvDefinitions.Values.OrderBy(column => column.Number))
				{
					tw.WriteLine("		[CsvColumnIndex({0})]", column.Number);
					tw.WriteLine("		public {0} {1}", column.FieldClrType, column.FieldName);
					tw.WriteLine("		{");
					tw.WriteLine("			get;");
					tw.WriteLine("			set;");
					tw.WriteLine("		}");
				}

				tw.WriteLine("	}");
				tw.WriteLine("}");

				tw.Flush();
				fs.Close();
			}
		}

		private static IEnumerable<string[]> CreateCsvContext(string csvDefinitionPath, Encoding encoding)
		{
			using (var fs = new FileStream(csvDefinitionPath, FileMode.Open, FileAccess.Read, FileShare.Read))
			{
				using (var parser = new TextFieldParser(fs, encoding, true))
				{
					parser.Delimiters = new[] { "," };
					parser.HasFieldsEnclosedInQuotes = true;
					parser.TextFieldType = FieldType.Delimited;
					parser.TrimWhiteSpace = true;

					// レコードが存在する限り続ける
					while (parser.EndOfData == false)
					{
						// 一行読み取る
						var fields = parser.ReadFields();

						yield return fields;
					}
				}
			}
		}

		static void Main(string[] args)
		{
			var csvDefinitions = LoadCsvDefinitions("csvdefinition.xlsx");
			WriteModelSource(@"..\..\csvmodel.cs", "CenterCLR.Demo2", "CsvModel", csvDefinitions);

			foreach (var fields in CreateCsvContext("ken_all.csv", Encoding.GetEncoding("Shift_JIS")))
			{
				// モデルクラスに変換する
				var model = new CsvModel
				{
					コード = ParseInt32(fields[0], -1),
					旧郵便番号 = fields[1],
					郵便番号 = fields[2],

					// ...
				};
			}
		}
	}
}
