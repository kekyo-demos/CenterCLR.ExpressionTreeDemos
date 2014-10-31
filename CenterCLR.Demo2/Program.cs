using System.ComponentModel;
using System.ComponentModel.Design.Serialization;
using System.Data.Common;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Charts;
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

				// コンストラクタを定義
				tw.WriteLine("		public {0}(string[] fields)", modelName);
				tw.WriteLine("		{");

				foreach (var column in csvDefinitions.Values.OrderBy(column => column.Number))
				{
					tw.WriteLine(
						"			this.{0} = ({2})Utilities.ConvertTo(fields[{1}], typeof({2}));",
						column.FieldName,
						column.Number,
						column.FieldClrType);
				}

				tw.WriteLine("		}");

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

		private static IEnumerable<T> CreateCsvContext<T>(string csvDefinitionPath, Encoding encoding)
			where T : new()
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

						var model = new T();

						// リフレクションで動的に値を代入する
						foreach (var entry in
							from pi in typeof(T).GetProperties()
							let csvColumnIndexAttribute = pi.GetCustomAttribute<CsvColumnIndexAttribute>()
							where csvColumnIndexAttribute != null
							select new
							{
								PropertyInfo = pi,
								Index = csvColumnIndexAttribute.Index
							})
						{
							// CsvColumnIndex属性で指定された位置の値文字列を得る
							var stringValue = fields[entry.Index];

							// 文字列からプロパティの型に変換する
							var value = Utilities.ConvertTo(stringValue, entry.PropertyInfo.PropertyType);

							// プロパティに設定（setterを呼び出す）
							entry.PropertyInfo.SetValue(model, value);
						}

						yield return model;
					}
				}
			}
		}

		private static IEnumerable<T> CreateCsvContext2<T>(string csvDefinitionPath, Encoding encoding)
		{
			// コンストラクタのパラメータを示す式木 (string[] fields)
			var parameters = new[] { Expression.Parameter(typeof(string[])) };

			// var model = (fields) => new T(fields);
			var root = Expression.Lambda<Func<string[], T>>(
				Expression.New(
				typeof(T).GetConstructor(new[] { typeof(string[]) }), parameters),
				parameters);

			// 式木をコンパイルして、実行可能なデリゲートを得る
			var creator = root.Compile();

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

						// インスタンスを生成
						var model = creator(fields);
						yield return model;
					}
				}
			}
		}

		static void Main(string[] args)
		{
			var csvDefinitions = LoadCsvDefinitions("csvdefinition.xlsx");
			WriteModelSource(@"..\..\csvmodel.cs", "CenterCLR.Demo2", "CsvModel", csvDefinitions);

			// コンテキストを生成
			var context = CreateCsvContext2<CsvModel>("ken_all.csv", Encoding.GetEncoding("Shift_JIS"));

			// 都道府県別にグルーピングして昇順にソートし、レコード数をカウントする
			var groupedCounts =
				from model in context
				group model by model.都道府県名 into g
				orderby g.Key
				select new
				{
					都道府県名 = g.Key,
					レコード数 = g.Count()
				};

			// 結果を出力する
			foreach (var groupedCount in groupedCounts)
			{
				Console.WriteLine("{0} = {1}", groupedCount.都道府県名, groupedCount.レコード数);
			}
		}
	}
}
