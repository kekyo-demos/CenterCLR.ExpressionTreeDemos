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
			where T : new()
		{
			// コンストラクタのパラメータを示す式木 (string[] fields)
			var fieldsExpression = Expression.Parameter(typeof(string[]));
			var parameters = new[] { fieldsExpression };

			// 生成したいのは、こういうラムダ式に相当する式木（但しクラスはT型）
			//Expression<Func<string[], CsvModel>> creatorExpression = fields => new CsvModel
			//{
			//	コード = int.Parse(fields[0], CultureInfo.InvariantCulture),
			//	旧郵便番号 = fields[1],
			//	郵便番号 = fields[2],
			//	// ...
			//};

			// T型のインスタンスを生成する式木を生成
			// new T()
			var newExpression = Expression.New(
				typeof(T).GetConstructor(Type.EmptyTypes));

			// T型のプロパティ群を初期化するための代入式の式木群を生成
			// T {
			//	コード = int.Parse(fields[0], CultureInfo.InvariantCulture),
			//	旧郵便番号 = fields[1],
			//	郵便番号 = fields[2],
			//	// ...
			// }
			var memberAssignmentExpressions = typeof(T).GetProperties().
				Select(pi =>
				{
					// 属性からフィールドのインデックス値を入手
					var index = pi.GetCustomAttribute<CsvColumnIndexAttribute>().Index;

					// インデックス値を示す式木を生成
					var indexExpression = Expression.Constant(index);

					// インデックスを指定して配列にアクセスする式木を生成
					// fields[index]
					var arrayAccessExpression = Expression.ArrayAccess(
						fieldsExpression,
						indexExpression);

					// ConvertToに相当する式木の処理：式木内で変換操作を行う必要がある。
					Expression convertToExpression;

					// 真偽値なら
					if (pi.PropertyType == typeof(bool))
					{
						// int.Parseでintに変換する
						// int.Parse(fields[index])
						var callExpression = Expression.Call(
							null,
							typeof(int).GetMethod("Parse", new[] { typeof(string) }),
							arrayAccessExpression);

						// 数値が0かどうかを判定する式で、真偽値に変換する
						// int.Parse(fields[index]) != 0
						convertToExpression = Expression.NotEqual(
							callExpression,
							Expression.Constant(0));
					}
					// Enumなら
					else if (pi.PropertyType.IsEnum == true)
					{
						// Enum.Parseメソッドを使う（P型はプロパティの型）
						// Enum.Parse(typeof(P), fields[index])
						var callExpression = Expression.Call(
							null,
							typeof(Enum).GetMethod("Parse", new[] { typeof(Type), typeof(string) }),
							Expression.Constant(pi.PropertyType),
							arrayAccessExpression);

						// 結果をキャストする
						// (P)Convert.ChangeType(fields[index], typeof(P))
						convertToExpression = Expression.Convert(
							callExpression,
							pi.PropertyType);
					}
					// 文字列なら
					else if (pi.PropertyType == typeof(string))
					{
						// fieldsは元々文字列なので、変換操作は不要
						convertToExpression = arrayAccessExpression;
					}
					// その他
					else
					{
						// Convert.ChangeTypeメソッドを使う（P型はプロパティの型）
						// Convert.ChangeType(fields[index], typeof(P))
						var callExpression = Expression.Call(
							null,
							typeof(Convert).GetMethod("ChangeType", new[] { typeof(object), typeof(Type) }),
							arrayAccessExpression,
							Expression.Constant(pi.PropertyType));

						// 結果をキャストする
						// (P)Convert.ChangeType(fields[index], typeof(P))
						convertToExpression = Expression.Convert(
							callExpression,
							pi.PropertyType);
					}

					// 変換の結果をプロパティに代入する式木を生成
					return Expression.Bind(
						pi,
						convertToExpression);
				});

			// T型のインスタンスを生成し、fieldsの値をプロパティに群に代入する式木を生成
			// new T { ... }
			var memberInitExpression = Expression.MemberInit(
					newExpression,
					memberAssignmentExpressions);

			// T型のインスタンスを生成して初期化するラムダ式を示す式木を生成
			// fields => new T { ... }
			var lambdaExpression = Expression.Lambda<Func<string[], T>>(
				memberInitExpression,
				parameters);

			// 式木をコンパイルして、実行可能なデリゲートを得る
			var creator = lambdaExpression.Compile();

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
