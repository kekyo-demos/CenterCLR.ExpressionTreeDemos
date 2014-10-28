using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CenterCLR.Demo1
{
	public sealed class MessageManager
	{
		private readonly Dictionary<string, string> overrideMessages_;

		public MessageManager(string messagePath)
		{
			XLWorkbook workbook;
			using (var fs = new FileStream(messagePath, FileMode.Open, FileAccess.Read, FileShare.Read))
			{
				workbook = new XLWorkbook(fs);
			}

			overrideMessages_ =
				(from worksheet in workbook.Worksheets
				 from rowIndex in Enumerable.Range(1, worksheet.RowCount())
				 let key = worksheet.Cell(rowIndex, 1).Value.ToString().Trim()
				 where key.Length >= 1
				 select new
				 {
					 Key = key,
					 Message = worksheet.Cell(rowIndex, 2).Value.ToString().Trim()
				 }).
				ToDictionary(entry => entry.Key, entry => entry.Message);
		}

		public string GetMessage(Expression<Func<string>> field)
		{
			// 戻り値がstringとなる式である事は、タイプセーフ性で保証されている。
			// フィールドを示す式である事だけを確認する。

			var memberExpression = field.Body as MemberExpression;
			Debug.Assert(memberExpression != null);

			var fieldInfo = memberExpression.Member as FieldInfo;
			Debug.Assert(fieldInfo != null);
			Debug.Assert(fieldInfo.IsStatic == true);

			var key = string.Format("{0}.{1}.{2}", fieldInfo.DeclaringType.Namespace, fieldInfo.DeclaringType.Name, fieldInfo.Name);
			string value;
			if (overrideMessages_.TryGetValue(key, out value) == true)
			{
				return value;
			}
			else
			{
				return (string)fieldInfo.GetValue(null);
			}
		}

		public static void WriteMessages(string messagePath)
		{
			var dataTable = new DataTable("Message");
			dataTable.Columns.Add("キー", typeof(string));
			dataTable.Columns.Add("メッセージ", typeof(string));

			// 読み込まれている全てのクラスを調べて、フィールド定義を抽出する
			var fields =
				from assembly in AppDomain.CurrentDomain.GetAssemblies().AsParallel()
				from type in assembly.GetTypes()
				where
					(type.Namespace != null) &&
					(type.Namespace.StartsWith("CenterCLR.") == true) &&	// 簡易絞り込み
					(type.IsClass == true) &&
					(type.IsNested == false) &&
					(type.IsGenericType == false)
				from fieldInfo in type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.DeclaredOnly)
				where fieldInfo.FieldType == typeof(string)
				select fieldInfo;

			fields.ForEach(fieldInfo =>
				{
					var key = string.Format("{0}.{1}.{2}", fieldInfo.DeclaringType.Namespace, fieldInfo.DeclaringType.Name, fieldInfo.Name);
					var value = (string)fieldInfo.GetValue(null);
					dataTable.Rows.Add(key, value);
				});
			
			using (var fs = new FileStream(messagePath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
			{
				var workbook = new XLWorkbook();
				workbook.Worksheets.Add(dataTable);

				workbook.SaveAs(fs);
				fs.Close();
			}
		}
	}

	class Program
	{
		private static readonly string TitleMessage = "○×商事CRMシステム";
		private static readonly string DescriptionMessage = "宜しいですか？";

		static void Main(string[] args)
		{
			MessageManager.WriteMessages("SampleMessage.xlsx");

			var manager = new MessageManager("Message.xlsx");

			var title = manager.GetMessage(() => TitleMessage);
			var description = manager.GetMessage(() => DescriptionMessage);

			MessageBox.Show(description, title, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
		}
	}
}
