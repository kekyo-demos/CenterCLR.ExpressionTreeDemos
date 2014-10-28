using ClosedXML.Excel;
using System;
using System.Collections.Generic;
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
	}

	class Program
	{
		private static readonly string TitleMessage = "○×商事CRMシステム";
		private static readonly string DescriptionMessage = "宜しいですか？";

		static void Main(string[] args)
		{
			var manager = new MessageManager("Message.xlsx");

			var title = manager.GetMessage(() => TitleMessage);
			var description = manager.GetMessage(() => DescriptionMessage);

			MessageBox.Show(description, title, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
		}
	}
}
