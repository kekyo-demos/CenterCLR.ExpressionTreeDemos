using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CenterCLR.Demo1
{
	public static class MessageManager
	{
		public static string GetMessage(Expression<Func<string>> field)
		{
			// 戻り値がstringとなる式である事は、タイプセーフ性で保証されている。
			// フィールドを示す式である事だけを確認する。

			var memberExpression = field.Body as MemberExpression;
			Debug.Assert(memberExpression != null);

			var fieldInfo = memberExpression.Member as FieldInfo;
			Debug.Assert(fieldInfo != null);
			Debug.Assert(fieldInfo.IsStatic == true);

			var key = string.Format("{0}.{1}.{2}", fieldInfo.DeclaringType.Namespace, fieldInfo.DeclaringType.Name, fieldInfo.Name);
				
			var value = (string)fieldInfo.GetValue(null);

			// TODO:キーを使って何かする

			return value;
		}
	}

	class Program
	{
		private static readonly string TitleMessage = "○×商事CRMシステム";
		private static readonly string DescriptionMessage = "宜しいですか？";

		static void Main(string[] args)
		{
			var title = MessageManager.GetMessage(() => TitleMessage);
			var description = MessageManager.GetMessage(() => DescriptionMessage);

			MessageBox.Show(description, title, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
		}
	}
}
