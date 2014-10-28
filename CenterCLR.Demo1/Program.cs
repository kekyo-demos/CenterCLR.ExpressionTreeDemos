using System;
using System.Collections.Generic;
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
		public static string GetMessage(Type type, string fieldName)
		{
			var fieldInfo = type.GetField(fieldName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.DeclaredOnly);
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
			var title = MessageManager.GetMessage(typeof(Program), "TitleMessage");
			var description = MessageManager.GetMessage(typeof(Program), "DescriptionMessage");

			MessageBox.Show(description, title, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
		}
	}
}
