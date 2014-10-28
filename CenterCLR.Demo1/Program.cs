using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CenterCLR.Demo1
{
	public static class MessageManager
	{
		public static string GetMessage(string message)
		{
			return message;
		}
	}

	class Program
	{
		private static readonly string TitleMessage = "○×商事CRMシステム";
		private static readonly string DescriptionMessage = "宜しいですか？";

		static void Main(string[] args)
		{
			var title = MessageManager.GetMessage(TitleMessage);
			var description = MessageManager.GetMessage(DescriptionMessage);

			MessageBox.Show(description, title, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
		}
	}
}
