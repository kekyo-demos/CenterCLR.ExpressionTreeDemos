using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CenterCLR.Demo2
{
	public static class Utilities
	{
		public static object ConvertTo(string stringValue, Type targetType)
		{
			// boolなら
			if (targetType == typeof(bool))
			{
				// 数値からboolに変換
				return int.Parse(stringValue, CultureInfo.InvariantCulture) != 0;
			}
			// Enumなら
			else if (targetType.IsEnum == true)
			{
				// Enum.Parseを使用して変換
				return Enum.Parse(targetType, stringValue);
			}
			// それ以外
			else
			{
				// Convertクラスで変換
				return Convert.ChangeType(stringValue, targetType, CultureInfo.InvariantCulture);
			}
		}
	}
}
