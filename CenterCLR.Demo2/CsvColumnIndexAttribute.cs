using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CenterCLR.Demo2
{
	[AttributeUsage(AttributeTargets.Property)]
	public sealed class CsvColumnIndexAttribute : Attribute
	{
		public CsvColumnIndexAttribute(int index)
		{
			this.Index = index;
		}

		public int Index
		{
			get;
			private set;
		}
	}
}
