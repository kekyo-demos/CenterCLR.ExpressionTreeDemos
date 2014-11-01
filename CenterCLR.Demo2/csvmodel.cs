namespace CenterCLR.Demo2
{
	public enum 更新Values
	{
		変更なし = 0,
		変更あり = 1,
		廃止 = 2,
	}
	public enum 変更理由Values
	{
		変更なし = 0,
		指令都市施行 = 1,
		住居表示 = 2,
		区画整理 = 3,
		郵便区調整 = 4,
		訂正 = 5,
		廃止 = 6,
	}
	public sealed class CsvModel
	{
		[CsvColumnIndex(0)]
		public int コード
		{
			get;
			set;
		}
		[CsvColumnIndex(1)]
		public string 旧郵便番号
		{
			get;
			set;
		}
		[CsvColumnIndex(2)]
		public string 郵便番号
		{
			get;
			set;
		}
		[CsvColumnIndex(3)]
		public string 都道府県カナ名
		{
			get;
			set;
		}
		[CsvColumnIndex(4)]
		public string 市区町村カナ名
		{
			get;
			set;
		}
		[CsvColumnIndex(5)]
		public string 町域カナ名
		{
			get;
			set;
		}
		[CsvColumnIndex(6)]
		public string 都道府県名
		{
			get;
			set;
		}
		[CsvColumnIndex(7)]
		public string 市区町村名
		{
			get;
			set;
		}
		[CsvColumnIndex(8)]
		public string 町域名
		{
			get;
			set;
		}
		[CsvColumnIndex(9)]
		public bool 重郵便番号
		{
			get;
			set;
		}
		[CsvColumnIndex(10)]
		public bool 小字別町域
		{
			get;
			set;
		}
		[CsvColumnIndex(11)]
		public bool 丁目
		{
			get;
			set;
		}
		[CsvColumnIndex(12)]
		public bool 重町域
		{
			get;
			set;
		}
		[CsvColumnIndex(13)]
		public 更新Values 更新
		{
			get;
			set;
		}
		[CsvColumnIndex(14)]
		public 変更理由Values 変更理由
		{
			get;
			set;
		}
	}
}
