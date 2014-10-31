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
		public int コード
		{
			get;
			set;
		}
		public string 旧郵便番号
		{
			get;
			set;
		}
		public string 郵便番号
		{
			get;
			set;
		}
		public string 都道府県カナ名
		{
			get;
			set;
		}
		public string 市区町村カナ名
		{
			get;
			set;
		}
		public string 町域カナ名
		{
			get;
			set;
		}
		public string 都道府県名
		{
			get;
			set;
		}
		public string 市区町村名
		{
			get;
			set;
		}
		public string 町域名
		{
			get;
			set;
		}
		public bool 重郵便番号
		{
			get;
			set;
		}
		public bool 小字別町域
		{
			get;
			set;
		}
		public bool 丁目
		{
			get;
			set;
		}
		public bool 重町域
		{
			get;
			set;
		}
		public 更新Values 更新
		{
			get;
			set;
		}
		public 変更理由Values 変更理由
		{
			get;
			set;
		}
	}
}
