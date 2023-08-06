using System;
using System.Text;

namespace BillingTool
{
	public class ResultMessage
	{
		public bool IsDone { get; set; }
		public int InsertStartRow { get; set; }
		public int InsertEndRow { get; set; }
		public int InsertCount { get; set; }
		public int StartReadRow { get; set; }
		public int EndReadRow { get; set; }
		public int ReadRowCount => EndReadRow - StartReadRow;
		public TimeSpan ConmudeTime { get; set; }

		public string Message { get; set; }

		public string ConmudeTimeStr { get => TimeCovert(); }

		private string TimeCovert()
		{
			StringBuilder sb = new StringBuilder();
			if (ConmudeTime.TotalMinutes < 1)
			{
				sb.Append($"{ConmudeTime.TotalSeconds}s");
			}
			else
			{
				sb.Append($"{ConmudeTime.TotalMinutes}m");
			}
			return sb.ToString();
		}
	}
}