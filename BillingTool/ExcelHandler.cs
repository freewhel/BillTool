using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace BillingTool
{
	public class ESourceType
	{
		public const string WECHAT = "微信";
		public const string ALIPAY = "alipay";
	}

	public class ExcelHandler
	{
		private readonly int[] weChatIndexs = { 1, 6, 5, 3, 4, 8 };
		private readonly int[] aliPayIndexs = { 5, 10, 11, 8, 9, 16 };
		private readonly int[] targetIndexs = { 1, 2, 3, 4, 5, 6, 7, 8 };
		private readonly string[] targetHead = { "付款时间", "金额（元）", "收/支", "交易对方", "收支类型", "商品名称", "交易状态", "支付方式" };
		private Dictionary<string, int[]> sourceTypes;
		private Dictionary<string, List<string>> ruleDict = new Dictionary<string, List<string>>();
		private string sourceType = string.Empty;

		private ResultMessage resultMessage;

		/// <summary>
		/// 源数据表
		/// </summary>
		private ExcelPackage sourceExcel;

		/// <summary>
		/// 目标表
		/// </summary>
		private ExcelPackage targetExcel;

		private ExcelWorksheet sourceSheet1;
		private ExcelWorksheet targetSheet1;
		private readonly string sourcePath;
		private readonly string targetPath;
		private string ruleFilePath = $"{Environment.CurrentDirectory}\\rule.ini";


		public ExcelHandler(string sourcePath, string targetPath)
		{
			this.sourcePath = sourcePath;
			this.targetPath = targetPath;
			Init();
		}

		private void Init()
		{
			//申明非商用
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			//weChatIndexs = new int[] { 1, 6, 5, 3, 4, 8 };
			//aliPayIndexs = new int[] { 5, 10, 11, 8, 9, 16 };
			//targetIndexs = new int[] { 1, 2, 3, 4, 5, 6, 7, 8 };
			//targetHead = new string[] { "付款时间", "金额（元）", "收/支", "交易对方", "收支类型", "商品名称", "交易状态", "支付方式" };
			resultMessage = new ResultMessage();
			sourceTypes = new Dictionary<string, int[]>
			{
				{ ESourceType.ALIPAY, aliPayIndexs },
				{ ESourceType.WECHAT, weChatIndexs }
			};
			InitRuleFile();
		}

		private void InitRuleFile()
		{
			try
			{
				if (!File.Exists(ruleFilePath))
				{
					File.WriteAllText(ruleFilePath, "");
				}
				IniFileHandler iniFileManager = new IniFileHandler(ruleFilePath);
				Dictionary<string, List<string>> temp = iniFileManager.GetSectionsKeys();

				if (temp.Count > 0)
				{
					ruleDict = temp;
				}
			}
			catch (Exception)
			{


			}

		}

		private bool LoadBills()
		{
			bool result = false;
			try
			{

				//确定数据文件类型
				if (sourcePath.Contains(ESourceType.ALIPAY))
				{
					sourceType = ESourceType.ALIPAY;
				}
				else if (sourcePath.Contains(ESourceType.WECHAT))
				{
					sourceType = ESourceType.WECHAT;
				}
				else
				{
					resultMessage.IsDone = false;
					resultMessage.Message = "数据文件类型检测错误，请检查数据文件内容是否正确！";
					return result;
				}


				if (!sourcePath.EndsWith("xlsx"))
				{
					sourceExcel = LoadCsv(sourcePath, sourceType);
					sourceSheet1 = sourceExcel.Workbook.Worksheets[0];
				}
				else
				{
					sourceExcel = new ExcelPackage(new FileInfo(sourcePath));
					sourceSheet1 = sourceExcel.Workbook.Worksheets[0];
				}
				targetExcel = new ExcelPackage(new FileInfo(targetPath));
				targetSheet1 = targetExcel.Workbook.Worksheets[0];
				result = true;
			}
			catch (Exception e)
			{
				resultMessage.IsDone = false;
				resultMessage.Message = e.Message;
				result = false;
			}
			return result;
		}


		private ExcelPackage LoadCsv(string filePath, string type)
		{

			string fileContext = File.ReadAllText(filePath, Encoding.GetEncoding("GB2312"));

			ExcelTextFormat format = new ExcelTextFormat();
			switch (type)
			{
				case ESourceType.ALIPAY:

					break;
				case ESourceType.WECHAT:
					fileContext = fileContext.Replace("¥", "");
					format.TextQualifier = '\"';
					break;
				default:
					break;
			}


			ExcelPackage excelPackage = new ExcelPackage();

			ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

			worksheet.Cells[1, 1].LoadFromText(fileContext, format);
			return excelPackage;
		}


		public ResultMessage SoucerToTargetBill()
		{

			resultMessage.IsDone = true;
			DateTime startTime = DateTime.Now;

			if (!LoadBills())
			{
				return resultMessage;
			}

			for (int i = 1; i <= targetSheet1.Dimension.End.Column; i++)
			{
				bool isTrue = targetSheet1.Cells[1, i].Value.ToString().Equals(targetHead[i - 1]);
				if (!isTrue)
				{
					resultMessage.IsDone = false;
					resultMessage.Message = "目标文件格式错误，请检查目标文件是否正确！";
					return resultMessage;
				}
			}

			//获取数据表读取行位置
			int readRow = -1;
			for (int i = 1; i < sourceSheet1.Dimension.End.Row; i++)
			{
				if (sourceSheet1.Cells[i, 1].Value != null
					&& sourceSheet1.Cells[i, 1].Value.ToString().IndexOf("明细列表") > -1)
				{
					readRow = i + 2;
					resultMessage.StartReadRow = readRow;
					break;
				}
			}

			if (readRow < 0)
			{
				resultMessage.IsDone = false;
				resultMessage.Message = "获取数据表起始行失败，请检查数据表是否正确！";
			}

			//获取目标表插入位置
			int insertRow = targetSheet1.Dimension.End.Row + 1;
			resultMessage.InsertStartRow = insertRow;
			if (!sourceTypes.TryGetValue(sourceType, out int[] sourceTypeIndexs))
			{
				resultMessage.IsDone = false;
				resultMessage.Message = "获取目标表索引失败！";
				return resultMessage;
			}

			//重新加载规则文件
			InitRuleFile();

			//复制数据
			for (; readRow <= sourceSheet1.Dimension.End.Row; readRow++)
			{
				if (null == sourceSheet1.Cells[readRow, sourceTypeIndexs[0]].Value)
				{
					break;
				}
				resultMessage.InsertCount++;
				targetSheet1.Cells[insertRow, targetIndexs[0]].Value
					= sourceSheet1.Cells[readRow, sourceTypeIndexs[0]].Value;

				targetSheet1.Cells[insertRow, targetIndexs[2]].Value =
					RemoveSpace(sourceSheet1.Cells[readRow, sourceTypeIndexs[2]].Value.ToString());
				targetSheet1.Cells[insertRow, targetIndexs[3]].Value =
					RemoveSpace(sourceSheet1.Cells[readRow, sourceTypeIndexs[3]].Value.ToString());
				targetSheet1.Cells[insertRow, targetIndexs[5]].Value =
					RemoveSpace(sourceSheet1.Cells[readRow, sourceTypeIndexs[4]].Value.ToString());
				targetSheet1.Cells[insertRow, targetIndexs[6]].Value =
					RemoveSpace(sourceSheet1.Cells[readRow, sourceTypeIndexs[5]].Value.ToString());

				targetSheet1.Cells[insertRow, targetIndexs[7]].Value = sourceType;

				if (sourceSheet1.Cells[readRow, sourceTypeIndexs[2]].Value.ToString().IndexOf("支出") > -1)
				{
					double cost = (double)sourceSheet1.Cells[readRow, sourceTypeIndexs[1]].Value;
					targetSheet1.Cells[insertRow, targetIndexs[1]].Value = (0 - cost);
				}
				else
				{
					targetSheet1.Cells[insertRow, targetIndexs[1]].Value
						= sourceSheet1.Cells[readRow, sourceTypeIndexs[1]].Value;
				}

				string commodityName = targetSheet1.Cells[insertRow, targetIndexs[5]].Value.ToString() +
									   targetSheet1.Cells[insertRow, targetIndexs[3]].Value.ToString();
				targetSheet1.Cells[insertRow, targetIndexs[4]].Value = GetRECategory(commodityName);
				insertRow++;
			}
			resultMessage.EndReadRow = readRow;
			resultMessage.InsertEndRow = insertRow;

			try
			{
				targetExcel.Save();
			}
			catch (Exception e)
			{
				resultMessage.IsDone = false;
				resultMessage.Message = e.InnerException.InnerException.Message;
				return resultMessage;
			}

			DateTime endTime = DateTime.Now;
			resultMessage.ConmudeTime = endTime - startTime;
			return resultMessage;
		}

		private string RemoveSpace(string str)
		{
			string result = str.Replace(" ", "");

			return result;
		}

		private string GetRECategory(string category)
		{
			string result = string.Empty;
			foreach (KeyValuePair<string, List<string>> item in ruleDict)
			{
				for (int i = 0; i < item.Value.Count; i++)
				{
					if (category.Contains(item.Value[i]))
					{
						result = item.Key;
						return result;
					}
				}
			}

			return result;

		}
	}
}