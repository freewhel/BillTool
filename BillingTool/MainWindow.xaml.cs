using OfficeOpenXml;
using System.Windows;
using System.IO;
using System.Linq;
using System.Windows.Controls;

namespace BillingTool
{
	/// <summary>
	/// MainWindow.xaml 的交互逻辑
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			string sourcePath = sourceTextBox.Text;
			string targetPath = targeTextBox.Text;

			ExcelHandler excelHander = new ExcelHandler(sourcePath, targetPath);
			ResultMessage result = excelHander.SoucerToTargetBill();
			if (result.IsDone)
			{
				resultTextBlock.Text = $"转换成功\r\n" +
					"----------------------------------------------------------\r\n" +
					$"读取文件路径：{sourcePath}\r\n" +
					$"读取起始行：{result.StartReadRow}\r\n" +
					$"读取结束行：{result.EndReadRow}\r\n" +
					$"读取数量：{result.ReadRowCount}\r\n" +
					"-----------------------------------------------------------\r\n" +
					$"插入文件路径：{targetPath}\r\n" +
					$"插入起始行：{result.InsertStartRow} \r\n" +
					$"插入结束行：{result.InsertEndRow} \r\n" +
					$"插入总行数：{result.InsertCount} \r\n" +
					"-----------------------------------------------------------\r\n" +
					$"耗时：{result.ConmudeTimeStr}";
			}
			else
			{
				resultTextBlock.Text = $"转换失败\r\n" +
					"----------------------------------------------------------\r\n" +
					$"错误信息：{result.Message}\r\n";
			}

		}

		private void TextBox_Drop(object sender, DragEventArgs e)
		{
			string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
			if (files != null && File.Exists(files[0]))
			{
				TextBox thisTextBox = sender as TextBox;
				thisTextBox.Text = files[0];
			}
		}

		private void TextBox_DragOver(object sender, DragEventArgs e)
		{
			e.Effects = DragDropEffects.Copy;
			e.Handled = true;
		}
	}
}