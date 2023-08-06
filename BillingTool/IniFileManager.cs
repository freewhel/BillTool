using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BillingTool
{
	public class IniFileManager
	{
		private string filePath;

		public IniFileManager(string filePath)
		{
			this.filePath = filePath;
		}

		public Dictionary<string, List<string>> GetSectionsKeys()
		{
			Dictionary<string, List<string>> result = new Dictionary<string, List<string>>();
			List<string> keys = new List<string>();
			try
			{
				StreamReader reader = new StreamReader(this.filePath);
				string line = string.Empty;
				string section = string.Empty;
				string lastSection = string.Empty;
				while (null != (line = reader.ReadLine()))
				{
					if (!string.IsNullOrEmpty(line))
					{
						if (line.Contains("["))
						{

							section = line.Substring(1, line.Length - 2);

							if (!string.IsNullOrEmpty(lastSection) && !lastSection.Equals(section))
							{

								result.Add(lastSection, keys);
								keys = new List<string>();
							}
							lastSection = section.Equals(lastSection) ? lastSection : section;

						}
						else
						{
							keys.Add(line.Trim());
						}
					}

				}
			}
			catch (Exception)
			{
			}

			return result;
		}
	}
}
