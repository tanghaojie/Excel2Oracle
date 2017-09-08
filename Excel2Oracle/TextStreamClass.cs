using System;
using System.IO;
using System.Text;

namespace Excel2Oracle
{
	class TextStreamClass
	{
		public static void Append(string filepath, string content)
		{
			StreamWriter sw = File.AppendText(filepath);
			sw.Write(content);
			sw.Flush();
			sw.Close();
		}
		public static void Write(string filepath, string content)
		{
			FileStream fs = new FileStream(filepath, FileMode.Create);
			StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
			sw.Write(content);
			sw.Flush();
			sw.Close();
			fs.Close();
		}

		public static string Read(string path)
		{
			string content = "";
			StreamReader sr = new StreamReader(path, Encoding.UTF8);
			String line;
			while ((line = sr.ReadLine()) != null)
			{
				content += line + "\r\n";
			}
			sr.Close();
			return content;
		}
	}
}
