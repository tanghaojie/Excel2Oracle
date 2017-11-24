using System;
using System.IO;
using System.Text;

namespace Excel2Oracle
{
	class TextStreamClass
	{
		public static void Append(string filepath, string content)
		{
            using (StreamWriter sw = File.AppendText(filepath))
            {
                sw.Write(content);
                sw.Flush();
            }
		}
		public static void Write(string filepath, string content)
		{
            using (FileStream fs = new FileStream(filepath, FileMode.Create))
            {
                using (StreamWriter sw = new StreamWriter(fs, Encoding.UTF8))
                {
                    sw.Write(content);
                    sw.Flush();
                }
            }
		}
		public static string Read(string path)
		{
			string content = "";
            using (StreamReader sr = new StreamReader(path, Encoding.UTF8))
            {
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    content += line + "\r\n";
                }
            }
			return content;
		}
	}
}
