using System;

namespace Excel2Oracle
{
	public class Log
	{
		static string logFile = "Log.txt";
		public enum LogType {
			Error,
			Warn,
			Info,
			Other
		}
		public static void WriteLog(string text, LogType logType)
		{
			DateTime now = DateTime.Now;
			string strDate = now.ToLongDateString();
			string strTime = now.ToLongTimeString();

			string strLogType = Enum.GetName(typeof(LogType), logType).PadRight(6, ' ');

			string total = strDate + " " + strTime + "\t" + strLogType + "\t" + text + "\r\n";
			TextStreamClass.Append("./" + logFile, total);
		}
	}
}
