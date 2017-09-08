using System;

namespace Excel2Oracle
{
	class Program
	{
		private static PropertyManager pm = new PropertyManager();
		private static ExcelManage e = new ExcelManage();
		public static OracleManage o = new OracleManage();
		static void Main(string[] args)
		{
			try
			{
				Log.WriteLog("启动", Log.LogType.Info);

				Log.WriteLog("检查" + PropertyManager.propertyJSON + "", Log.LogType.Info);
				var property = pm.GetProperties();
				if (!pm.CheckProperty(property))
				{
					Log.WriteLog("停止\r\n\r\n", Log.LogType.Warn);
					return;
				}
				Log.WriteLog("检查" + PropertyManager.propertyJSON + " 通过", Log.LogType.Info);

				Log.WriteLog("检查Excel", Log.LogType.Info);
				if (!e.CheckExcel(property))
				{
					Log.WriteLog("停止\r\n\r\n", Log.LogType.Warn);
					return;
				}
				Log.WriteLog("检查Excel 通过", Log.LogType.Info);

				Log.WriteLog("检查Oracle", Log.LogType.Info);
				if (!o.CheckOracle(property))
				{
					Log.WriteLog("停止\r\n\r\n", Log.LogType.Warn);
					return;
				}
				Log.WriteLog("检查Oracle 通过", Log.LogType.Info);

				Log.WriteLog("开始操作", Log.LogType.Info);
				E2O.Excel2Oracle(property);

				Log.WriteLog("停止\r\n\r\n", Log.LogType.Info);
			}
			catch (Exception ex)
			{
				Log.WriteLog(ex.Message + "\t" + ex.Source, Log.LogType.Error);
				Log.WriteLog("系统Crashed\r\n\r\n", Log.LogType.Error);
			}
		}


	}
}
