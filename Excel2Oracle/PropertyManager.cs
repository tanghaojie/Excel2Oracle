using Newtonsoft.Json;
using System.IO;

namespace Excel2Oracle
{
	class PropertyManager
	{
		public static string propertyJSON = "Property.json";

		private bool CheckFileExist()
		{
			return File.Exists("./" + propertyJSON);
		}

		private void CreateEmptyFile()
		{
			Property p = new Property() { ExcelPath = "", OracleDb = "", OraclePassword = "", OracleUsername = "", TableRelations = new TableRelation[1] { new TableRelation() { ExcelTablename = "", OracleTablename = "", FieldRelations = new FieldRelation[1] { new FieldRelation() { ExcelField = "", OracleField = "" } } } } };
			string json = JsonConvert.SerializeObject(p, new JsonSerializerSettings() { Formatting = Formatting.Indented, StringEscapeHandling = StringEscapeHandling.EscapeNonAscii });
			TextStreamClass.Write("./" + propertyJSON, json);
		}

		public Property GetProperties() {
			if (!CheckFileExist()) {
				CreateEmptyFile();
			}
			string json = TextStreamClass.Read("./" + propertyJSON);
			JsonSerializerSettings jsonSerializerSettings = new JsonSerializerSettings() { Formatting = Formatting.Indented, StringEscapeHandling = StringEscapeHandling.EscapeNonAscii };
			Property p = JsonConvert.DeserializeObject<Property>(json, jsonSerializerSettings);
			return p;
		}

		public bool CheckProperty(Property p) {
			if (p == null) {
				Log.WriteLog("配置信息错误，删除配置文件重试", Log.LogType.Error);
				return false;
			}
			if (string.IsNullOrEmpty(p.ExcelPath)) {
				Log.WriteLog("空的Excel路径", Log.LogType.Error);
				return false;
			}
			if (string.IsNullOrEmpty(p.OracleDb))
			{
				Log.WriteLog("空的Oracle数据库", Log.LogType.Error);
				return false;
			}
			if (string.IsNullOrEmpty(p.OracleUsername))
			{
				Log.WriteLog("空的Oracle用户名", Log.LogType.Error);
				return false;
			}
			if (string.IsNullOrEmpty(p.OraclePassword))
			{
				Log.WriteLog("空的Oracle密码", Log.LogType.Error);
				return false;
			}

			TableRelation[] tableRelations = p.TableRelations;
			if (tableRelations==null || tableRelations.Length<=0)
			{
				Log.WriteLog("缺少表关系", Log.LogType.Error);
				return false;
			}
			foreach (var tr in tableRelations)
			{
				if (string.IsNullOrEmpty(tr.ExcelTablename))
				{
					Log.WriteLog("空的Excel表名", Log.LogType.Error);
					return false;
				}
				if (string.IsNullOrEmpty(tr.OracleTablename))
				{
					Log.WriteLog("空的Oracle表名", Log.LogType.Error);
					return false;
				}
				FieldRelation[] fieldRelations = tr.FieldRelations;
				if (fieldRelations == null || fieldRelations.Length <= 0) {
					Log.WriteLog("缺少字段关系", Log.LogType.Error);
					return false;
				}
				foreach (var fr in fieldRelations)
				{
					if (string.IsNullOrEmpty(fr.ExcelField))
					{
						Log.WriteLog("空的Excel字段名", Log.LogType.Error);
						return false;
					}
					if (string.IsNullOrEmpty(fr.OracleField))
					{
						Log.WriteLog("空的Oracle字段名", Log.LogType.Error);
						return false;
					}
				}
			}

			return true;
		}

	}
}
