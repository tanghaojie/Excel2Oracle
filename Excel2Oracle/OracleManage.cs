using Oracle.DataAccess.Client;
using System;
using System.Data;

namespace Excel2Oracle
{
	class OracleManage
	{
		public bool CheckOracle(Property property)
		{
			string db = property.OracleDb;
			string username = property.OracleUsername;
			string password = property.OraclePassword;
			string connStr = GetConnectionString(db, username, password);
			if (string.IsNullOrEmpty(connStr))
			{
				Log.WriteLog("Oracle连接配置错误", Log.LogType.Error);
				return false;
			}

			TableRelation[] trs = property.TableRelations;
			foreach (var tr in trs)
			{
				string table = tr.OracleTablename;
				if (!CheckTableExist(connStr, table))
				{
					Log.WriteLog("Oracle数据库中不存在表[" + table + "]", Log.LogType.Error);
					return false;
				}
				FieldRelation[] frs = tr.FieldRelations;
				DataTable dt = GetDataTableWithFirstRow(connStr, table);
				if (dt == null) {
					Log.WriteLog("未能正确打开Oracle表[" + table + "]", Log.LogType.Error);
					return false;
				}
				foreach (var fr in frs)
				{
					string of = fr.OracleField;
					if (!dt.Columns.Contains(of)) {
						Log.WriteLog("Oracle表[" + table + "]中，不存在字段[" + of + "]", Log.LogType.Error);
						return false;
					}
				}
			}
			return true;
		}

		public string GetConnectionString(string db, string username, string password)
		{
			db = "tcp://" + db;
			Uri u = new Uri(db);
			string ip = u.Host;
			string port = u.Port.ToString();
			string d = u.AbsolutePath?.Replace("/", "");
			if (string.IsNullOrEmpty(ip) || string.IsNullOrEmpty(port) || string.IsNullOrEmpty(d))
			{
				return null;
			}
			return "Data Source = (DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = " + ip + ")(PORT = " + port + "))(CONNECT_DATA = (SERVICE_NAME = " + d + "))); User Id = " + username + "; Password = " + password + ";";
		}

		private DataTable GetDataTableWithFirstRow(string connStr, string tablename)
		{
			string sql = "select * from " + tablename + " where rownum=1";
			DataSet dataSet = Query(connStr, sql);
			if (dataSet == null || dataSet.Tables.Count <= 0)
			{
				return null;
			}
			return dataSet.Tables[0];
		}

		private bool CheckTableExist(string connStr, string tablename)
		{
			string sql = "select count(*) from user_tables where table_name = '" + tablename.ToUpper() + "'";
			DataSet dataSet = Query(connStr, sql);
			if (dataSet == null || dataSet.Tables.Count <= 0)
			{
				return false;
			}
			DataTable dt = dataSet.Tables[0];
			object ov = dt.Rows[0][0];
			if (ov == null || ov == DBNull.Value)
			{
				return false;
			}
			int count = 0;
			if (int.TryParse(ov.ToString(), out count))
			{
				if (count > 0)
				{
					return true;
				}
				else
				{
					return false;
				}
			}
			else
			{
				return false;
			}
		}

		private DataSet Query(string connStr, string cmdText)
		{
			DataSet ds = null;
			using (OracleConnection connection = new OracleConnection(connStr))
			{
				connection.Open();
				using (OracleCommand cmd = new OracleCommand())
				{
					cmd.Connection = connection;
					cmd.CommandText = cmdText;
					using (OracleDataAdapter adapter = new OracleDataAdapter(cmd))
					{
						ds = new DataSet();
						adapter.Fill(ds);
					}
				}
			}
			return ds;
		}

	}
}
