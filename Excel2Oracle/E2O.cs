using ExcelDataReader;
using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Excel2Oracle
{
	class E2O
	{
		public static void Excel2Oracle(Property property)
		{
			string ePath = property.ExcelPath;
			string oDb = property.OracleDb;
			string oUsername = property.OracleUsername;
			string oPassword = property.OraclePassword;

			using (var stream = File.Open(ePath, FileMode.Open, FileAccess.Read))
			{
				using (var excelReader = ExcelReaderFactory.CreateReader(stream))
				{
					if (excelReader == null)
					{
						Log.WriteLog("未能正确打开Excel文件", Log.LogType.Error);
						return;
					}
					string connStr = Program.o.GetConnectionString(oDb, oUsername, oPassword);
					using (OracleConnection connection = new OracleConnection(connStr))
					{
						connection.Open();
						foreach (var tr in property.TableRelations)
						{
							string eTablename = tr.ExcelTablename;
							string oTablename = tr.OracleTablename;
							FieldRelation[] frs = tr.FieldRelations;
							excelReader.Reset();
							bool find = false;
							do
							{
								if (excelReader.Name.ToUpper() == eTablename.ToUpper())
								{
									find = true;
									break;
								}
							} while (excelReader.NextResult());
							if (!find)
							{
								Log.WriteLog("Excel中未找到表[" + eTablename + "]，已经跳过此表导入", Log.LogType.Error);
								continue;
							}

							Dictionary<int, string> eoField = new Dictionary<int, string>();
							if (!excelReader.Read())
							{
								Log.WriteLog("Excel表[" + eTablename + "]，没有读取到字段信息，已经跳过此表导入", Log.LogType.Error);
								continue;
							}
							int count = excelReader.FieldCount;
							for (int i = 0; i < count; i++)
							{
								object x = excelReader.GetValue(i);
								if (x != null && x != DBNull.Value)
								{
									string exf = x.ToString();
									foreach (var fr in frs)
									{
										string ef = fr.ExcelField;
										string of = fr.OracleField;
										if (exf.ToUpper() == ef.ToUpper())
										{
											eoField.Add(i, of);
										}
									}
								}
							}

							Dictionary<string, object> oracleRow = new Dictionary<string, object>();
							int rowNum = 1;
							while (excelReader.Read())
							{
								rowNum++;
								oracleRow = new Dictionary<string, object>();
								foreach (var eo in eoField)
								{
									int cNum = eo.Key;
									object value = excelReader.GetValue(cNum);
									if (value == null || value == DBNull.Value)
									{
										value = "";
									}
									string of = eo.Value;
									oracleRow.Add(of, value);
								}
								string strF = "";
								string strV = "";
								DataColumnCollection oracleColumns = null;
								try
								{
									string sql1 = "select * from " + oTablename + " where rownum=1";
									DataSet ds = null;
									using (OracleCommand cmd = new OracleCommand())
									{
										cmd.Connection = connection;
										cmd.CommandText = sql1;
										using (OracleDataAdapter adapter = new OracleDataAdapter(cmd))
										{
											ds = new DataSet();
											adapter.Fill(ds);
										}
									}
									if (ds == null || ds.Tables.Count <= 0)
									{
										Log.WriteLog("未能正确打开Oracle表[" + oTablename + "]，跳过此表导入", Log.LogType.Error);
										continue;
									}
									oracleColumns = ds.Tables[0].Columns;
								}
								catch (Exception ex)
								{
									Log.WriteLog("未能正确打开Oracle表[" + oTablename + "]，跳过此表导入。错误信息：" + ex.Message + "  " + ex.Source, Log.LogType.Error);
									continue;
								}
								foreach (var x in oracleRow)
								{
									string f = x.Key;
									object v = x.Value;
									DataColumn c = oracleColumns[f];
									strF += f + ",";

									string xx = "";
									if (c != null)
									{
										if (c.DataType == typeof(DateTime))
										{
											xx = "to_Date('" + x.Value.ToString() + "','yyyy/mm/dd hh24:mi:ss')";
										}
										else if (c.DataType == typeof(int) || c.DataType == typeof(double) || c.DataType == typeof(float) || c.DataType == typeof(decimal))
										{
											xx = x.Value.ToString();
										}
										else
										{
											xx = "'" + x.Value.ToString() + "'";
										}
									}
									strV += xx + ",";
								}
								strF = strF.Substring(0, strF.Length - 1);
								strV = strV.Substring(0, strV.Length - 1);

								string sql = "INSERT INTO " + oTablename + "(" + strF + ") VALUES(" + strV + ")";
								try
								{
									using (OracleCommand cmd = new OracleCommand())
									{
										cmd.Connection = connection;
										cmd.CommandText = sql;
										if (cmd.ExecuteNonQuery() <= 0)
										{
											Log.WriteLog("Excel表[" + eTablename + "]行号:" + rowNum + "，数据出错，没有正确导入，跳过此行。错误信息：没有插入数据库。\t sql：" + sql, Log.LogType.Error);
											continue;
										}
									}
								}
								catch (Exception ex)
								{
									Log.WriteLog("Excel表[" + eTablename + "]行号:" + rowNum + "，数据出错，没有正确导入，跳过此行。错误信息：" + ex.Message + " " + ex.Source + "\t sql：" + sql, Log.LogType.Error);
									continue;
								}
							}
						}
					}
				}
			}
		}


	}
}
