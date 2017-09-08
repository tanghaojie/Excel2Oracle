using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;

namespace Excel2Oracle
{
	class ExcelManage
	{
		public bool CheckExcel(Property property)
		{
			var ePath = property.ExcelPath;

			if (!File.Exists(ePath) || !(ePath.ToUpper().EndsWith(".XLS") || ePath.ToUpper().EndsWith(".XLSX")))
			{
				Log.WriteLog("Excel文件不存在或目录错误", Log.LogType.Error);
				return false;
			}
			bool flag = true;
			using (var stream = File.Open(ePath, FileMode.Open, FileAccess.Read))
			{
				using (var excelReader = ExcelReaderFactory.CreateReader(stream))
				{
					if (excelReader == null)
					{
						Log.WriteLog("未能正确打开Excel文件", Log.LogType.Error);
						flag = false;
					}
					else
					{
						int tableCount = excelReader.ResultsCount;
						if (tableCount <= 0)
						{
							Log.WriteLog("Excel文件中没有表", Log.LogType.Error);
							flag = false;
						}
						else
						{
							List<string> listTablenames = new List<string>();
							excelReader.Reset();
							do
							{
								string n = excelReader.Name.ToUpper();
								if (!listTablenames.Contains(n))
								{
									listTablenames.Add(n);
								}
							} while (excelReader.NextResult());

							var trs = property.TableRelations;
							foreach (var tr in trs)
							{
								if (!flag)
								{
									break;
								}
								excelReader.Reset();
								string eTName = tr.ExcelTablename.ToUpper();
								if (!listTablenames.Contains(eTName))
								{
									Log.WriteLog("Excel文件中未找到表[" + eTName + "]", Log.LogType.Error);
									flag = false;
									break;
								}
								else
								{
									do
									{
										if (!flag)
										{
											break;
										}
										string name = excelReader.Name;
										if (name.ToUpper() == eTName)
										{
											int count = excelReader.FieldCount;
											FieldRelation[] frs = tr.FieldRelations;
											List<string> eFields = new List<string>();
											if (excelReader.Read())
											{
												for (int i = 0; i < count; i++)
												{
													object x = excelReader.GetValue(i);
													if (x != null && x != DBNull.Value)
													{
														string ef = x.ToString();
														if (!String.IsNullOrEmpty(ef))
														{
															ef = ef.ToUpper();
														}
														if (!eFields.Contains(ef))
														{
															eFields.Add(ef);
														}
													}
												}
											}
											else
											{
												Log.WriteLog("Excel文件[" + eTName + "]表中没有字段数据信息", Log.LogType.Error);
												flag = false;
												break;
											}
											foreach (var fr in frs)
											{
												string eFName = fr.ExcelField;
												if (!eFields.Contains(eFName.ToUpper()))
												{
													Log.WriteLog("Excel文件[" + eTName + "]表中，不包含[" + eFName + "]字段", Log.LogType.Error);
													flag = false;
													break;
												}
											}
											break;
										}
									} while (excelReader.NextResult());
								}
							}
						}
					}
				}
			}
			return flag;
		}
	}
}
