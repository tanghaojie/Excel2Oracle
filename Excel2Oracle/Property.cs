namespace Excel2Oracle
{
	class Property
	{
		public string ExcelPath { get; set; }
		public string OracleDb { get; set; }
		public string OracleUsername { get; set; }
		public string OraclePassword { get; set; }
		public TableRelation[] TableRelations { get; set; }
	}

	class TableRelation
	{
		public string ExcelTablename { get; set; }
		public string OracleTablename { get; set; }
		public FieldRelation[] FieldRelations { get; set; }
	}

	class FieldRelation
	{
		public string ExcelField { get; set; }
		public string OracleField { get; set; }
	}
}
