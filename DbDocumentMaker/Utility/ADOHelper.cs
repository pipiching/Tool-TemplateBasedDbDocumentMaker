using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbDocumentMaker.Utility
{
    public class QueryField
    {
        public string ColumnName { get; set; }
        public int ColumnOrdinal { get; set; }
        public int ColumnSize { get; set; }
        public int NumericPrecision { get; set; }
        public int NumericScale { get; set; }
        public bool IsUnique { get; set; }
        public string BaseColumnName { get; set; }
        public string BaseTableName { get; set; }
        public string DataType { get; set; }
        public bool AllowDBNull { get; set; }
        public string ProviderType { get; set; }
        public bool IsIdentity { get; set; }
        public bool IsAutoIncrement { get; set; }
        public bool IsRowVersion { get; set; }
        public bool IsLong { get; set; }
        public bool IsReadOnly { get; set; }
        public string ProviderSpecificDataType { get; set; }
        public string DataTypeName { get; set; }
        public string UdtAssemblyQualifiedName { get; set; }
        public int NewVersionedProviderType { get; set; }
        public string IsColumnSet { get; set; }
        public string RawProperties { get; set; }
        public string NonVersionedProviderType { get; set; }
    }

    public class ADOHelper
    {

        public ADOHelper()
        {
        }

        //https://docs.microsoft.com/zh-tw/dotnet/api/system.data.sqlclient.sqldatareader.getschematable?view=dotnet-plat-ext-6.0
        private DataTable GetQuerySchema(string strconn, string strSQL)
        {
            // Returns a DataTable filled with the results of the query
            // Function returns the count of records in the datatable
            // ----- dt (datatable) needs to be empty & no schema defined

            var sconQuery = new SqlConnection();
            var scmdQuery = new SqlCommand();
            SqlDataReader srdrQuery = null;
            int intRowsCount = 0;
            var dtSchema = new DataTable();

            try
            {

                // Open the SQL connnection to the SWO database
                sconQuery.ConnectionString = strconn;
                sconQuery.Open();

                // Execute the SQL command against the database & return a resultset
                scmdQuery.Connection = sconQuery;
                scmdQuery.CommandText = strSQL;
                srdrQuery = scmdQuery.ExecuteReader(CommandBehavior.SchemaOnly);

                dtSchema = srdrQuery.GetSchemaTable();
            }
            catch (Exception ex)
            {
                Information.Err().Raise(-1000, Description: "Error = '" + ex.Message + " ': sql = " + strSQL);
            }
            finally
            {
                if (!(srdrQuery == null))
                {
                    if (!srdrQuery.IsClosed)
                        srdrQuery.Close();
                }
                scmdQuery.Dispose();
                sconQuery.Close();
                sconQuery.Dispose();
            }

            return dtSchema;
        }
        public List<QueryField> GetFields(string ConnectionString, string Query, ref DataTable SchemaTable)
        {
            var dt = new DataTable();
            if (Query.Trim().ToLower().StartsWith("select "))
            {
                SchemaTable = GetQuerySchema(ConnectionString, Query);
            }
            else
            {
                SchemaTable = GetSPSchema(Query, ConnectionString);
            }
            var result = new List<QueryField>();

            for (int i = 0, loopTo = SchemaTable.Rows.Count - 1; i <= loopTo; i++)
            {
                var qf = new QueryField();
                string properties = string.Empty;
                for (int j = 0, loopTo1 = SchemaTable.Columns.Count - 1; j <= loopTo1; j++)
                {
                    properties += SchemaTable.Columns[j].ColumnName + Strings.Chr(254) + SchemaTable.Rows[i][j].ToString();
                    if (j < SchemaTable.Columns.Count - 1)
                        properties += Convert.ToString(Strings.Chr(255));

                    if (SchemaTable.Rows[i][j] is DBNull == false)
                    {
                        switch (SchemaTable.Columns[j].ColumnName ?? "")
                        {
                            case "ColumnName":
                                {
                                    qf.ColumnName = Convert.ToString(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "ColumnOrdinal":
                                {                             
                                    qf.ColumnOrdinal = Convert.ToInt32(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "ColumnSize":
                                {
                                    qf.ColumnSize = Convert.ToInt32(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "NumericPrecision":
                                {
                                    qf.NumericPrecision = Convert.ToInt32(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "NumericScale":
                                {
                                    qf.NumericScale = Convert.ToInt32(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "IsUnique":
                                {
                                    qf.IsUnique = Convert.ToBoolean(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "BaseColumnName":
                                {
                                    qf.BaseColumnName = Convert.ToString(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "BaseTableName":
                                {
                                    qf.BaseTableName = Convert.ToString(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "DataType":
                                {
                                    qf.DataType = ((Type)SchemaTable.Rows[i][j]).FullName;
                                    break;
                                }
                            case "AllowDBNull":
                                {
                                    qf.AllowDBNull = Convert.ToBoolean(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "ProviderType":
                                {
                                    qf.ProviderType = Convert.ToString(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "IsIdentity":
                                {
                                    qf.IsIdentity = Convert.ToBoolean(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "IsAutoIncrement":
                                {
                                    qf.IsAutoIncrement = Convert.ToBoolean(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "IsRowVersion":
                                {
                                    qf.IsRowVersion = Convert.ToBoolean(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "IsLong":
                                {
                                    qf.IsLong = Convert.ToBoolean(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "IsReadOnly":
                                {
                                    qf.IsReadOnly = Convert.ToBoolean(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "ProviderSpecificDataType":
                                {
                                    qf.ProviderSpecificDataType = ((Type)SchemaTable.Rows[i][j]).FullName;
                                    break;
                                }
                            case "DataTypeName":
                                {
                                    qf.DataTypeName = Convert.ToString(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "UdtAssemblyQualifiedName":
                                {
                                    qf.UdtAssemblyQualifiedName = Convert.ToString(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "IsColumnSet":
                                {
                                    qf.IsColumnSet = Convert.ToString(SchemaTable.Rows[i][j]);
                                    break;
                                }
                            case "NonVersionedProviderType":
                                {
                                    qf.NonVersionedProviderType = Convert.ToString(SchemaTable.Rows[i][j]);
                                    break;
                                }

                            default:
                                {
                                    break;
                                }
                        }
                    }
                }
                qf.RawProperties = properties;
                result.Add(qf);
            }

            return result;
        }
        public string[] GenerateCodeCS(ref List<QueryField> Columns, string ObjectName, string LinePrefix = "    ")
        {
            var result = new string[Columns.Count + 2 + 1];
            result[0] = string.Format("public class {0} {{", ObjectName);
            for (int i = 0, loopTo = Columns.Count - 1; i <= loopTo; i++)
            {
                try
                {

                    if (char.IsNumber(Convert.ToChar(Columns[i].ColumnName.Substring(0, 1))))
                    {
                        Columns[i].ColumnName = "_" + Columns[i].ColumnName;
                    }

                    string AllowNull = ", null";
                    if (Columns[i].AllowDBNull == false)
                        AllowNull = ", not null";

                    switch (Columns[i].DataTypeName ?? "")
                    {

                        case "bigint":
                            {
                                result[i + 1] = string.Format("{0}public long {1} {{ get; set; }} //(bigint{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "binary":
                            {
                                result[i + 1] = string.Format("{0}public byte[] {1} {{ get; set; }} //(binary({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull);
                                break;
                            }

                        case "bit":
                            {
                                result[i + 1] = string.Format("{0}public bool {1} {{ get; set; }} //(bit{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "char":
                            {
                                result[i + 1] = string.Format("{0}public string {1} {{ get; set; }} //(char({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull);
                                break;
                            }

                        case "date":
                            {
                                result[i + 1] = string.Format("{0}public DateTime {1} {{ get; set; }} //(date{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "datetime":
                            {
                                result[i + 1] = string.Format("{0}public DateTime {1} {{ get; set; }} //(datetime{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "datetime2":
                            {
                                result[i + 1] = string.Format("{0}public DateTime {1} {{ get; set; }} //(datetime2({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].NumericScale, AllowNull);
                                break;
                            }

                        case "datetimeoffset":
                            {
                                result[i + 1] = string.Format("{0}public DateTimeOffset {1} {{ get; set; }} //(datetimeoffset{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "decimal":
                            {
                                result[i + 1] = string.Format("{0}public decimal {1} {{ get; set; }} //(decimal({2},{3}){4})", LinePrefix, Columns[i].ColumnName, Columns[i].NumericPrecision, Columns[i].NumericScale, AllowNull);
                                break;
                            }

                        case "float":
                            {
                                result[i + 1] = string.Format("{0}public double {1} {{ get; set; }} //(float{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "image":
                            {
                                result[i + 1] = string.Format("{0}public byte[] {1} {{ get; set; }} //(image{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "int":
                            {
                                result[i + 1] = string.Format("{0}public int {1} {{ get; set; }} //(int{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "money":
                            {
                                result[i + 1] = string.Format("{0}public decimal {1} {{ get; set; }} //(money{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "nchar":
                            {
                                if (Columns[i].IsLong)
                                {
                                    result[i + 1] = string.Format("{0}public string {1} {{ get; set; }} //(nchar(max){2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                }
                                else
                                {
                                    result[i + 1] = string.Format("{0}public string {1} {{ get; set; }} //(nchar({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull);
                                }

                                break;
                            }

                        case "ntext":
                            {
                                result[i + 1] = string.Format("{0}public string {1} {{ get; set; }} //(ntext{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "nvarchar":
                            {
                                if (Columns[i].IsLong)
                                {
                                    result[i + 1] = string.Format("{0}public string {1} {{ get; set; }} //(nvarchar(max){2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                }
                                else
                                {
                                    result[i + 1] = string.Format("{0}public string {1} {{ get; set; }} //(nvarchar({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull);
                                }

                                break;
                            }

                        case "real":
                            {
                                result[i + 1] = string.Format("{0}public Single {1} {{ get; set; }} //(real({2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "smalldatetime":
                            {
                                result[i + 1] = string.Format("{0}public DateTime {1} {{ get; set; }} //(smalldatetime{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "sql_variant":
                            {
                                result[i + 1] = string.Format("{0}public object {1} {{ get; set; }} //(sql_variant{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "text":
                            {
                                result[i + 1] = string.Format("{0}public string {1} {{ get; set; }} //(text{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "time":
                            {
                                result[i + 1] = string.Format("{0}public DateTime {1} {{ get; set; }} //(time({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].NumericScale, AllowNull);
                                break;
                            }

                        case "timestamp":
                            {
                                result[i + 1] = string.Format("{0}public byte[] {1} {{ get; set; }} //(timestamp{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "tinyint":
                            {
                                result[i + 1] = string.Format("{0}public byte {1} {{ get; set; }} //(tinyint{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "uniqueidentifier":
                            {
                                result[i + 1] = string.Format("{0}public Guid {1} {{ get; set; }} //(uniqueidentifier{2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                break;
                            }

                        case "varbinary":
                            {
                                if (Columns[i].IsLong)
                                {
                                    result[i + 1] = string.Format("{0}public byte[] {1} {{ get; set; }} //(varbinary(max){2})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull);
                                }
                                else
                                {
                                    result[i + 1] = string.Format("{0}public byte[] {1} {{ get; set; }} //(varbinary({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, Columns[i].ColumnSize, AllowNull);
                                }

                                break;
                            }

                        case "varchar":
                            {
                                if (Columns[i].IsLong)
                                {
                                    result[i + 1] = string.Format("{0}public string {1} {{ get; set; }} //(varchar(max){2})", LinePrefix, Columns[i].ColumnName, AllowNull);
                                }
                                else
                                {
                                    result[i + 1] = string.Format("{0}public string {1} {{ get; set; }} //(varchar({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull);
                                }

                                break;
                            }

                        case "xml":
                            {
                                result[i + 1] = string.Format("{0}public string {1} {{ get; set; }} //(XML(.){2})", LinePrefix, Columns[i].ColumnName, AllowNull); // sql variant
                                break;
                            }

                        default:
                            {
                                switch (Columns[i].DataType ?? "")
                                {

                                    case "Microsoft.SqlServer.Types.SqlGeography": // geography
                                        {
                                            result[i + 1] = string.Format("{0}public Microsoft.SqlServer.Types.SqlGeography {1} {{ get; set; }} //({2}{3})", LinePrefix, Columns[i].ColumnName, Columns[i].DataTypeName, AllowNull);
                                            break;
                                        }

                                    case "Microsoft.SqlServer.Types.SqlHierarchyId": // heirarchyid
                                        {
                                            result[i + 1] = string.Format("{0}public Microsoft.SqlServer.Types.SqlGeography {1} {{ get; set; }} //({2}{3})", LinePrefix, Columns[i].ColumnName, Columns[i].DataTypeName, AllowNull);
                                            break;
                                        }

                                    case "Microsoft.SqlServer.Types.SqlGeometry": // geometry
                                        {
                                            result[i + 1] = string.Format("{0}public Microsoft.SqlServer.Types.SqlGeography {1} {{ get; set; }} //({2}{3})", LinePrefix, Columns[i].ColumnName, Columns[i].DataTypeName, AllowNull);
                                            break;
                                        }

                                }

                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }
            result[Columns.Count + 1] = "}";
            return result;
        }
        public string StringArrayToText(string[] lines, string LineDelimiter = Constants.vbCrLf)
        {
            // Determine number of characters needed in result string
            int charCount = 0;
            foreach (var s in lines)
            {
                if (!(s == null))
                {
                    charCount += s.Length + LineDelimiter.Length;
                }
            }

            // Preallocate needed string space plus a little extra
            var sb = new System.Text.StringBuilder(charCount + lines.Count());
            for (int i = 0, loopTo = lines.Count() - 1; i <= loopTo; i++)
            {
                if (!(lines[i] == null))
                {
                    sb.Append(lines[i] + LineDelimiter);
                }
            }
            return sb.ToString();
        }

        // Run a stored procedure with optional parameters and return as datatable of results
        public DataTable GetSPSchema(string spName, string strConn)
        {

            // Dim the dataset to hold the result table
            var dtSchema = new DataTable();

            // Return if missing important information (SP Name)
            if (spName.Trim().Length == 0)
                return dtSchema;

            // Default the connection string to the public class variable if not specified
            if (strConn.Length == 0)
                return dtSchema;

            // Create the connection to the database
            var sconSP = new SqlConnection();
            var scmdSP = new SqlCommand();
            SqlDataReader srdrSP = null;

            try
            {
                // Set the connection string on the connection object
                sconSP.ConnectionString = strConn;
                sconSP.Open();

                // Set up the SqlCommand object
                scmdSP.CommandText = spName;
                scmdSP.Connection = sconSP;
                scmdSP.CommandType = CommandType.StoredProcedure;

                if (spName.Contains("|"))
                {
                    var spParms = spName.Split('|');
                    scmdSP.CommandText = spParms[0].Trim();
                    for (int i = 1, loopTo = spParms.Count() - 1; i <= loopTo; i++)
                    {

                        var spFields = spParms[i].Replace("`,", Constants.vbTab).Replace("`", "").Split(Convert.ToChar(Constants.vbTab));
                        scmdSP.Parameters.Add(new SqlParameter(spFields[0].Trim(), spFields[1].Trim()));
                    }
                }

                srdrSP = scmdSP.ExecuteReader(CommandBehavior.SchemaOnly);
                dtSchema = srdrSP.GetSchemaTable();
            }
            catch (Exception ex)
            {
                Information.Err().Raise(-1000, Description: "Error = '" + ex.Message + " ': stored procedure = " + spName);
            }
            finally
            {
                if (!(srdrSP == null))
                {
                    if (!srdrSP.IsClosed)
                        srdrSP.Close();
                }
                scmdSP.Dispose();
                sconSP.Close();
                sconSP.Dispose();
            }

            return dtSchema;
        }
    }
}
