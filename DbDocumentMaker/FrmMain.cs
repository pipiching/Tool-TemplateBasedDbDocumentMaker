using DbDocumentMaker.Models;
using DbDocumentMaker.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

namespace DbDocumentMaker
{
    public partial class FrmMain : Form
    {
        // Fields
        private DbManager _dbManager;


        // Constructors
        public FrmMain()
        {
            InitializeComponent();

            // init worker
            _dbManager = new DbManager(Config.GetInstance().Content.CurrentConnection.Str);
            _dbManager.SystemName = Config.GetInstance().Content.CurrentConnection.SystemName;
            _dbManager.SystemDescription = Config.GetInstance().Content.CurrentConnection.SystemDescription;

            InitRichTextBoxContextMenu(txtPOCO); //增加右鍵選單

        }


        // Methods
        private void ShowTables(List<Table> dbTables)
        {
            // reset ui
            clbTables.Items.Clear();
            dgvColumns.DataSource = null;

            // find selected table names for current connection in config
            var checkedTableNames = new List<string>();
            var connName = Config.GetInstance().Content.CurrentConnectionName;
            var docTablePackages = Config.GetInstance().Content.DocTablePackages;
            if (docTablePackages.ContainsKey(connName))
            {
                checkedTableNames = docTablePackages[connName];
            }

            // display table list
            clbTables.Items.Clear();
            foreach (var table in dbTables)
            {
                var isChecked = checkedTableNames.Contains(table.TableName);
                clbTables.Items.Add(table, isChecked);
            }
        }

        private void ShowTableColumns(string tableName)
        {
            dgvColumns.DataSource =
                _dbManager.DbTables.Where(t => t.TableName == tableName).First()
                        .Columns
                        .Select(c => new
                        {
                            Index = c.No,
                            ColumnName = c.ColumnName,
                            Type = c.FullDataType,
                            Nullable = c.IsNullable,
                            PK = c.IsPrimaryKey,
                            FK = c.IsForeignKey,
                            FKReference = c.FkReferencedInfo,
                            Identity = c.IsIdentity,
                            Default = c.Default,
                            Description = c.Description
                        }).ToList();

            dgvColumns.AutoResizeColumns();
        }


        public List<string> GenerateCodeCS(string ObjectName, string LinePrefix = "    ")
        {
            var Columns = _dbManager.DbTables.Where(t => t.TableName == ObjectName).First()
                        .Columns
                        .Select(c => new
                        {
                            Index = c.No,
                            ColumnName = c.ColumnName,
                            Type = c.FullDataType,
                            Nullable = c.IsNullable,
                            PK = c.IsPrimaryKey,
                            FK = c.IsForeignKey,
                            FKReference = c.FkReferencedInfo,
                            Identity = c.IsIdentity,
                            Default = c.Default,
                            Description = c.Description,
                            ColumnSize = c.Length,
                            NumericScale = c.NumericScale,
                            NumericPrecision = c.NumericPrecision

                        }).ToArray();



            var result = new List<string>();
            result.Add(string.Format("public class {0} {{", ObjectName));

            for (int i = 0, loopTo = Columns.Count() - 1; i <= loopTo; i++)
            {
                try
                {
                    /*
                     if (char.IsNumber(Convert.ToChar(Columns[i].ColumnName.Substring(0, 1))))
                     {   
                         Columns[i].ColumnName = "_" + Columns[i].ColumnName;
                     }
                     */
                    string AllowNull = ", null";
                    if (Columns[i].Nullable == false)
                        AllowNull = ", not null";
                    result.Add("/// <summary>");
                    /// <summary>
                    result.Add(string.Format("///   {0} ", Columns[i].Description).Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", ""));
                    /// 自動序號
                    /// </summary>
                    result.Add("/// </summary>");

                    //nvarchar 先另外處理  這三個

                    if (Columns[i].Type.IndexOf("nchar") >= 0 || Columns[i].Type.IndexOf("nvarchar") >= 0 || Columns[i].Type.IndexOf("varchar") >= 0)
                    {

                        result.Add(string.Format("{0}public string {1} {{ get; set; }} //(({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].Type, AllowNull));
                    }

                    switch (Columns[i].Type ?? "")
                    {

                        case "bigint":
                            {
                                result.Add(string.Format("{0}public long {1} {{ get; set; }} //(bigint{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "binary":
                            {
                                result.Add(string.Format("{0}public byte[] {1} {{ get; set; }} //(binary({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull));
                                break;
                            }

                        case "bit":
                            {
                                result.Add(string.Format("{0}public bool {1} {{ get; set; }} //(bit{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "char":
                            {
                                result.Add(string.Format("{0}public string {1} {{ get; set; }} //(char({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull));
                                break;
                            }

                        case "date":
                            {
                                result.Add(string.Format("{0}public DateTime {1} {{ get; set; }} //(date{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "datetime":
                            {
                                result.Add(string.Format("{0}public DateTime {1} {{ get; set; }} //(datetime{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "datetime2":
                            {
                                result.Add(string.Format("{0}public DateTime {1} {{ get; set; }} //(datetime2({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].NumericScale, AllowNull));
                                break;
                            }

                        case "datetimeoffset":
                            {
                                result.Add(string.Format("{0}public DateTimeOffset {1} {{ get; set; }} //(datetimeoffset{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "decimal":
                            {
                                result.Add(string.Format("{0}public decimal {1} {{ get; set; }} //(decimal({2},{3}){4})", LinePrefix, Columns[i].ColumnName, Columns[i].NumericPrecision, Columns[i].NumericScale, AllowNull));
                                break;
                            }

                        case "float":
                            {
                                result.Add(string.Format("{0}public double {1} {{ get; set; }} //(float{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "image":
                            {
                                result.Add(string.Format("{0}public byte[] {1} {{ get; set; }} //(image{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "int":
                            {
                                result.Add(string.Format("{0}public int {1} {{ get; set; }} //(int{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "smallint":
                            {
                                result.Add(string.Format("{0}public short {1} {{ get; set; }} //(smallint{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "money":
                            {
                                result.Add(string.Format("{0}public decimal {1} {{ get; set; }} //(money{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "nchar":
                            {
                                //if (Columns[i].IsLong)
                                //{
                                //    result.Add( string.Format("{0}public string {1} {{ get; set; }} //(nchar(max){2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                //}
                                //else
                                //{
                                result.Add(string.Format("{0}public string {1} {{ get; set; }} //(nchar({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull));
                                //}

                                break;
                            }

                        case "ntext":
                            {
                                result.Add(string.Format("{0}public string {1} {{ get; set; }} //(ntext{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "nvarchar":
                            {
                                //if (Columns[i].IsLong)
                                //{
                                //    result.Add( string.Format("{0}public string {1} {{ get; set; }} //(nvarchar(max){2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                //}
                                //else
                                //{
                                result.Add(string.Format("{0}public string {1} {{ get; set; }} //(nvarchar({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull));
                                //}

                                break;
                            }

                        case "real":
                            {
                                result.Add(string.Format("{0}public Single {1} {{ get; set; }} //(real({2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "smalldatetime":
                            {
                                result.Add(string.Format("{0}public DateTime {1} {{ get; set; }} //(smalldatetime{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "sql_variant":
                            {
                                result.Add(string.Format("{0}public object {1} {{ get; set; }} //(sql_variant{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "text":
                            {
                                result.Add(string.Format("{0}public string {1} {{ get; set; }} //(text{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "time":
                            {
                                result.Add(string.Format("{0}public DateTime {1} {{ get; set; }} //(time({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].NumericScale, AllowNull));
                                break;
                            }

                        case "timestamp":
                            {
                                result.Add(string.Format("{0}public byte[] {1} {{ get; set; }} //(timestamp{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "tinyint":
                            {
                                result.Add(string.Format("{0}public byte {1} {{ get; set; }} //(tinyint{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "uniqueidentifier":
                            {
                                result.Add(string.Format("{0}public Guid {1} {{ get; set; }} //(uniqueidentifier{2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                break;
                            }

                        case "varbinary":
                            {
                                //if (Columns[i].IsLong)
                                //{
                                //    result.Add( string.Format("{0}public byte[] {1} {{ get; set; }} //(varbinary(max){2})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull));
                                //}
                                //else
                                //{
                                result.Add(string.Format("{0}public byte[] {1} {{ get; set; }} //(varbinary({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, Columns[i].ColumnSize, AllowNull));
                                //}

                                break;
                            }

                        case "varchar":
                            {
                                //if (Columns[i].IsLong)
                                //{
                                //    result.Add( string.Format("{0}public string {1} {{ get; set; }} //(varchar(max){2})", LinePrefix, Columns[i].ColumnName, AllowNull));
                                //}
                                //else
                                //{
                                result.Add(string.Format("{0}public string {1} {{ get; set; }} //(varchar({2}){3})", LinePrefix, Columns[i].ColumnName, Columns[i].ColumnSize, AllowNull));
                                //}

                                break;
                            }

                        case "xml":
                            {
                                result.Add(string.Format("{0}public string {1} {{ get; set; }} //(XML(.){2})", LinePrefix, Columns[i].ColumnName, AllowNull)); // sql variant
                                break;
                            }

                        default:
                            {
                                switch (Columns[i].Type ?? "")
                                {

                                    case "Microsoft.SqlServer.Types.SqlGeography": // geography
                                        {
                                            result.Add(string.Format("{0}public Microsoft.SqlServer.Types.SqlGeography {1} {{ get; set; }} //({2}{3})", LinePrefix, Columns[i].ColumnName, Columns[i].Type, AllowNull));
                                            break;
                                        }

                                    case "Microsoft.SqlServer.Types.SqlHierarchyId": // heirarchyid
                                        {
                                            result.Add(string.Format("{0}public Microsoft.SqlServer.Types.SqlGeography {1} {{ get; set; }} //({2}{3})", LinePrefix, Columns[i].ColumnName, Columns[i].Type, AllowNull));
                                            break;
                                        }

                                    case "Microsoft.SqlServer.Types.SqlGeometry": // geometry
                                        {
                                            result.Add(string.Format("{0}public Microsoft.SqlServer.Types.SqlGeography {1} {{ get; set; }} //({2}{3})", LinePrefix, Columns[i].ColumnName, Columns[i].Type, AllowNull));
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
            result.Add("}");
            //   result[Columns.Count + 1] = "}";
            return result;
        }

        private static void InitRichTextBoxContextMenu(RichTextBox textBox)

        {

            //創建剪下子選單

            var cutMenuItem = new System.Windows.Forms.MenuItem("剪下");

            cutMenuItem.Click += (sender, eventArgs) => textBox.Cut();

            //創建複製子選單

            var copyMenuItem = new System.Windows.Forms.MenuItem("複製");

            copyMenuItem.Click += (sender, eventArgs) => textBox.Copy();

            //創建貼上子選單

            var pasteMenuItem = new System.Windows.Forms.MenuItem("貼上");

            pasteMenuItem.Click += (sender, eventArgs) => textBox.Paste();

            //創建右鍵選單並將子選單加入到右鍵選單中

            var contextMenu = new ContextMenu();

            contextMenu.MenuItems.Add(cutMenuItem);

            contextMenu.MenuItems.Add(copyMenuItem);

            contextMenu.MenuItems.Add(pasteMenuItem);

            textBox.ContextMenu = contextMenu;

        }


        private List<string> GetCheckedTableNames()
        {
            var checkedTables = new List<string>();
            foreach (var table in clbTables.CheckedItems.Cast<Table>())
            {
                checkedTables.Add(table.TableName);
            }

            return checkedTables;
        }


        // Events
        private void btnReload_Click(object sender, EventArgs e)
        {
            _dbManager.LoadTables();
            ShowTables(_dbManager.DbTables);
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {

            var checkedTableNames = GetCheckedTableNames();
            if (checkedTableNames.Count > 0)
            {
                try
                {
                    // generate db document
                    string templatePath = Config.GetInstance().Content.CurrentDocTemplatePath;
                    string outputPath = Config.GetInstance().Content.OutputDocLocation + Guid.NewGuid().ToString() + ".xlsx";
                    _dbManager.GenerateDocument(checkedTableNames, templatePath, outputPath);

                    // open it
                    Process.Start(outputPath);
                }
                catch (Exception ex)
                {
                    MsgBoxHelper.Error("Failure: " + ex.Message);
                }


            }
            else
            { MsgBoxHelper.Warning("Please select table!"); }


        }

        private void clbTables_SelectedValueChanged(object sender, EventArgs e)
        {
            var tableName = string.Empty;
            if (clbTables.SelectedIndex > -1)
            {
                tableName = (clbTables.Items[clbTables.SelectedIndex] as Table).TableName;
                ShowTableColumns(tableName);

                var POCO = GenerateCodeCS(tableName);
                txtPOCO.Text = "";
                foreach (var item in POCO)
                {
                    txtPOCO.AppendText("\r\n" + item);
                }

            }
        }

        private void btnCheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < clbTables.Items.Count; i++)
                clbTables.SetItemChecked(i, true);
        }

        private void btnUnCheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < clbTables.Items.Count; i++)
                clbTables.SetItemChecked(i, false);
        }

        private void btnRememberCheckedTables_Click(object sender, EventArgs e)
        {
            var checkedTables = GetCheckedTableNames();
            if (checkedTables.Count > 0)
            {
                try
                {
                    // save to config
                    var connName = Config.GetInstance().Content.CurrentConnectionName;
                    Config.GetInstance().Content.DocTablePackages[connName] = checkedTables;
                    Config.GetInstance().SaveConfig();

                    MsgBoxHelper.Done();
                }
                catch (Exception ex)
                {
                    MsgBoxHelper.Error("Failure: " + ex.Message);
                }
            }
            else
            { MsgBoxHelper.Warning("Please select table!"); }
        }

        private void btnSetting_Click(object sender, EventArgs e)
        {
            using (var frmSetting = new FrmSetting())
            {
                frmSetting.ShowDialog();
                if (frmSetting.DialogResult == DialogResult.OK)
                {
                    if (frmSetting.IsCurrentConnectionChanged)
                    {
                        // change connect string
                        _dbManager.LoadTables(Config.GetInstance().Content.CurrentConnection.Str);
                        ShowTables(_dbManager.DbTables);
                    }

                }
            }
        }

        private bool CheckIsNullable(string str)
        {
            return str == ", null";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            QueryForm QueryForm = new QueryForm();
            QueryForm.SetValue(Config.GetInstance().Content.CurrentConnection.Str);
            QueryForm.StartPosition = FormStartPosition.CenterParent;
            QueryForm.ShowDialog(this);
            /*
            switch (QueryForm.ShowDialog(this))
            {
                case DialogResult.Yes:
                    this.Show();
                    break;
                case DialogResult.No:
                    this.Show();

                    break;
                default:
                    break;
            }
            */
        }
    }
}
