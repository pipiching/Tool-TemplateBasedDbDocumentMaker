using DbDocumentMaker.Utility;
using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DbDocumentMaker
{
    public partial class QueryForm : Form
    {
        private string _ConnectStr = string.Empty;

        public QueryForm()
        {
            InitializeComponent();
        }

        public void SetValue(String ConnectStr)
        {
            this._ConnectStr = ConnectStr;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            string CodeText = string.Empty;
            var SchemaTable = new DataTable();

            try
            {
                GenerateCode(_ConnectStr, txtQuery.Text, txtObjectName.Text, ref SchemaTable, ref CodeText, null);
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message);
            }

            dgvSchemaColumns.DataSource = SchemaTable;
            txtCode.Text = CodeText;
        }

        private void GenerateCode(string ConnectionString, string Query, string ObjectName, ref DataTable SchemaTable, ref string Code, ArrayList spParms = null)
        {
            if (string.IsNullOrEmpty(ConnectionString))
                throw new Exception("Database Connection String is Required");

            if (string.IsNullOrEmpty(Query))
                throw new Exception("Query String is Required");

            if (string.IsNullOrEmpty(ObjectName))
                throw new Exception("Code Object Name is Required");

            var adoHelper = new ADOHelper();
            var Columns = adoHelper.GetFields(ConnectionString, Query, ref SchemaTable);

            if (!(Columns == null) && Columns.Count > 0)
            {
                string[] codeLines;
                    codeLines = adoHelper.GenerateCodeCS(ref Columns, ObjectName);
                Code = adoHelper.StringArrayToText(codeLines);
            }
        }

    }
}
