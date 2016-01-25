using System;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

namespace DataGridView_Import_Excel
{
    public partial class Form1 : Form
    {
        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=NO'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";

        public Form1()
        {
            InitializeComponent();
            //firstCombobox();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MessageBox.Show("Hello");
            //comboBox1.Items.Add("weekdays");
            //comboBox1.Items.Add("year");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //comboBox2.Items.Clear();
            //if (comboBox1.SelectedItem == "weekdays")
            //{
            //    comboBox2.Items.Add("Sunday");
            //    comboBox2.Items.Add("Monday");
            //    comboBox2.Items.Add("Tuesday");
            //}
            //else if (comboBox1.SelectedItem == "year")
            //{
            //    comboBox2.Items.Add("2012");
            //    comboBox2.Items.Add("2013");
            //    comboBox2.Items.Add("2014");
            //}
        }

        //public void firstCombobox()
        //{
        //    reportFolder.DisplayMember = "Text";
        //    reportFolder.ValueMember = "Value";

        //    var items = new[] { 
        //        new { Text = "Báo cáo doanh số", Value = "1" }, 
        //        new { Text = "Báo cáo giám sát", Value = "2" }, 
        //        new { Text = "Báo cáo kho", Value = "3" }
        //    };

        //    reportFolder.DataSource = items;
        //}

        //public void secondCombobox(int reportFolderID)
        //{
        //    reportType.DisplayMember = "Text";
        //    reportType.ValueMember = "Value";
        //    switch (reportFolderID)
        //    {
        //        case 1:
        //            var items = new[] { 
        //                new { Text = "Báo cáo sản lượng và doanh số bán hàng theo nvbh- khách hàng- sản phẩm", Value = "1" }, 
        //                new { Text = "Báo cáo sản lượng và doanh số bán hàng theo nvbh", Value = "2" }, 
        //                new { Text = "Báo cáo thực hiện chỉ tiêu kpi", Value = "3" }, 
        //                new { Text = "Báo cáo thực hiện kế hoạch tiêu thụ", Value = "4" }, 
        //                new { Text = "Báo cáo danh sách đơn hàng", Value = "5" }, 
        //                new { Text = "Bảng quyết toán chương trình khuyến mãi", Value = "6" }, 
        //                new { Text = "Báo cáo khách hàng phát sinh doanh số", Value = "7" }, 
        //                new { Text = "Báo cáo thực đặt thực giao", Value = "8" }, 
        //                new { Text = "Báo cáo doanh số", Value = "9" }, 
        //                new { Text = "Báo cáo giám sát", Value = "10" }, 
        //                new { Text = "Báo cáo kho", Value = "11" }
        //            };

        //            reportType.DataSource = items;

        //            break;
        //        case 2:
        //            break;
        //        case 3:
        //            break;
        //    }

        //}
        //private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        //{
        //    ComboBox comboBox = (ComboBox)sender;
        //    MessageBox.Show("123");
        //}

        //private void reportFolder_SelectionChangeCommitted(object sender, EventArgs e)
        //{
        //    int reportFolderID = Convert.ToInt32(reportFolder.SelectedValue);
        //    secondCombobox(reportFolderID);
            
        //}

        private void btnSelect_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        
        private void checkType()
        {
            //if (radioButton1.Checked == true)
            //{
            //    type = 0;
            //    return;
            //}
            //else if (radioButton2.Checked == true)
            //{
            //    type = 1;
            //    return;
            //}
            //else if (radioButton3.Checked == true)
            //{
            //    type = 2;
            //    return;
            //}

        }

        int type;

        int _flag_add_fist_to_data;
        
        private void ImportData()
        {
            type = Int32.Parse(textBox1.Text);

            _flag_add_fist_to_data = 0;

            string filePath = openFileDialog1.FileName;
            string filename = openFileDialog1.SafeFileName;
            string extension = Path.GetExtension(filePath);
            string conStr, sheetName;

            string connectionString = "data source=localhost; initial catalog=dmsreport; persist security info=True; Integrated Security=SSPI;";

            conStr = string.Empty;
            switch (extension)
            {
                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, "YES");
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, "YES");
                    break;
            }
            
            //Read Data 
            DataTable dt = new DataTable();
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        con.Open();
                        DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                        for (int i_sheet = 0; i_sheet < dtExcelSchema.Rows.Count; i_sheet++)
                        {
                            sheetName = dtExcelSchema.Rows[i_sheet].Field<string>("TABLE_NAME");
                            cmd.CommandText = "SELECT * From [" + sheetName + "]";
                            cmd.Connection = con;
                            oda.SelectCommand = cmd;
                            oda.Fill(dt);
                            checkType();
                            switch (type)
                            {
                                case 11: //SAN_LUONG_DOANH_SO_BAN_HANG_THEO_NVBH_KH_SP
                                    using (SqlConnection connection = new SqlConnection(connectionString))
                                    {
                                        DataColumn Col = dt.Columns.Add("", typeof(System.String));
                                        Col.SetOrdinal(6);

                                        foreach (DataRow row in dt.Rows)
                                        {
                                            int num;

                                            bool isNum = Int32.TryParse(row[7].ToString(), out num);

                                            if (!isNum)
                                            {
                                                row[6] = row[7];
                                                row[7] = "";
                                            }
                                        }

                                        for (int i = 4; i < dt.Rows.Count; i++)
                                        {
                                            DataRow row = dt.Rows[i];
                                            DataRow prevRow = dt.Rows[i - 1];
                                            string tprevRow = prevRow[6].ToString();
                                            if (row[6].ToString() == "")
                                            {
                                                row[6] = tprevRow;
                                            }
                                        }
                                        for (int i = 3; i < dt.Rows.Count; i++)
                                        {
                                            DataRow row = dt.Rows[i];
                                            if (row[dt.Columns.Count - 1].ToString() == "")
                                            {
                                                row.Delete();
                                                dt.AcceptChanges();
                                            }
                                        }
                                        connection.Open();
                                        try
                                        {
                                            string month_report = filename.Substring(4, 2);
                                            string year_report = filename.Substring(0, 4);

                                            SqlCommand cmdCheck = new SqlCommand("DELETE FROM [dbo].[SAN_LUONG_DOANH_SO_BAN_HANG_THEO_NVBH_KH_SP] WHERE month_report = " + month_report + " AND year_report = " + year_report);
                                            cmdCheck.Connection = connection;
                                            cmdCheck.ExecuteNonQuery();

                                            for (int i = 3; i < dt.Rows.Count; i++)
                                            {
                                                if (i == dt.Rows.Count)
                                                {
                                                    break;
                                                }
                                                SqlCommand cmdSql = new SqlCommand("INSERT INTO [dbo].[SAN_LUONG_DOANH_SO_BAN_HANG_THEO_NVBH_KH_SP] (code_sup, sup, code_saler, saler, code_customer, customer,  address, code_product, product, label, type_goods, type_goods_child, encapsulation, taste, num, price, total, date_add, month_report, year_report) VALUES ('" + dt.Rows[i][0] + "', N'" + dt.Rows[i][1] + "', '" + dt.Rows[i][2] + "', N'" + dt.Rows[i][3] + "', '" + dt.Rows[i][4] + "', N'" + dt.Rows[i][5] + "', N'" + dt.Rows[i][6] + "', '" + dt.Rows[i][7] + "', N'" + dt.Rows[i][8] + "', '" + dt.Rows[i][9] + "', '" + dt.Rows[i][10] + "', '" + dt.Rows[i][11] + "', '" + dt.Rows[i][12] + "', '" + dt.Rows[i][13] + "', '" + dt.Rows[i][14] + "', '" + dt.Rows[i][15] + "', '" + dt.Rows[i][16] + "', '" + DateTime.Now + "', '" + filename.Substring(4, 2) + "', '" + filename.Substring(0, 4) + "')");
                                                cmdSql.Connection = connection;
                                                cmdSql.ExecuteNonQuery();
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine(ex.Message);
                                        }
                                        finally
                                        {
                                            connection.Close();
                                        }

                                    }
                                    break;

                                case 12:
                                    using (SqlConnection connection = new SqlConnection(connectionString))
                                    {
                                        for (int i = 3; i < dt.Rows.Count; i++)
                                        {
                                            DataRow row = dt.Rows[i];
                                            if (row[7].ToString() == "")
                                            {
                                                row.Delete();
                                                dt.AcceptChanges();
                                            }
                                        }

                                        for (int i = dt.Rows.Count - 1; i > 3; i--)
                                        {
                                            DataRow row = dt.Rows[i];
                                            if (row[3].ToString() == "")
                                            {
                                                row.Delete();
                                                dt.AcceptChanges();
                                            }
                                        }

                                        string strTemp = dt.Rows[2][15].ToString();

                                        for (int i = 3; i < dt.Rows.Count; i++)
                                        {
                                            DataRow row = dt.Rows[i];
                                            row[15] = strTemp;                                            
                                        }

                                        connection.Open();
                                        try
                                        {
                                            string month_report = filename.Substring(4, 2);
                                            string year_report = filename.Substring(0, 4);

                                            SqlCommand cmdCheck = new SqlCommand("DELETE FROM [dbo].[DOANH_SO_SAN_LUONG_NVBH] WHERE month_report = " + month_report + " AND year_report = " + year_report);
                                            cmdCheck.Connection = connection;
                                            cmdCheck.ExecuteNonQuery();

                                            for (int i = 3; i < dt.Rows.Count; i++)
                                            {
                                                if (i == dt.Rows.Count)
                                                {
                                                    break;
                                                }
                                                SqlCommand cmdSql = new SqlCommand("INSERT INTO [dbo].[DOANH_SO_SAN_LUONG_NVBH] (code_sup, sup, code_saler, saler, code_place, place, code_product, product, convert_to_box, num, money_sale, total_weight, net_weight, volume, supplier, date_add, month_report, year_report) VALUES ('" + dt.Rows[i][1] + "', N'" + dt.Rows[i][2] + "', '" + dt.Rows[i][3] + "', N'" + dt.Rows[i][4] + "', '" + dt.Rows[i][5] + "', N'" + dt.Rows[i][6] + "', '" + dt.Rows[i][7] + "', N'" + dt.Rows[i][8] + "', '" + dt.Rows[i][9] + "', '" + dt.Rows[i][10] + "', '" + dt.Rows[i][11] + "', '" + dt.Rows[i][12] + "', '" + dt.Rows[i][13] + "', '" + dt.Rows[i][14] + "', N'" + dt.Rows[i][15] + "', '" + DateTime.Now + "', '" + filename.Substring(4, 2) + "', '" + filename.Substring(0, 4) + "')");
                                                cmdSql.Connection = connection;
                                                cmdSql.ExecuteNonQuery();
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine(ex.Message);
                                        }
                                        finally
                                        {
                                            connection.Close();
                                        }

                                    }
                                    break;
                                case 31:
                                    if (dt.Columns.Count == 31) // TONG HOP NXT NPP
                                    {
                                        using (SqlConnection connection = new SqlConnection(connectionString))
                                        {
                                            connection.Open();
                                            try
                                            {
                                                string month_report = filename.Substring(4, 2);
                                                string year_report = filename.Substring(0, 4);

                                                SqlCommand cmdCheck = new SqlCommand("DELETE FROM [dbo].[TONG_HOP_NXT_NPP] WHERE month_report = " + month_report + " AND year_report = " + year_report);
                                                cmdCheck.Connection = connection;
                                                cmdCheck.ExecuteNonQuery();

                                                for (int i = 6; i < dt.Rows.Count; i++)
                                                {
                                                    SqlCommand cmdSql = new SqlCommand("INSERT INTO TONG_HOP_NXT_NPP (code_stock, name_stock, code_supplier, name_supplier, date_from, date_to, code_product, name_product, type_goods, specification, unit, price, opening_stock, import_company, import_return_product_sale, import_return_product_promotion, import_configure, import_stock_vansale, import_total, export_sale, export_product_promotion, export_configure, export_vansale, export_vansale_product_sale, export_vansale_product_promotion, export_vansale_product_return, export_vansale_product_stock, export_total, closing_stock, total_money, date_add, month_report, year_report) VALUES ('" + dt.Rows[i][1] + "', N'" + dt.Rows[i][2] + "', '" + dt.Rows[i][3] + "', N'" + dt.Rows[i][4] + "', '" + dt.Rows[i][5] + "', '" + dt.Rows[i][6] + "', N'" + dt.Rows[i][7] + "', N'" + dt.Rows[i][8] + "', N'" + dt.Rows[i][9] + "', '" + dt.Rows[i][10] + "', N'" + dt.Rows[i][11] + "', '" + dt.Rows[i][12] + "', '" + dt.Rows[i][13] + "', '" + dt.Rows[i][14] + "', '" + dt.Rows[i][15] + "', '" + dt.Rows[i][16] + "', '" + dt.Rows[i][17] + "', '" + dt.Rows[i][18] + "', N'" + dt.Rows[i][19] + "', '" + dt.Rows[i][20] + "', N'" + dt.Rows[i][21] + "', '" + dt.Rows[i][22] + "', '" + dt.Rows[i][23] + "', '" + dt.Rows[i][24] + "', '" + dt.Rows[i][25] + "', '" + dt.Rows[i][26] + "', '" + dt.Rows[i][27] + "', '" + dt.Rows[i][28] + "', '" + dt.Rows[i][29] + "', '" + dt.Rows[i][30] + "', '" + DateTime.Now + "', '" + filename.Substring(4, 2) + "', '" + filename.Substring(0, 4) + "')");
                                                    cmdSql.Connection = connection;
                                                    cmdSql.ExecuteNonQuery();
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine(ex.Message);
                                            }
                                            finally
                                            {
                                                connection.Close();
                                            }

                                        }
                                    }
                                    else if (dt.Columns.Count == 14 ) // KHO CONG TY
                                    {
                                        using (SqlConnection connection = new SqlConnection(connectionString))
                                        {
                                            connection.Open();
                                            try
                                            {
                                                string month_report = filename.Substring(4, 2);
                                                string year_report = filename.Substring(0, 4);

                                                SqlCommand cmdCheck = new SqlCommand("DELETE FROM [dbo].[KHO_CONG_TY] WHERE month_report = " + month_report + " AND year_report = " + year_report);
                                                cmdCheck.Connection = connection;
                                                cmdCheck.ExecuteNonQuery();

                                                for (int i = 4; i < dt.Rows.Count; i++)
                                                {
                                                    if (i == dt.Rows.Count - 1)
                                                    {
                                                        break;
                                                    }
                                                    SqlCommand cmdSql = new SqlCommand("INSERT INTO KHO_CONG_TY (code_product, name_product, type_goods, specification, unit, price, opening_stock, import_company, import_total, export_company, export_total, closing_stock, total_money, date_add, month_report, year_report) VALUES ('" + dt.Rows[i][1] + "', N'" + dt.Rows[i][2] + "', N'" + dt.Rows[i][3] + "', '" + dt.Rows[i][4] + "', N'" + dt.Rows[i][5] + "', '" + dt.Rows[i][6] + "', '" + dt.Rows[i][7] + "', '" + dt.Rows[i][8] + "', '" + dt.Rows[i][9] + "', '" + dt.Rows[i][10] + "', '" + dt.Rows[i][11] + "', '" + dt.Rows[i][12] + "', '" + dt.Rows[i][13] + "', '" + DateTime.Now + "', '" + filename.Substring(4, 2) + "', '" + filename.Substring(0, 4) + "')");
                                                    cmdSql.Connection = connection;
                                                    cmdSql.ExecuteNonQuery();
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine(ex.Message);
                                            }
                                            finally
                                            {
                                                connection.Close();
                                            }

                                        }
                                    }
                                    else if (dt.Columns.Count == 25) // TONG HOP NPP & TUNG NHA PP 
                                    {
                                        if (sheetName.Contains("TỔNG HỢP"))
                                        {
                                            using (SqlConnection connection = new SqlConnection(connectionString))
                                            {
                                                connection.Open();
                                                try
                                                {
                                                    string month_report = filename.Substring(4, 2);
                                                    string year_report = filename.Substring(0, 4);

                                                    SqlCommand cmdCheck = new SqlCommand("DELETE FROM [dbo].[TONG_HOP_NPP] WHERE month_report = " + month_report + " AND year_report = " + year_report);
                                                    cmdCheck.Connection = connection;
                                                    cmdCheck.ExecuteNonQuery();

                                                    for (int i = 6; i < dt.Rows.Count - 1; i++)
                                                    {
                                                        if (i == dt.Rows.Count - 1)
                                                        {
                                                            break;
                                                        }
                                                        SqlCommand cmdSql = new SqlCommand("INSERT INTO TONG_HOP_NPP (code_product, name_product, type_goods, specification, unit, price, opening_stock, import_company, import_return_product_sale, import_return_product_promotion, import_configure, import_stock_vansale, import_total, export_sale, export_product_promotion, export_configure, export_vansale, export_vansale_product_sale, export_vansale_product_promotion, export_vansale_product_return, export_vansale_product_stock, export_total, closing_stock, total_money, date_add, month_report, year_report) VALUES ('" + dt.Rows[i][1] + "', N'" + dt.Rows[i][2] + "', N'" + dt.Rows[i][3] + "', N'" + dt.Rows[i][4] + "', '" + dt.Rows[i][5] + "', '" + dt.Rows[i][6] + "', '" + dt.Rows[i][7] + "', '" + dt.Rows[i][8] + "', '" + dt.Rows[i][9] + "', '" + dt.Rows[i][10] + "', '" + dt.Rows[i][11] + "', '" + dt.Rows[i][12] + "', '" + dt.Rows[i][13] + "', '" + dt.Rows[i][14] + "', '" + dt.Rows[i][15] + "', '" + dt.Rows[i][16] + "', '" + dt.Rows[i][17] + "', '" + dt.Rows[i][18] + "', '" + dt.Rows[i][19] + "', '" + dt.Rows[i][20] + "', '" + dt.Rows[i][21] + "', '" + dt.Rows[i][22] + "', '" + dt.Rows[i][23] + "', '" + dt.Rows[i][24] + "', '" + DateTime.Now + "', '" + filename.Substring(4, 2) + "', '" + filename.Substring(0, 4) + "')");
                                                        cmdSql.Connection = connection;
                                                        cmdSql.ExecuteNonQuery();
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine(ex.Message);
                                                }
                                                finally
                                                {
                                                    connection.Close();
                                                }

                                            }
                                        }
                                        else
                                        {
                                            for (int i = 6; i < dt.Rows.Count; i++ )
                                            {
                                                dt.Rows[i][0] = sheetName;
                                            }

                                            using (SqlConnection connection = new SqlConnection(connectionString))
                                            {
                                                connection.Open();
                                                try
                                                {
                                                    if (_flag_add_fist_to_data == 0)
                                                    {
                                                        string month_report = filename.Substring(4, 2);
                                                        string year_report = filename.Substring(0, 4);

                                                        SqlCommand cmdCheck = new SqlCommand("DELETE FROM [dbo].[TONG_HOP_TUNG_NPP] WHERE month_report = " + month_report + " AND year_report = " + year_report);
                                                        cmdCheck.Connection = connection;
                                                        cmdCheck.ExecuteNonQuery();
                                                        _flag_add_fist_to_data = 1;
                                                    }

                                                    for (int i = 6; i < dt.Rows.Count; i++)
                                                    {
                                                        if(i == dt.Rows.Count - 1)
                                                        {
                                                            break;
                                                        }
                                                        string _strTemp = dt.Rows[i][0].ToString();

                                                        var charsToRemove = new string[] { "$", "'" };
                                                        
                                                        foreach (var c in charsToRemove)
                                                        {
                                                            _strTemp = _strTemp.Replace(c, string.Empty);
                                                        }
                                                        SqlCommand cmdSql = new SqlCommand("INSERT INTO TONG_HOP_TUNG_NPP (name_supplier, code_product, name_product, type_goods, specification, unit, price, opening_stock, import_company, import_return_product_sale, import_return_product_promotion, import_configure, import_stock_vansale, import_total, export_sale, export_product_promotion, export_configure, export_vansale, export_vansale_product_sale, export_vansale_product_promotion, export_vansale_product_return, export_vansale_product_stock, export_total, closing_stock, total_money, date_add, month_report, year_report) VALUES ( N'" + _strTemp + "', '" + dt.Rows[i][1] + "', N'" + dt.Rows[i][2] + "', N'" + dt.Rows[i][3] + "', N'" + dt.Rows[i][4] + "', N'" + dt.Rows[i][5] + "', '" + dt.Rows[i][6] + "', '" + dt.Rows[i][7] + "', '" + dt.Rows[i][8] + "', '" + dt.Rows[i][9] + "', '" + dt.Rows[i][10] + "', '" + dt.Rows[i][11] + "', '" + dt.Rows[i][12] + "', '" + dt.Rows[i][13] + "', '" + dt.Rows[i][14] + "', '" + dt.Rows[i][15] + "', '" + dt.Rows[i][16] + "', '" + dt.Rows[i][17] + "', '" + dt.Rows[i][18] + "', '" + dt.Rows[i][19] + "', '" + dt.Rows[i][20] + "', '" + dt.Rows[i][21] + "', '" + dt.Rows[i][22] + "', '" + dt.Rows[i][23] + "', '" + dt.Rows[i][24] + "', '" + DateTime.Now + "', '" + filename.Substring(4, 2) + "', '" + filename.Substring(0, 4) + "')");
                                                        cmdSql.Connection = connection;
                                                        cmdSql.ExecuteNonQuery();
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine(ex.Message);
                                                }
                                                finally
                                                {
                                                    connection.Close();
                                                }

                                            }
                                        }
    
                                    }
                                    break;
                                default: MessageBox.Show("Please choose file type");
                                    break;
                            }
                            dt.Reset();
                            con.Close();
                        }
                    }
                }
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ImportData();
        }
    }
}
