using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Data.OleDb;
using System.Diagnostics;
using System.Data.SqlClient;



namespace online_offline_Scanning
{
    public partial class Form1 : Form
    {
        UInt64 num = 0;
        UInt64 num2 = 0;        
        public Form1()
        {           
            InitializeComponent();
            //MessageBox.Show(MSSQLHelper.connectionString);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            dataGridView1.ReadOnly = true;
            label1.Text = "0";
            label5.Text = "WAIT";
            label11.Text = "0";
            label7.Text = "WAIT";
            label5.BackColor = Color.Orange;
            label7.BackColor = Color.Orange;
            Tip.Text = "Please scan the TVSN";
            label6.Text = "Please scan the TVSN";
            this.tabControl1.SelectedIndex = Convert.ToInt32(ConfigurationManager.AppSettings["Pagestate"]);
            saveFileDialog1.Filter = " Excel files(*.xls)|*.xls";
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["Pagestate"].Value = tabControl1.SelectedIndex.ToString();
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        // get code rules
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                numericUpDown1.Value = Convert.ToInt32(ConfigurationManager.AppSettings[comboBox1.SelectedItem.ToString()]);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        // set code rules
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings[comboBox1.SelectedItem.ToString()].Value = numericUpDown1.Value.ToString();
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
                MessageBox.Show("Save success");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                switch (listBox1.Items.Count)
                {
                    case 0:
                        {
                            if (textBox1.Text.Trim().Length != Convert.ToInt32(ConfigurationManager.AppSettings["TVSN"]))
                            {
                                MessageBox.Show("The length of the TVSN:  " +textBox1.Text.Trim()+ "  is not match!");
                            }
                            else
                            {
                                listBox1.Items.Add(textBox1.Text.Trim());
                                Tip.Text = "Please scan the Mainboard number";
                            }
                            textBox1.Text = "";
                            
                            break;
                        }
                    case 1:
                        {
                            if (textBox1.Text.Trim().Length != Convert.ToInt32(ConfigurationManager.AppSettings["Mainboard"]))
                            {
                                MessageBox.Show("The length of the Mainboard:  " + textBox1.Text.Trim() + "   is not match!");
                            }
                            else
                            {
                                listBox1.Items.Add(textBox1.Text.Trim());
                                Tip.Text = "Please scan the Panel number";
                            }
                            textBox1.Text = "";                         
                            break;
                        }
                    case 2:
                        {
                            if (textBox1.Text.Trim().Length != Convert.ToInt32(ConfigurationManager.AppSettings["Panel"]))
                            {
                                MessageBox.Show("The length of the Panel:  " + textBox1.Text.Trim() + "   is not match!");
                            }
                            else
                            {
                                listBox1.Items.Add(textBox1.Text.Trim());
                                try
                                {
                                    int n = MSSQLHelper.ExecuteSql("insert into OnlineScanning (TVSN,Mainboard,Panel)values( '" + listBox1.Items[0].ToString() + "','" + listBox1.Items[1].ToString() + "','" + listBox1.Items[2].ToString() + "')");
                                    if (n > 0)
                                    {
                                        if (listBox2.Items.Count >= 100)
                                            listBox2.Items.Clear();
                                        listBox2.Items.Add(listBox1.Items[0].ToString());
                                        listBox1.Items.Clear();
                                        num += 1;
                                        label1.Text = num.ToString();
                                        label5.Text = "PASS";
                                        label5.BackColor = Color.Green;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                    listBox1.Items.Clear();
                                    label5.Text = "FAIL";
                                    label5.BackColor = Color.Red;
                                    Tip.Text=ex.Message;
                                }
                                
                            }
                            textBox1.Text = "";
                            break;
                        }
                    default:
                        {
                            MessageBox.Show("Program error,please restart the software");
                            Tip.Text = "Please scan the TVSN";
                            break;
                        }
                }



            }
        }

        //Clear online data
        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            Tip.Text = "Please scan the TVSN";
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                switch (listBox4.Items.Count)
                {
                    case 0:
                        {
                            if (textBox2.Text.Trim().Length != Convert.ToInt32(ConfigurationManager.AppSettings["TVSN"]))
                            {
                                MessageBox.Show("The length of the TVSN:  " + textBox2.Text.Trim() + "  is not match!");
                            }
                            else
                            {
                                listBox4.Items.Add(textBox2.Text.Trim());
                                label6.Text = "Please scan the left stand";
                            }
                            textBox2.Text = "";

                            break;
                        }
                    case 1:
                        {
                            if (textBox2.Text.Trim().Length != Convert.ToInt32(ConfigurationManager.AppSettings["Left stand"]))
                            {
                                MessageBox.Show("The length of the left stand:  " + textBox2.Text.Trim() + "   is not match!");
                            }
                            else
                            {
                                listBox4.Items.Add(textBox2.Text.Trim());
                                label6.Text = "Please scan the right stand";
                            }
                            textBox2.Text = "";
                            break;
                        }
                    case 2:
                        {
                            if (textBox2.Text.Trim().Length != Convert.ToInt32(ConfigurationManager.AppSettings["Right stand"]))
                            {
                                MessageBox.Show("The length of the right stand:  " + textBox2.Text.Trim() + "   is not match!");
                            }
                            else
                            {
                                listBox4.Items.Add(textBox2.Text.Trim());
                                label6.Text = "Please scan the accessory bag";
                            }
                            textBox2.Text = "";
                            break;
                        }
                    case 3:
                        {
                            if (textBox2.Text.Trim().Length != Convert.ToInt32(ConfigurationManager.AppSettings["Accessory bag"]))
                            {
                                MessageBox.Show("The length of the accessory bag:  " + textBox2.Text.Trim() + "   is not match!");
                            }
                            else
                            {
                                listBox4.Items.Add(textBox2.Text.Trim());
                                try
                                {
                                    int onlineresult = Convert.ToUInt16(MSSQLHelper.GetSingle("select count(1) TVSN from OnlineScanning where TVSN='" + listBox4.Items[0].ToString() + "'"));
                                    int resetresult = Convert.ToUInt16(MSSQLHelper.GetSingle("select count(1) TVSN from XMDataControl where TVSN='" + listBox4.Items[0].ToString().Replace("/","") + "' and TestResult is not null"));
                                    if(onlineresult>0&&resetresult>0)
                                    {
                                        int n = MSSQLHelper.ExecuteSql("insert into OffLineScanning (TVSN,LeftStand,RightStand,Accessorybag)values( '" + listBox4.Items[0].ToString() + "','" + listBox4.Items[1].ToString() + "','" + listBox4.Items[2].ToString() + "','" + listBox4.Items[3].ToString() + "')");
                                        if (n > 0)
                                        {
                                            if (listBox3.Items.Count >= 100)
                                                listBox3.Items.Clear();
                                            listBox3.Items.Add(listBox4.Items[0].ToString());
                                            listBox4.Items.Clear();
                                            num2 += 1;
                                            label11.Text = num2.ToString();
                                            label7.Text = "PASS";
                                            label7.BackColor = Color.Green;
                                            label6.Text = "Please scan the TVSN";
                                        }
                                    }
                                    else if(onlineresult==0&&resetresult==0)
                                    {
                                        MessageBox.Show("Missing online scanning and factory reset");
                                        listBox4.Items.Clear();
                                        label7.Text = "FAIL";
                                        label7.BackColor = Color.Red;
                                        break;
                                    }
                                    else if(onlineresult==0)
                                    {
                                        MessageBox.Show("Missing online scanning");
                                        listBox4.Items.Clear();
                                        label7.Text = "FAIL";
                                        label7.BackColor = Color.Red;
                                        break;
                                    }
                                    else if (resetresult == 0)
                                    {
                                        MessageBox.Show("Missing factory reset");
                                        listBox4.Items.Clear();
                                        label7.Text = "FAIL";
                                        label7.BackColor = Color.Red;
                                        break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                    listBox4.Items.Clear();
                                    label7.Text = "FAIL";
                                    label7.BackColor = Color.Red;
                                    label6.Text = ex.Message;
                                }

                            }
                            textBox2.Text = "";
                            break;
                        }
                    default:
                        {
                            MessageBox.Show("Program error,please restart the software");
                            label6.Text = "Please scan the TVSN";
                            break;
                        }
                }



            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (this.tabControl1.SelectedIndex)
            {
                case 3:
                    {
                        validation validation1 = new validation();
                        validation1.ShowDialog();                       
                        break;
                    }
            }
        }

        //check data
        private void button5_Click(object sender, EventArgs e)
        {
            DataSet ds = null;
            switch(comboBox2.SelectedIndex)
            {
                case 0:
                    {
                        switch(comboBox3.SelectedIndex)
                        {
                            case 0:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where " + comboBox3.SelectedItem.ToString() + "='" + textBox3.Text.Trim() + "'");
                                    break;
                                }
                            case 1:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where TVSN in (select TVSN from XMCard where Batchnum='" + textBox3.Text.Trim() + "')");
                                    break;
                                }
                            case 2:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where " + comboBox3.SelectedItem.ToString() + "='" + textBox3.Text.Trim() + "'");
                                    break;
                                }
                            case 3:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where " + comboBox3.SelectedItem.ToString() + "='" + textBox3.Text.Trim() + "'");
                                    break;
                                }
                            default:
                                break;
                        }
                        break;
                    }
                case 1:
                    {
                        switch (comboBox3.SelectedIndex)
                        {
                            case 0:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where " + comboBox3.SelectedItem.ToString() + "='" + textBox3.Text.Trim() + "'");
                                    break;
                                }
                            case 1:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where TVSN in (select TVSN from XMCard where Batchnum='" + textBox3.Text.Trim() + "')");
                                    break;
                                }
                            case 4:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where " + comboBox3.SelectedItem.ToString() + "='" + textBox3.Text.Trim() + "'");
                                    break;
                                }
                            case 5:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where " + comboBox3.SelectedItem.ToString() + "='" + textBox3.Text.Trim() + "'");
                                    break;
                                }
                            case 6:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where " + comboBox3.SelectedItem.ToString() + "='" + textBox3.Text.Trim() + "'");
                                    break;
                                }
                            default:
                                break;
                        }
                        break;
                    }
                case 2:
                    {
                        ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where " + comboBox3.SelectedItem.ToString() + "='" + textBox3.Text.Trim() + "'");
                        break;
                    }
                case 7:
                    {
                        ds= MSSQLHelper.Query("select OffLineScanning.TVSN,XMDataControl.TVMN,XMDataControl.ChipSN,XMDataControl.ChipMN,XMDataControl.EthernetMac,XMDataControl.BTMac,XMDataControl.FWVersion,OnlineScanning.Mainboard,OnlineScanning.Panel,OffLineScanning.LeftStand,OffLineScanning.RightStand,OffLineScanning.Accessorybag,OnlineScanning.UpdateTime,OffLineScanning.UpdateTime from XMDataControl,OnlineScanning,OffLineScanning where OnlineScanning.TVSN=OffLineScanning.TVSN AND Replace(OnlineScanning.TVSN,'/','')=XMDataControl.TVSN AND OnlineScanning.TVSN in (select TVSN from XMCard where Batchnum ='"+textBox3.Text.Trim()+"')");
                        break;
                    }
                default:
                    {
                        switch (comboBox3.SelectedIndex)
                        {
                            case 0:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where " + comboBox3.SelectedItem.ToString() + "='" + textBox3.Text.Trim() + "'");
                                    break;
                                }
                            case 1:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where TVSN in (select replace(TVSN,'/','') from XMCard where Batchnum='" + textBox3.Text.Trim() + "')");
                                    break;
                                }
                            default:
                                {
                                    ds = MSSQLHelper.Query("select * from " + comboBox2.SelectedItem.ToString() + " where " + comboBox3.SelectedItem.ToString() + "='" + textBox3.Text.Trim() + "'");
                                    break;
                                }
                        }
                        break;
                    }
            }
            
            dataGridView1.DataSource = ds.Tables[0];
        }

        //data to Excel
        private void button6_Click(object sender, EventArgs e)
        {
            ToExcel.dataGVToExcel(dataGridView1);
            
        }

        //delete uploaded data 
        private void button8_Click(object sender, EventArgs e)
        {
            if (textBox4.Text.Trim().Length == 0)
            {
                MessageBox.Show("please import the batchnum");
            }
            else if (textBox4.Text.Trim().Contains("'") || textBox4.Text.Trim().Contains("="))
            {
                MessageBox.Show("Do not import ' or  =");
            }
            else
            {
                if (MSSQLHelper.ExecuteSql("Delete from XMCard where Batchnum ='" + textBox4.Text.Trim() + "'") > 0)
                    MessageBox.Show("Delete success");
                else
                    MessageBox.Show("Can't find the match data");
            }
        }

        //upload Excel data
        private void button7_Click(object sender, EventArgs e)
        {
            button7.Text = "Importing...";
            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "Excel 97-2003|*.xls|Excel file|*.xlsx";
            if (fd.ShowDialog() == DialogResult.OK)
            {
                TransferData(fd.FileName, "Sheet1", MSSQLHelper.connectionString);
            }
            button7.Text = "Import";
        }

        public void TransferData(string excelFile, string sheetName, string connectionString)
        {
            Stopwatch st = new System.Diagnostics.Stopwatch();
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + excelFile + ";" + "Extended Properties=Excel 12.0;";
            OleDbConnection conn = new OleDbConnection(strConn);
            SqlTransaction tran = null;
            SqlConnection sqlconn = null;
            try
            {
                st.Start();
                //获取全部数据  
                conn.Open();
                int BTMacnum = 0;
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                strExcel = string.Format("select * from [{0}$]", sheetName);
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                myCommand.Fill(ds, sheetName);
                /*ds.Tables[0].Columns.Add("UpdateTime", typeof(DateTime));
                var timer = GetDateTimeFromSQL();*/
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    if (ds.Tables[0].Rows[i]["BTMac"].ToString().Length != 0)
                        BTMacnum++;
                }
                if (BTMacnum != 0 && BTMacnum != ds.Tables[0].Rows.Count)
                    goto cc;
                ds.Tables[0].Columns["TVSN"].Unique = true;
                ds.Tables[0].Columns["TVSN"].AllowDBNull = false;
                ds.Tables[0].Columns["TVMN"].Unique = true;
                ds.Tables[0].Columns["TVMN"].AllowDBNull = false;
                ds.Tables[0].Columns["EthernetMac"].Unique = true;
                ds.Tables[0].Columns["EthernetMac"].AllowDBNull = false;
                if (BTMacnum == ds.Tables[0].Rows.Count)
                    ds.Tables[0].Columns["BTMac"].Unique = true;
                dt.Columns.Add("TVSN", typeof(String));
                dt.Columns.Add("AllMac", typeof(String));
                dt.Columns.Add("MACType", typeof(String));
                dt.Columns["AllMac"].Unique = true;
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    DataRow newRow = dt.NewRow();
                    newRow["TVSN"] = ds.Tables[0].Rows[i]["TVSN"];
                    newRow["AllMac"] = ds.Tables[0].Rows[i]["EthernetMac"];
                    newRow["MACType"] = "EthernetMac";
                    dt.Rows.Add(newRow);
                    if (BTMacnum == ds.Tables[0].Rows.Count)
                    {
                        DataRow newRow1 = dt.NewRow();
                        newRow1["TVSN"] = ds.Tables[0].Rows[i]["TVSN"];
                        newRow1["AllMac"] = ds.Tables[0].Rows[i]["BTMac"];
                        newRow1["MACType"] = "BTMac";
                        dt.Rows.Add(newRow1);
                    }
                }
                sqlconn = new SqlConnection(MSSQLHelper.connectionString);
                sqlconn.Open();
                using (tran = sqlconn.BeginTransaction())
                {
                    int num = ds.Tables[0].Rows.Count;
                    /*if (num >= 100)
                        progressBar1.Maximum = (ds.Tables[0].Rows.Count + dt.Rows.Count) / 100;
                    else
                        progressBar1.Maximum = ds.Tables[0].Rows.Count + dt.Rows.Count;*/
                    using (System.Data.SqlClient.SqlBulkCopy bcp1 = new System.Data.SqlClient.SqlBulkCopy(sqlconn, SqlBulkCopyOptions.Default, tran))
                    {
                        /*bcp1.SqlRowsCopied += new System.Data.SqlClient.SqlRowsCopiedEventHandler(bcp_SqlRowsCopied);
                        if (num >= 100)
                        {
                            bcp1.BatchSize = 100;//每次传输的行数
                            bcp1.NotifyAfter = 100;
                        }
                        else if (num >= 10 && num < 100)
                        {
                            bcp1.BatchSize = 10;//每次传输的行数
                            bcp1.NotifyAfter = 10;
                        }
                        else
                        {
                            bcp1.BatchSize = 1;//每次传输的行数
                            bcp1.NotifyAfter = 1;
                        }*/
                        bcp1.DestinationTableName = "UsingMac";//目标表
                        bcp1.ColumnMappings.Add("TVSN", "TVSN");
                        bcp1.ColumnMappings.Add("AllMac", "AllMac");
                        bcp1.ColumnMappings.Add("MACType", "MACType");
                        bcp1.WriteToServer(dt);
                        bcp1.ColumnMappings.Clear();
                        bcp1.DestinationTableName = "XMCard";

                        bcp1.ColumnMappings.Add("Batchnum", "Batchnum");
                        bcp1.ColumnMappings.Add("TVSN", "TVSN");
                        bcp1.ColumnMappings.Add("TVMN", "TVMN");
                        bcp1.ColumnMappings.Add("EthernetMac", "EthernetMac");
                        bcp1.ColumnMappings.Add("Model", "Model");
                        bcp1.ColumnMappings.Add("FWVersion", "FWVersion");
                        bcp1.ColumnMappings.Add("BTMac", "BTMac");
                        bcp1.WriteToServer(ds.Tables[0]);
                    }
                    tran.Commit();
                    st.Stop();
                    sqlconn.Close();
                    button1.Text = "Import";
                }
            cc:
                {
                    if (BTMacnum != 0 && BTMacnum != ds.Tables[0].Rows.Count)
                        MessageBox.Show("Missing BTMac!");
                    else
                        MessageBox.Show("Import success,using" + st.ElapsedMilliseconds.ToString() + " ms");
                }
                conn.Close();

            }
            catch (Exception ex)
            {
                st.Stop();
                conn.Close();
                if (tran != null)
                    tran.Rollback();
                if (sqlconn != null)
                    sqlconn.Close();
                button1.Text = "Import";
                MessageBox.Show(ex.Message.ToString());
            }

        }

        //clear offline data
        private void button4_Click(object sender, EventArgs e)
        {
            listBox4.Items.Clear();
            label6.Text = "Please scan the TVSN";
        }

        // delete online data
        private void button9_Click(object sender, EventArgs e)
        {
            if (textBox3.Text.Trim().Length == 0)
            {
                MessageBox.Show("please import the TVSN");
            }
            else if (textBox3.Text.Trim().Contains("'") || textBox3.Text.Trim().Contains("="))
            {
                MessageBox.Show("Do not import ' or  =");
            }
            else
            {
                if (MSSQLHelper.ExecuteSql("Delete from OnlineScanning where TVSN ='" + textBox3.Text.Trim() + "'") > 0)
                    MessageBox.Show("Delete success");
                else
                    MessageBox.Show("Can't find the match data");
                if(listBox2.Items.Contains(textBox3.Text.Trim()))
                {
                    listBox2.Items.Remove(textBox3.Text.Trim());
                    num -= 1;
                    label1.Text = num.ToString();
                }
            }
        }





    }
}
