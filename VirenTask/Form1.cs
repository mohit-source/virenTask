using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace VirenTask
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        void clear()
        {


            txtId.Text = "";
            txtname.Text = "";
            txtphone.Text = "";
            txtemail.Text = "";
            txtadd.Text = "";
            txtclg.Text = "";
        }
        private void btnsubmit_Click(object sender, EventArgs e)
        {

            try
            {
                System.Data.OleDb.OleDbConnection myconnection;
                System.Data.OleDb.OleDbCommand mycommand = new System.Data.OleDb.OleDbCommand();
                String sql = null;
                myconnection = new System.Data.OleDb.OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source='D:\\virentask.xlsx'; Extended Properties = Excel 8.0  ");
                myconnection.Open();

                mycommand.Connection = myconnection;
                sql = "Insert into [Sheet1$] (id,name,email,phone,address,college) values ('" + txtId.Text + "','" + txtname.Text + "','" + txtemail.Text + "','" + txtphone.Text + "','" + txtadd.Text + "','" + txtclg.Text + "') ";
                mycommand.CommandText = sql;
                mycommand.ExecuteNonQuery();
                     MessageBox.Show("Data Inserted");
                myconnection.Close();
                clear();

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }
        }
    }
}
