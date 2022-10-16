using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data.OleDb;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        System.Media.SoundPlayer player = new System.Media.SoundPlayer(@"01_Hand_of_Fate_II_The_Beginning_.wav");
        public static string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=telephones.mdb;";

        // поле - ссылка на экземпляр класса OleDbConnection для соединения с БД
        private OleDbConnection myConnection;
        //OleDbConnection con = new OleDbConnection();
        OleDbDataAdapter da = new OleDbDataAdapter();
        BindingSource bs = new BindingSource();
        DataSet ds = new DataSet();
        private void conn(OleDbConnection mycon)
        {
            try
            {
                string nazvanie = comboBox1.Text;
                da = new OleDbDataAdapter("select * from " + nazvanie, mycon);
                ds = new DataSet();
                da.Fill(ds);
                bs = new BindingSource(ds, ds.Tables[0].TableName);
                bindingNavigator1.BindingSource = bs;
                dgw1.DataSource = bs;
            }
            catch (Exception)//если в try возникнет исключение, обрабатываем его ниже в catch, к примеру выводим сообщение с текстом ошибки
            {
                MessageBox.Show("Нет подключения к базе данных");
            }
        }
        private void renn()
        {
            dgw1.Columns[0].HeaderText = "Номер записи";
            dgw1.Columns[1].HeaderText = "Наименование";
            dgw1.Columns[2].HeaderText = "Адрес";
            dgw1.Columns[3].HeaderText = "Email";
            dgw1.Columns[4].HeaderText = "ФИО";
            dgw1.Columns[5].HeaderText = "Должность";
            dgw1.Columns[6].HeaderText = "Телефон";
            dgw1.Columns[7].HeaderText = "Примечание";

        }
        public Form1()
        {
            InitializeComponent();
            myConnection = new OleDbConnection(connectString);

            // открываем соединение с БД
            myConnection.Open();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            dgw1.ReadOnly = true;
            comboBox1.Items.Clear();
            DataTable tbls = myConnection.GetSchema("Tables", new string[] { null, null, null, "TABLE" }); //список всех таблиц
            foreach (DataRow row in tbls.Rows)
            {
                string TableName = row["TABLE_NAME"].ToString();
                comboBox1.Items.Add(TableName);
            }
        }   

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        //private void toolStripButton3_Click(object sender, EventArgs e)
        //{
       
        //}

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            myConnection.Close();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            string nazvanie = comboBox1.Text;
 
            OleDbCommand command = new OleDbCommand("INSERT INTO " + nazvanie + " (naim,adres,email,fio,dolznost,tel,prim) VALUES (@na,@ad,@ema,@fi,@dol,@tel,@prim)", myConnection);
            command.Parameters.AddWithValue("@na",tb1.Text);
            command.Parameters.AddWithValue("@ad", tb2.Text);
            command.Parameters.AddWithValue("@ema", tb3.Text);
            command.Parameters.AddWithValue("@fi", tb4.Text);
            command.Parameters.AddWithValue("@dol", tb5.Text);
            command.Parameters.AddWithValue("@tel", tb6.Text);
            command.Parameters.AddWithValue("@prim", tb7.Text);
            command.ExecuteNonQuery();
            conn(myConnection);
            renn();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            string nazvanie = comboBox1.Text;           
            OleDbCommand command = new OleDbCommand("DELETE FROM "+nazvanie+" WHERE id_num=@id", myConnection);
            command.Parameters.AddWithValue("@id", dgw1.CurrentRow.Cells[0].Value);
            command.ExecuteNonQuery();
            conn(myConnection);
            renn();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {

            player.PlayLooping();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            player.Stop();
        }

       

        private void button2_Click(object sender, EventArgs e)
        {
            string nazvanie = textBox1.Text;
            OleDbCommand command = new OleDbCommand("CREATE TABLE " + nazvanie + " (Id_num INT IDENTITY (1, 1) PRIMARY KEY, naim NVARCHAR (50), adres NVARCHAR (50),email NVARCHAR (50), fio NVARCHAR (50), dolznost NVARCHAR (50),tel NVARCHAR (50),prim NVARCHAR (50)  )", myConnection);

            command.ExecuteNonQuery();
            comboBox1.Items.Clear();
            DataTable tbls = myConnection.GetSchema("Tables", new string[] { null, null, null, "TABLE" }); //список всех таблиц
            foreach (DataRow row in tbls.Rows)
            {
                string TableName = row["TABLE_NAME"].ToString();
                comboBox1.Items.Add(TableName);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            conn(myConnection);
            renn();

        }

        private void button3_Click(object sender, EventArgs e)
        {            
                for (int i = 0; i <dgw1.RowCount; i++)
                {
                   dgw1.Rows[i].Selected = false;
                    for (int j = 0; j < dgw1.ColumnCount; j++)
                        if (dgw1.Rows[i].Cells[j].Value != null)
                            if (dgw1.Rows[i].Cells[j].Value.ToString().Contains(tbStr.Text))
                            {
                               dgw1.Rows[i].Selected = true;
                                break;
                            }
                }
            }
        }
    }

