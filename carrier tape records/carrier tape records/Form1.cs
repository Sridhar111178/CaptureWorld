using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;


namespace carrier_tape_records
{
    public partial class Form1 : Form
    {
        SqlConnection con;
        SqlCommand cmd;
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            con = new SqlConnection("Persist Security Info=False;User ID=EOLUSER1;Password=EOLUSER02;Initial Catalog=UTAC_EOL;Server=168.232.225.161\\EOLDBSVR02");
            cmd = new SqlCommand("select * from JSDOWNLOAD_SUMMARY", con);
         
        }
        private void button1_Click(object sender, EventArgs e)
        {
            cmd.Parameters.Clear();
            //command.Parameters.Add("ArrivedDate", SqlDbType.DateTime).Value = ArrivedDate;
            //command.Parameters.Add("ReturnDate", SqlDbType.DateTime).Value = ReturnDate;
            //cmd.Parameters.Add("@date1", SqlDbType.DateTime).Value = dateTimePicker1.Value.ToString();
            //cmd.Parameters.Add("@date2", SqlDbType.DateTime).Value = dateTimePicker2.Value.ToString();
            cmd.Parameters.AddWithValue("@date1", dateTimePicker1.Value.ToString());
            cmd.Parameters.AddWithValue("@date2", dateTimePicker2.Value.ToString());            
            cmd.Parameters.AddWithValue("@LotID", textBoxLotID.Text.ToString());
            cmd.Parameters.AddWithValue("@KeyCode", textBoxKeyCode.Text.ToString());
            cmd.Parameters.AddWithValue("@CarPartNo", textBoxCarPartNo.Text.ToString());
            cmd.Parameters.AddWithValue("@CarBatNo", textBoxCarBatNo.Text.ToString());
            cmd.Parameters.AddWithValue("@CovPartNo", textBoxCovPartNo.Text.ToString());
            cmd.Parameters.AddWithValue("@CovBatNo", textBoxCovBatNo.Text.ToString());

            con.Open();

            if(string.IsNullOrEmpty(textBoxLotID.Text) && string.IsNullOrEmpty(textBoxKeyCode.Text) && string.IsNullOrEmpty(textBoxCarPartNo.Text) && string.IsNullOrEmpty(textBoxCarBatNo.Text) && string.IsNullOrEmpty(textBoxCovPartNo.Text) && string.IsNullOrEmpty(textBoxCovBatNo.Text))
            {
                cmd.CommandText = "select MachineID,UserID,LotId,KeyCode,RequestDate,CarrierTapePartNo1,CarrierTapeBatch1,CoverTapePartNo1," +
           "CoverTapeBatch1,CarrierTapePartNo2,CarrierTapeBatch2,CoverTapePartNo2,CoverTapeBatch2," +
           "RequestStatus from JSDOWNLOAD_SUMMARY where RequestDate between @date1 And @date2 and PackType = 'Reel' " +
           "and RequestStatus = 'Success' and LotID like 'U%' ";
            }

            else if(!string.IsNullOrEmpty(textBoxLotID.Text))
            {
                cmd.CommandText = "select MachineID,UserID,LotId,KeyCode,RequestDate,CarrierTapePartNo1,CarrierTapeBatch1,CoverTapePartNo1," +
             "CoverTapeBatch1,CarrierTapePartNo2,CarrierTapeBatch2,CoverTapePartNo2,CoverTapeBatch2," +
             "RequestStatus from JSDOWNLOAD_SUMMARY where RequestDate between @date1 And @date2 and PackType = 'Reel' " +
             "and RequestStatus = 'Success' and LotID like 'U%' and LotId = @LotID ";
            }
            else if (!string.IsNullOrEmpty(textBoxKeyCode.Text))
            {
                cmd.CommandText = "select MachineID,UserID,LotId,KeyCode,RequestDate,CarrierTapePartNo1,CarrierTapeBatch1,CoverTapePartNo1," +
             "CoverTapeBatch1,CarrierTapePartNo2,CarrierTapeBatch2,CoverTapePartNo2,CoverTapeBatch2," +
             "RequestStatus from JSDOWNLOAD_SUMMARY where RequestDate between @date1 And @date2 and PackType = 'Reel' " +
             "and RequestStatus = 'Success' and LotID like 'U%' and KeyCode = @KeyCode ";
            }
            else if (!string.IsNullOrEmpty(textBoxCarPartNo.Text))
            {
                cmd.CommandText = "select MachineID,UserID,LotId,KeyCode,RequestDate,CarrierTapePartNo1,CarrierTapeBatch1,CoverTapePartNo1," +
             "CoverTapeBatch1,CarrierTapePartNo2,CarrierTapeBatch2,CoverTapePartNo2,CoverTapeBatch2," +
             "RequestStatus from JSDOWNLOAD_SUMMARY where RequestDate between @date1 And @date2 and PackType = 'Reel' " +
             "and RequestStatus = 'Success' and LotID like 'U%' and CarrierTapePartNo1 = @CarPartNo or CarrierTapePartNo2 = @CarPartNo ";
            }
            else if (!string.IsNullOrEmpty(textBoxCarBatNo.Text))
            {
                cmd.CommandText = "select MachineID,UserID,LotId,KeyCode,RequestDate,CarrierTapePartNo1,CarrierTapeBatch1,CoverTapePartNo1," +
             "CoverTapeBatch1,CarrierTapePartNo2,CarrierTapeBatch2,CoverTapePartNo2,CoverTapeBatch2," +
             "RequestStatus from JSDOWNLOAD_SUMMARY where RequestDate between @date1 And @date2 and PackType = 'Reel' " +
             "and RequestStatus = 'Success' and LotID like 'U%' and CarrierTapeBatch1 = @CarBatNo or CarrierTapeBatch2 = @CarBatNo ";
            }
            else if (!string.IsNullOrEmpty(textBoxCovPartNo.Text))
            {
                cmd.CommandText = "select MachineID,UserID,LotId,KeyCode,RequestDate,CarrierTapePartNo1,CarrierTapeBatch1,CoverTapePartNo1," +
             "CoverTapeBatch1,CarrierTapePartNo2,CarrierTapeBatch2,CoverTapePartNo2,CoverTapeBatch2," +
             "RequestStatus from JSDOWNLOAD_SUMMARY where RequestDate between @date1 And @date2 and PackType = 'Reel' " +
             "and RequestStatus = 'Success' and LotID like 'U%' and CoverTapePartNo1 = @CovPartNo or CoverTapePartNo2 = @CovPartNo ";
            }
            else if (!string.IsNullOrEmpty(textBoxCovBatNo.Text))
            {
                cmd.CommandText = "select MachineID,UserID,LotId,KeyCode,RequestDate,CarrierTapePartNo1,CarrierTapeBatch1,CoverTapePartNo1," +
             "CoverTapeBatch1,CarrierTapePartNo2,CarrierTapeBatch2,CoverTapePartNo2,CoverTapeBatch2," +
             "RequestStatus from JSDOWNLOAD_SUMMARY where RequestDate between @date1 And @date2 and PackType = 'Reel' " +
             "and RequestStatus = 'Success' and LotID like 'U%' and CoverTapeBatch1 = @CovBatNo or CoverTapeBatch2 = @CovBatNo ";
            }



            //cmd.CommandText = "select * from JSDOWNLOAD_SUMMARY where lotId = 'UI38D362.1' ";
            SqlDataAdapter adaptor = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            adaptor.Fill(dt);
            dataGridView1.DataSource = dt;
            con.Close();
       
        }
  
        private void button1_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            app.Visible = true;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                    else
                    {
                        worksheet.Cells[i + 2, j + 1] = "";
                    }
                }
            }

        }

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    dataGridView1.DataSource = null;
        //    dataGridView1.Rows.Clear();
        //}
    }
}
