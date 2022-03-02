using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace App_Olap_Datawarehouse
{
    public partial class FormCustomQuery : Form
    {
        public FormCustomQuery()
        {
            InitializeComponent();
        }

        private void btnThucthi_Click(object sender, EventArgs e)
        {
            string txt = txtArea.Text;
            DataSet ds = new DataSet();
            AdomdConnection myconn = new AdomdConnection(@"Provider=solap;data source=DESKTOP-0G8H13M; initial catalog =MultidimensionalQuanLiDT_DW"); ;
            AdomdDataAdapter mycmd = new AdomdDataAdapter();
            mycmd.SelectCommand = new AdomdCommand();
            mycmd.SelectCommand.Connection = myconn;
            mycmd.SelectCommand.CommandText = txt;
            myconn.Open();
            mycmd.Fill(ds, "src");
            myconn.Close();
            dataGridView1.DataSource = new DataView(ds.Tables[0]);
        }
    }
}
