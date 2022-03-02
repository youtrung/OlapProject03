using DevExpress.XtraBars;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.DashboardCommon;
using DevExpress.DataAccess.ConnectionParameters;
using Microsoft.SqlServer.Management.IntegrationServices;
using System.Data.SqlClient;
using Microsoft.AnalysisServices;

namespace App_Olap_Datawarehouse
{
    public partial class RibbonForm1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        String cnnString = "Data Source=DESKTOP-0G8H13M;Initial Catalog=msdb;" +
                "Integrated Security=True;";
        public RibbonForm1()
        {
            InitializeComponent();
            loadDatasource();
            dashboardDesigner1.Dashboard = createTabcontaniner();
            //   addItemstoPivot();
            createTabcontaniner();

        }
        public void loadDatasource()
        {

            #region #OLAPDataSource




            OlapConnectionParameters olapParams = new OlapConnectionParameters();
            olapParams.ConnectionString = @"provider=MSOLAP;
                                  data source=DESKTOP-0G8H13M;
                                  initial catalog=MultidimensionalQuanLiDT_DW;
                                  cube name=QLDT DW;";
            DashboardOlapDataSource olapDataSource = new DashboardOlapDataSource(olapParams);

            dashboardDesigner1.Dashboard.DataSources.Add(olapDataSource);

            #endregion #OLAPDataSource
        }

        Dashboard createTabcontaniner()
        {

            // lấy data từ dashboard UI
            IDashboardDataSource olapDataSource = dashboardDesigner1.Dashboard.DataSources[0];
            string connstr = olapDataSource.ToString();
            /// Phân tích theo cửa hàng và cửa hàng theo time

            Dashboard dashboard = new Dashboard();
            if (connstr.Equals("DevExpress.DashboardCommon.DashboardOlapDataSource"))
            {
                PivotDashboardItem pivoteCH = new PivotDashboardItem();
                pivoteCH.DataSource = olapDataSource;
                pivoteCH.Values.AddRange(new DevExpress.DashboardCommon.Measure("[Measures].[Doanh Thu]"), new DevExpress.DashboardCommon.Measure("[Measures].[SOLUONG]"));
                pivoteCH.Rows.Add(new DevExpress.DashboardCommon.Dimension("[DIM CUAHANG].[MACH].[MACH]"));
                pivoteCH.ComponentName = "pivoteCH";
                CardDashboardItem cardCH = new CardDashboardItem();
                cardCH.Name = "Phan theo cua hang";
                cardCH.ComponentName = "cardCH";
                cardCH.DataSource = olapDataSource;
                cardCH.SeriesDimensions.Add(new DevExpress.DashboardCommon.Dimension("[DIM CUAHANG].[MACH].[MACH]"));
                cardCH.SparklineArgument = new DevExpress.DashboardCommon.Dimension("[Time].[Date].[Date]", DateTimeGroupInterval.DayMonthYear);
                DevExpress.DashboardCommon.Measure actualValue_Doanhthu = new DevExpress.DashboardCommon.Measure("[Measures].[Doanh Thu]");
                actualValue_Doanhthu.NumericFormat.FormatType = DataItemNumericFormatType.Currency;
                //co the them target sales dongia , chua add vao cube 
                Card valuecardCH = new Card(actualValue_Doanhthu);
                cardCH.Cards.Add(valuecardCH);

                ChartDashboardItem chartKV = new ChartDashboardItem();
                chartKV.DataSource = olapDataSource;
                chartKV.Arguments.Add(new DevExpress.DashboardCommon.Dimension("[DIM KHACHHANG].[MAKV].[MAKV]"));
                chartKV.Panes.Add(new ChartPane());
                SimpleSeries Doanhthu = new SimpleSeries(SimpleSeriesType.Bar);
                SimpleSeries Soluong = new SimpleSeries(SimpleSeriesType.Bar);
                Soluong.Value = new DevExpress.DashboardCommon.Measure("[Measures].[SOLUONG]");
                Doanhthu.Value = new DevExpress.DashboardCommon.Measure("[Measures].[Doanh Thu]");
                chartKV.Panes[0].Series.AddRange(Doanhthu, Soluong);
                chartKV.Name = "Phan tich theo khu vuc";
                chartKV.ComponentName = "chartKV";

                PivotDashboardItem pivoteKV = new PivotDashboardItem();
                pivoteKV.DataSource = olapDataSource;
                pivoteKV.Values.AddRange(new DevExpress.DashboardCommon.Measure("[Measures].[Doanh Thu]"), new DevExpress.DashboardCommon.Measure("[Measures].[SOLUONG]"));
                pivoteKV.Rows.AddRange(new DevExpress.DashboardCommon.Dimension("[DIM KHACHHANG].[MAKV].[MAKV]"), new DevExpress.DashboardCommon.Dimension("[DIM KHACHHANG].[TEN TP].[TEN TP]"));
                pivoteKV.Columns.Add(new DevExpress.DashboardCommon.Dimension("[DIM KHACHHANG].[GIOITINH].[GIOITINH]"));
                pivoteKV.ComponentName = "pivoteKV";

                CardDashboardItem cardKV = new CardDashboardItem();
                cardKV.Name = "Phan  khu vuc theo thang";
                cardKV.ComponentName = "cardKV";
                cardKV.DataSource = olapDataSource;
                cardKV.SeriesDimensions.Add(new DevExpress.DashboardCommon.Dimension("[DIM KHACHHANG].[MAKV].[MAKV]"));
                cardKV.SparklineArgument = new DevExpress.DashboardCommon.Dimension("[Time].[Month].[Month]", DateTimeGroupInterval.DayMonthYear);
                //co the them target sales dongia , chua add vao cube 
                cardKV.Cards.Add(valuecardCH);







                dashboard.DataSources.Add(olapDataSource);
                dashboard.Items.AddRange(pivoteCH, cardCH, chartKV, pivoteKV, cardKV);
                TabContainerDashboardItem tabContain = new TabContainerDashboardItem();
                tabContain.ComponentName = "CuahangTab";
                tabContain.TabPages.Add(new DashboardTabPage() { Name = "Phân Theo Cửa hàng ", ComponentName = "page1" });
                tabContain.TabPages["page1"].AddRange(dashboard.Items["pivoteCH"], dashboard.Items["cardCH"]);





                DashboardTabPage secondTapePage = tabContain.CreateTabPage();
                secondTapePage.Name = "Phan tich theo khu vuc ";
                secondTapePage.AddRange(dashboard.Items["chartKV"], dashboard.Items["pivoteKV"], dashboard.Items["cardKV"]);
                secondTapePage.ShowItemAsTabPage = true;


                PivotDashboardItem pivote_mix = new PivotDashboardItem();
                pivote_mix.DataSource = olapDataSource;
                pivote_mix.Values.AddRange(new DevExpress.DashboardCommon.Measure("[Measures].[Doanh Thu]"));
                pivote_mix.Rows.AddRange(new DevExpress.DashboardCommon.Dimension("[DIM KHACHHANG].[KHUVUCKH].[MAKV]"),
                                        new DevExpress.DashboardCommon.Dimension("[DIM KHACHHANG].[MAKH].[MAKH]"),
                                        new DevExpress.DashboardCommon.Dimension("[DIM HANGHOA].[MALOAIHH].[MALOAIHH]"),
                                        new DevExpress.DashboardCommon.Dimension("[DIM HANGHOA].[MAHH].[MAHH]"));
                pivote_mix.ComponentName = "pivoteMIX";

                ChartDashboardItem chartMix = new ChartDashboardItem();
                chartMix.DataSource = olapDataSource;
                chartMix.Arguments.AddRange(new DevExpress.DashboardCommon.Dimension("[DIM HANGHOA].[MALOAIHH].[MALOAIHH]"), new DevExpress.DashboardCommon.Dimension("[DIM HANGHOA].[MAHH].[MAHH]"));
                chartMix.Panes.Add(new ChartPane());
                chartMix.SeriesDimensions.Add(new DevExpress.DashboardCommon.Dimension("[Time].[Week].[Week]"));
                chartMix.Panes[0].Series.AddRange(Doanhthu, Soluong); // doanh thu va so luong da co tren kia
                chartMix.Name = "Phan tich theo hang hoa theo Week";
                chartMix.ComponentName = "chartMix";

                PieDashboardItem pieMix = new PieDashboardItem();
                pieMix.DataSource = olapDataSource;
                pieMix.Arguments.AddRange(new DevExpress.DashboardCommon.Dimension("[DIM KHACHHANG].[MAKH].[MAKH]"), new DevExpress.DashboardCommon.Dimension("[DIM DICHVU].[MADV].[MADV]"));
                pieMix.SeriesDimensions.Add(new DevExpress.DashboardCommon.Dimension("[Time].[Date].[Date]"));
                pieMix.Values.AddRange(new DevExpress.DashboardCommon.Measure("[Measures].[Doanh Thu]"), new DevExpress.DashboardCommon.Measure("[Measures].[SOLUONG]"));
                pieMix.ComponentName = "pieMix";

                dashboard.Items.AddRange(pivote_mix, chartMix, pieMix);



                DashboardTabPage TapePage03 = tabContain.CreateTabPage();
                TapePage03.Name = "Phan tich hon hop";
                TapePage03.AddRange(dashboard.Items["pivoteMIX"], dashboard.Items["chartMix"], dashboard.Items["pieMix"]);
                TapePage03.ShowItemAsTabPage = true;

                dashboard.Items.Add(tabContain);
            }


            return dashboard;
            //tabCuahang.Dashboard.Items.Add(pivot);
            //  tabCuahang.TabPages["page1"].Add(dashboard.Items);
            //dashboardDesigner1.Dashboard.Items.Add(tabCuahang);

        }


        private void dashboardDesigner1_Load(object sender, EventArgs e)
        {

        }

        //load button
        private void barButtonItem1_ItemClick(object sender, ItemClickEventArgs e)
        {
            dashboardDesigner1.Dashboard.Items.Clear();
            dashboardDesigner1.Dashboard = createTabcontaniner();
        }

        private async void btnImportExcel_ItemClick(object sender, ItemClickEventArgs e)
        {
       
            SqlConnection DbConn = new SqlConnection(cnnString);
            SqlCommand ExecJob = new SqlCommand();
            ExecJob.CommandType = CommandType.StoredProcedure;
            ExecJob.CommandText = "msdb.dbo.sp_start_job";
            ExecJob.Parameters.AddWithValue("@job_name", "execPkgDW");//ben trong sql agent 
            ExecJob.Connection = DbConn; //assign the connection to the command.

            using (DbConn)
            {
                DbConn.Open();
                using (ExecJob)
                {
                   ExecJob.ExecuteNonQuery();
                }
            }
            await Task.Delay(3000);

            Server server = new Server();
            server.Connect("Data source=DESKTOP-0G8H13M;Timeout=7200000;Integrated Security=SSPI");
            Database database = server.Databases.FindByName("MultidimensionalQuanLiDT_DW");
            database.Process(ProcessType.ProcessFull);
            MessageBox.Show("Đổ dữ liệu từ EXCEL THÀNH CÔNG");
            dashboardDesigner1.Dashboard.Items.Clear();
            dashboardDesigner1.Dashboard = createTabcontaniner();

        }

        private void btnProcessCube_ItemClick(object sender, ItemClickEventArgs e)
        {
            Server server = new Server();
            server.Connect("Data source=DESKTOP-0G8H13M;Timeout=7200000;Integrated Security=SSPI");
            Database database = server.Databases.FindByName("MultidimensionalQuanLiDT_DW");

            database.Process(ProcessType.ProcessFull);
          
            MessageBox.Show("Đổ dữ liệu lên DW thành công");
        }

        private async void btnImportSql_ItemClick(object sender, ItemClickEventArgs e)
        {
            SqlConnection DbConn = new SqlConnection(cnnString);
            SqlCommand ExecJob = new SqlCommand();
            ExecJob.CommandType = CommandType.StoredProcedure;
            ExecJob.CommandText = "msdb.dbo.sp_start_job";
            ExecJob.Parameters.AddWithValue("@job_name", "execDWImportSQL");//ben trong sql agent 
            ExecJob.Connection = DbConn; //assign the connection to the command.

            using (DbConn)
            {
                DbConn.Open();
                using (ExecJob)
                {
                    ExecJob.ExecuteNonQuery();
                }
            }
            await Task.Delay(3000);

            Server server = new Server();
            server.Connect("Data source=DESKTOP-0G8H13M;Timeout=7200000;Integrated Security=SSPI");
            Database database = server.Databases.FindByName("MultidimensionalQuanLiDT_DW");
            database.Process(ProcessType.ProcessFull);
            MessageBox.Show("Đổ dữ liệu từ SQL THÀNH CÔNG");
            dashboardDesigner1.Dashboard.Items.Clear();
            dashboardDesigner1.Dashboard = createTabcontaniner();
        }

        private void barButtonItem3_ItemClick(object sender, ItemClickEventArgs e)
        {
            FormCustomQuery frm = new FormCustomQuery();
            frm.Show();
        }
    }
}