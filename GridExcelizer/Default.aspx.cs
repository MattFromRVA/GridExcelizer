using GridExcelizer.ExcelExport;
using System;
using System.Collections.Generic;
using System.Web.UI;

namespace GridExcelizer
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                PopulateGridView();
            }
           
            Page.Title = "GridExcelizer";
        }

        private void PopulateGridView()
        {
            var demoData = new List<dynamic>
        {
            new { TextData = "Sample Text", NumericData = 123, DateData = DateTime.Now, CurrencyData = 99.99m, PercentageData = "10.28%" },
            new { TextData = "Another Text", NumericData = 456, DateData = DateTime.Now.AddDays(-30), CurrencyData = 149.49m, PercentageData = "0.72%" },
        };

            GridView1.DataSource = demoData;
            GridView1.DataBind();
        }

        protected void ExportButton_Click(object sender, EventArgs e)
        {
            // Create an instance of your ExcelExports class
            GridToExcel exporter = new GridToExcel(Response);

            // Call the ExportToExcel method
            exporter.ExportToExcel(GridView1, "SampleExport");
        }
    }
}