using OfficeOpenXml;
using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;
//using OfficeOpenXml;

namespace ImportExcelandBindIntoGrid
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            
        }

        private void BindGrid()
        {
            if (!string.IsNullOrEmpty(ViewState["fileName"].ToString()))
            {
                string filePath = Server.MapPath("~/App_Data/" + ViewState["fileName"].ToString());
                DataTable dt = ImportExcelToDataTable(filePath);
                ViewState["DataTable"] = dt;
                grdExcel.DataSource = dt;
                grdExcel.DataBind();
            }
        }

        private DataTable ImportExcelToDataTable(string filePath)
        {
            DataTable dt = new DataTable();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                // Add columns to DataTable
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dt.Columns.Add(firstRowCell.Text);
                }

                // Add rows to DataTable
                for (int rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                {
                    var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                    DataRow newRow = dt.NewRow();

                    foreach (var cell in row)
                    {
                        newRow[cell.Start.Column - 1] = cell.Text;
                    }

                    dt.Rows.Add(newRow);
                }
            }

            return dt;
        }
        protected void btnUploadDataIntoGrid_Click(object sender, EventArgs e)
        {
            string path = Path.GetFileName(FileUpload1.FileName);
            ViewState["fileName"] = path;
            if (!string.IsNullOrEmpty(path))
            {
                path = path.Replace(" ", "");
                FileUpload1.SaveAs(Server.MapPath("~/App_Data/") + path);
                String ExcelPath = Server.MapPath("~/App_Data/") + path;
                string Conn = ConfigurationManager.ConnectionStrings["excelConnection"].ConnectionString;
                Conn = string.Format(Conn, ExcelPath, "yes");
                OleDbConnection mycon = new OleDbConnection(Conn);
                mycon.Open();

                DataTable excelData = mycon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string excelSheetName = excelData.Rows[0]["TABLE_NAME"].ToString();
                OleDbCommand cmd = new OleDbCommand("select * from [" + excelSheetName + "]", mycon);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                grdExcel.DataSource = dt;
                grdExcel.DataBind();
                mycon.Close();
                Response.Write("<script>alert('Grid Populated');</script>");
            }
            else
            {
                grdExcel.DataSource = null;
                grdExcel.DataBind();
                ViewState["fileName"] = null;
                Response.Write("<script>alert('Please upload Excel');</script>");

            }


        }

        protected void btnExport_Click(object sender, EventArgs e)
        {



            if (grdExcel.Rows.Count > 0)
            {
                string strFileName = ViewState["fileName"].ToString();

                string filePath = Server.MapPath("~/App_Data/" + strFileName);
                string fileName = Path.GetFileName(filePath);
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("Content-Disposition", $"attachment; filename={fileName}");
                Response.Clear();
                Response.WriteFile(filePath);
                Response.End();
            }
            else
            {
                Response.Write("<script>alert('Data Not Available');</script>");
            }

        }
        private void UpdateExcelFile(string filePath, DataTable dt)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                for (int row = 0; row < dt.Rows.Count; row++)
                {
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dt.Rows[row][col].ToString();
                    }
                }
                package.Save();
            }
        }
        protected void GridView1_RowEditing(object sender, GridViewEditEventArgs e)
        {
            grdExcel.EditIndex = e.NewEditIndex;
            BindGrid();
        }

        protected void GridView1_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            int index = e.RowIndex;
            GridViewRow row = grdExcel.Rows[index];
            string productName = ((TextBox)row.FindControl("txtProductName")).Text;
            string quantity = ((TextBox)row.FindControl("txtQuantity")).Text;
            string price = ((TextBox)row.FindControl("txtPrice")).Text;

            DataTable dt = (DataTable)ViewState["DataTable"];
            dt.Rows[index]["ProductName"] = productName;
            dt.Rows[index]["Quantity"] = quantity;
            dt.Rows[index]["Price"] = price;

            grdExcel.EditIndex = -1;

            string filePath = Server.MapPath("~/App_Data/" + ViewState["fileName"].ToString());
            UpdateExcelFile(filePath, dt);
            BindGrid();
        }

        protected void GridView1_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            grdExcel.EditIndex = -1;
            BindGrid();
        }

    }
}