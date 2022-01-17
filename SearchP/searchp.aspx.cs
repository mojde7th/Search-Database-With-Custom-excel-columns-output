using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
namespace SearchP
{
    public partial class searchp : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            string mainconn = ConfigurationManager.ConnectionStrings["Myconnection"].ConnectionString;
            SqlConnection sqlconn = new SqlConnection(mainconn);
            SqlCommand sqlcomm = new SqlCommand("searchdate",sqlconn);
            sqlcomm.CommandType = CommandType.StoredProcedure;
            sqlcomm.Parameters.AddWithValue("@YYMM", TextBox1.Text);
            sqlconn.Open();
            GridView1.DataSource = sqlcomm.ExecuteReader();
            GridView1.DataBind();
            sqlconn.Close();

        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            Response.ClearContent();
            Response.AppendHeader("content-disposition", "attachment;filename=excel1.xls");
            Response.Charset = "";
            Response.ContentType = "application/excel";


            StringWriter stringWriter = new StringWriter();
            HtmlTextWriter htmlTextWriter = new HtmlTextWriter(stringWriter);
            GridView1.HeaderRow.Style.Add("background-color", "#FFFFFF");

            // Set background color of each cell of GridView1 header row
            foreach (TableCell tableCell in GridView1.HeaderRow.Cells)
            {
                tableCell.Style["background-color"] = "#A55129";
            }

            // Set background color of each cell of each data row of GridView1
            foreach (GridViewRow gridViewRow in GridView1.Rows)
            {
                gridViewRow.BackColor = System.Drawing.Color.White;
                foreach (TableCell gridViewRowTableCell in gridViewRow.Cells)
                {
                    gridViewRowTableCell.Style["background-color"] = "#FFF7E7";
                }
            }
            GridView1.RenderControl(htmlTextWriter);
            string style = @"<style> td { mso-number-format:\@;} </style>";
            Response.Write(style);
            Response.Write(stringWriter.ToString());
            Response.End();
        }
        public override void VerifyRenderingInServerForm(Control control)
        {

           
        }

  

        protected void Button3_Click1(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("RegNumber");
            dt.Columns.Add("DateM");
            dt.Columns.Add("6623");
            foreach (GridViewRow row in GridView1.Rows)
            {
                int i = row.RowIndex;
                DataRow dr = dt.NewRow();
                dr["RegNumber"] = GridView1.Rows[i].Cells[0].Text;
                dr["DateM"] = GridView1.Rows[i].Cells[1].Text;
                dr["6623"] = GridView1.Rows[i].Cells[2].Text;
                dt.Rows.Add(dr);
            }
            GridView gv = new GridView();
            gv.DataSource = dt;
            gv.DataBind();
            Response.ClearContent();
            Response.AppendHeader("content-disposition", "attachment;filename=excel1.xls");
            Response.ContentType = "application/excel";

            StringWriter stringWriter = new StringWriter();
            HtmlTextWriter htmlTextWriter = new HtmlTextWriter(stringWriter);
            GridView1.HeaderRow.Style.Add("background-color", "#FFFFFF");

            // Set background color of each cell of GridView1 header row
            foreach (TableCell tableCell in GridView1.HeaderRow.Cells)
            {
                tableCell.Style["background-color"] = "#A55129";
            }

            // Set background color of each cell of each data row of GridView1
            foreach (GridViewRow gridViewRow in GridView1.Rows)
            {
                gridViewRow.BackColor = System.Drawing.Color.White;
                foreach (TableCell gridViewRowTableCell in gridViewRow.Cells)
                {
                    gridViewRowTableCell.Style["background-color"] = "#FFF7E7";
                }
            }
            gv.RenderControl(htmlTextWriter);
            string style = @"<style> td { mso-number-format:\@;} </style>";
            Response.Write(style);
            Response.Write(stringWriter.ToString());
            Response.End();
        }
    }
}