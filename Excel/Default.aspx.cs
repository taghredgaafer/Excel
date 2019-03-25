using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace Excel
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void btnUpload_Click(object sender, EventArgs e)
        {

            if (FileUpload1.HasFile)

            {

                string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);

                string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName);

                string FolderPath = ConfigurationManager.AppSettings["FolderPath"];

                string FilePath = Server.MapPath(FolderPath + FileName);

                FileUpload1.SaveAs(FilePath);

                GetExcelSheets(FilePath, Extension, "Yes");

            }

        }

        private void GetExcelSheets(string FilePath, string Extension, string IsHDR)
        {
            string conStr = "";

            switch (Extension)

            {

                case ".xls": //Excel 97-03

                    conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;

                    break;

                case ".xlsx": //Excel 07

                    conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;

                    break;

                //case ".xlsx": //Excel 16

                //    conStr = ConfigurationManager.ConnectionStrings["Excel16ConString"].ConnectionString;

                //    break;

            }


            #region 

            //Get the Sheets in Excel WorkBoo

            conStr = String.Format(conStr, FilePath, IsHDR);

            OleDbConnection connExcel = new OleDbConnection(conStr);

            OleDbCommand cmdExcel = new OleDbCommand();

            OleDbDataAdapter oda = new OleDbDataAdapter();

            cmdExcel.Connection = connExcel;

            connExcel.Open();



            //Bind the Sheets to DropDownList

            ddlSheet.Items.Clear();

            ddlSheet.Items.Add(new ListItem("--Select Sheet--", ""));

            ddlSheet.DataSource = connExcel

                     .GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            ddlSheet.DataTextField = "TABLE_NAME";

            ddlSheet.DataValueField = "TABLE_NAME";

            ddlSheet.DataBind();

            connExcel.Close();

            txtTable.Text = "";

            lblFileName.Text = Path.GetFileName(FilePath);

            Panel2.Visible = true;

            Panel1.Visible = false;
            #endregion
        }

           
        protected void btnSave_Click(object sender, EventArgs e)

        {

            string FileName = lblFileName.Text;

            string Extension = Path.GetExtension(FileName);

            string FolderPath = Server.MapPath(ConfigurationManager.AppSettings["FolderPath"]);

            string CommandText = "";

            switch (Extension)

            {

                case ".xls": //Excel 97-03

                    CommandText = "spx_ImportFromExcel03";

                    break;

                case ".xlsx": //Excel 07

                    CommandText = "spx_ImportFromExcel07";

                    break;

                //case ".xlsx": //Excel 16

                //    CommandText = "spx_ImportFromExcel16";

                //    break;

            }
          


            #region

            //Read Excel Sheet using Stored Procedure

            //And import the data into Database Table

            String strConnString = ConfigurationManager.ConnectionStrings["conString"].ConnectionString;

            SqlConnection con = new SqlConnection(strConnString);

            SqlCommand cmd = new SqlCommand();

            cmd.CommandType = CommandType.StoredProcedure;

            cmd.CommandText = CommandText;

            cmd.Parameters.Add("@SheetName", SqlDbType.VarChar).Value =

                           ddlSheet.SelectedItem.Text;

            cmd.Parameters.Add("@FilePath", SqlDbType.VarChar).Value =

                           FolderPath + FileName;

            cmd.Parameters.Add("@HDR", SqlDbType.VarChar).Value =

                           rbHDR.SelectedItem.Text;

            cmd.Parameters.Add("@TableName", SqlDbType.VarChar).Value =

                           txtTable.Text;

            cmd.Connection = con;

            try

            {

                con.Open();

                object count = cmd.ExecuteNonQuery();

                lblMessage.ForeColor = System.Drawing.Color.Green;

                lblMessage.Text = count.ToString() + " records inserted.";

            }

            catch (Exception ex)

            {

                lblMessage.ForeColor = System.Drawing.Color.Red;

                lblMessage.Text = ex.Message;

            }

            finally

            {

                con.Close();

                con.Dispose();

                Panel1.Visible = true;

                Panel2.Visible = false;

            }

        }
        #endregion
        protected void btnCancel_Click(object sender, EventArgs e)
        {
            Response.Redirect("Default.aspx");


        }


    }


}