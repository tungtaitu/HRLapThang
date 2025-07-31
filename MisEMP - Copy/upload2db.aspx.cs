using System;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Configuration;
using System.Data.SqlClient;
using System.Web.Configuration;

public partial class uplod2DB : System.Web.UI.Page
{
    string cs = WebConfigurationManager.ConnectionStrings["connstr"].ConnectionString.ToString();
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
            
            Import_To_Grid(FilePath, Extension, rbHDR.SelectedItem.Text);
            GetExcelSheets(FilePath, Extension, "Yes");
        }
    }

    private void Import_To_Grid(string FilePath, string Extension, string isHDR)
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
        }
		conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
        conStr = String.Format(conStr, FilePath, isHDR);
        OleDbConnection connExcel = new OleDbConnection(conStr);
        OleDbCommand cmdExcel = new OleDbCommand();
        OleDbDataAdapter oda = new OleDbDataAdapter();
        DataTable dt = new DataTable();
        cmdExcel.Connection = connExcel;

        //Get the name of First Sheet
        connExcel.Open();
        DataTable dtExcelSchema;
        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
        connExcel.Close();

        //Read Data from First Sheet
        connExcel.Open();
        cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
        oda.SelectCommand = cmdExcel;
        oda.Fill(dt);
        connExcel.Close();

        //Bind Data to GridView
        GridView1.Caption = Path.GetFileName(FilePath);
        GridView1.DataSource = dt;
        GridView1.DataBind();
				GridView1.Visible= false ;
    } 

    protected void PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
        string FileName = GridView1.Caption;
        string Extension = Path.GetExtension(FileName);
        string FilePath = Server.MapPath(FolderPath + FileName);

        Import_To_Grid(FilePath, Extension, rbHDR.SelectedItem.Text);
        GridView1.PageIndex = e.NewPageIndex;
        GridView1.DataBind();
    }
    private void GetExcelSheets(string FilePath, string Extension, string isHDR)
    {
        string conStr = "";
        switch (Extension)
        {
            case ".xls": //Excel 97-03
                //conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;  //20171019 change 
				conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                break;
            case ".xlsx": //Excel 07
                conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                break;
        }

        //Get the Sheets in Excel WorkBoo
        conStr = String.Format(conStr, FilePath, isHDR);
        OleDbConnection connExcel = new OleDbConnection(conStr);
        OleDbCommand cmdExcel = new OleDbCommand();
        OleDbDataAdapter oda = new OleDbDataAdapter();
        cmdExcel.Connection = connExcel;
        connExcel.Open();

        //Bind the Sheets to DropDownList
        ddlSheets.Items.Clear();
        //ddlSheets.Items.Add(new ListItem("--Select Sheet--", ""));
        ddlSheets.DataSource = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        ddlSheets.DataTextField = "TABLE_NAME";
        ddlSheets.DataValueField = "TABLE_NAME";
        ddlSheets.DataBind();
        connExcel.Close();
        txtTable.Text = "";
        lblFileName.Text = Path.GetFileName(FilePath) ;
				lblrelt.Text = " (Success OK !!)" ;
        Panel2.Visible = true;
        Panel1.Visible = false;

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
                //CommandText = "spx_ImportFromExcel03";
				CommandText = "spx_ImportFromExcel07";
                break;
            case ".xlsx": //Excel 07
                CommandText = "spx_ImportFromExcel07";
                break;
        }
        //Read Excel Sheet using Stored Procedure
        //And import the data into Database Table 
		// 20171019 change 
		//string sql = string.Format("exec spx_ImportFromExcel03 N'{0}','{1}','No','{2}' " , ddlSheets.SelectedItem.Text.Replace("'",""), FolderPath + FileName  , "AempXls2db" ) ; 
		string sql = string.Format("exec spx_ImportFromExcel07 N'{0}','{1}','No','{2}' " , ddlSheets.SelectedItem.Text.Replace("'",""), FolderPath + FileName  , "AempXls2db" ) ; 		
        String strConnString = cs; // ConfigurationManager.ConnectionStrings["conString"].ConnectionString;
        SqlConnection con = new SqlConnection(strConnString);
        SqlCommand cmd = new SqlCommand(sql,con);
        //cmd.CommandType = CommandType.StoredProcedure;
        cmd.CommandText = sql;
        //cmd.Parameters.Add("@SheetName", SqlDbType.VarChar).Value = ddlSheets.SelectedItem.Text;
        //cmd.Parameters.Add("@FilePath", SqlDbType.VarChar).Value = FolderPath + FileName;
        //cmd.Parameters.Add("@HDR", SqlDbType.VarChar).Value = rbHDR.SelectedItem.Text;
        //cmd.Parameters.Add("@TableName", SqlDbType.VarChar).Value = txtTable.Text; 
        //cmd.Connection = con; 
				
				//using (SqlCommand cmd = new SqlCommand(sql,conn))
        //       {
        //           
        //           cmd.CommandText = sql;
        //           cmd.ExecuteNonQuery();
        //
        //           cmd.Dispose();
        //       }
				
				
				
				//Response.Write (sql); 
				//Response.End();
        try
        {
            con.Open();
            //object count = cmd.ExecuteNonQuery(); 
						cmd.ExecuteNonQuery(); 
						int totalRowsCount = GridView1.Rows.Count; 
						lblMessage.ForeColor = System.Drawing.Color.Green;
						lblMessage.Text = totalRowsCount.ToString() + " records inserted. Success(OK)."; 
						
        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = ex.Message +"<br>"+ sql ;
        }
        finally
        {
            con.Close();
            con.Dispose();
            Panel1.Visible = true;
            Panel2.Visible = false;

        }
    }
    protected void btnCancel_Click(object sender, EventArgs e)
    {
        Panel1.Visible = true;
        Panel2.Visible = false;
    }
}
