using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Configuration;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

public partial class DynameicWriteSQL : System.Web.UI.Page
{
    public int select_color_count = 1;
    public string InputDate;
    private static string val = "";
    PurchaseDocuments CreatePD;
    private static int index;
    private static string total;
    //private static DataTable dtCurrentTable;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            Reset_Format();
            CreateDocuments();
            SetInitialRow();
            GVgetData();
        }
    }

    //Reset Format
    private void Reset_Format()
    {
        if (ViewState["CurrentTable"] != null)
        {
            savebtn.Enabled = true;
        }
        else
        {
            savebtn.Enabled = false;
        }
        listAddbtn.Enabled = true;  //新增Row按鈕Enabled
        savebtn.Enabled = false;   //儲存按鈕Enabled
        modify.Visible = false;     //修改按鈕Visible
        printOK.Visible = false;    //預覽列印按鈕Visible
        //Customer
        BM_Id.Text = "";
        BM_Name.Text = "";
        BM_Tel.Text = "";
        BM_cell.Text = "";
        BM_Fax.Text = "";
        BM_remark.Text = "";
        BM_address.Text = "";

        //Customer Print
        txt_name_print.Text = "";
        txt_Tel_print.Text = "";
        txt_cell_print.Text = "";
        txt_addr_print.Text = "";
        txt_FAX_print.Text = "";
        txt_remark_print.Text = "";

        ViewState["CurrentTable"] = null;
        DataTable Initial_dt = new DataTable("Initial_dt");
        Initial_dt = (DataTable)ViewState["CurrentTable"];
        Gridview1.DataSource = Initial_dt;
        Gridview1.DataBind();
        Gridview_print.DataSource = Initial_dt;
        Gridview_print.DataBind();

    }
    //Create Documents
    private void CreateDocuments()
    {
        CreatePD = new PurchaseDocuments();
        txt_CreateTime.Text = CreatePD.CreatDocument_date();
        txt_CreateTime_print.Text = CreatePD.CreatDocument_date();
        InputDate = string.Format("{0:yyyy/MM/dd}", txt_CreateTime.Text);
        txt_id.Text = CreatePD.CreatDocument_Id();
        txt_id_print.Text = CreatePD.CreatDocument_Id();

    }

    //新增列Btn
    protected void listAddbtn_Click(object sender, EventArgs e)
    {
        savebtn.Enabled = true;
        AddNewRowToGrid();
        if (ViewState["CurrentTable"] != null)
        {
            DataTable dtCurrentTable = (DataTable)ViewState["CurrentTable"];
            TextBox model = (TextBox)Gridview1.Rows[dtCurrentTable.Rows.Count - 1].Cells[1].FindControl("Model");
            model.Focus();
        }
    }
    //儲存Btn
    protected void savebtn_Click(object sender, EventArgs e)
    {
        txtMsg.Text = "產生表單如下";
        
        SaveToSQL();
        GVgetData();
        savebtn.Enabled = false;
        listAddbtn.Enabled = false;
        printOK.Visible = true;
        modify.Visible = true;
        
        myBMprint.Style["display"] = "block";
        myIMnoprint.Style["display"] = "none";
        myBMnoprint.Style["display"] = "none";

        Documents CreateDocuments = new Documents((DataTable)ViewState["CurrentTable"], BM_Id.Text, BM_Name.Text, BM_Tel.Text, BM_cell.Text, BM_address.Text, BM_Fax.Text, BM_remark.Text);
        Documents_PD.Controls.Add(CreateDocuments.Create_Document());
    }
    //捨棄Btn
    protected void delbtn_Click(object sender, EventArgs e)
    {
        Reset_Format();
    }
    //下一筆Btn
    protected void nextbtn_Click(object sender, EventArgs e)
    {
        CreateDocuments();
        Reset_Format();
    }
    //修改Btn
    protected void modify_Click(object sender, EventArgs e)
    {
        BM_Id.Text = txt_cid_print.Text;
        BM_Name.Text = txt_name_print.Text;
        BM_Tel.Text = txt_Tel_print.Text;
        BM_cell.Text = txt_cell_print.Text;
        BM_remark.Text = txt_remark_print.Text;
        BM_address.Text = txt_addr_print.Text;
        
        txtMsg.Text = "";
        listAddbtn.Enabled = true;
        savebtn.Enabled = true;
        modify.Visible = false;
        printOK.Visible = false;

        ListContent.Style["display"] = "block";
        myBMprint.Style["display"] = "none";
        myIMnoprint.Style["display"] = "block";
        myBMnoprint.Style["display"] = "block";

        if (HasId(txt_id.Text))
        {
            DelSqlRow(txt_id.Text);
        }
    }
    //DelSqlRow
    private void DelSqlRow(string cid)
    {
        SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["YiChangDocumentsConnectionString"].ConnectionString);
        SqlCommand comm = null;
        string execute = "DELETE FROM PurchaseDocuments WHERE p_cid=@p_cid";
        try
        {
            conn.Open();
            //Set SelectCommand And Parameters
            comm = new SqlCommand(execute, conn);
            //Create Select SQLcommand
            comm.Parameters.Add("@p_cid", SqlDbType.NVarChar).Value = cid;

            SqlDataReader dr = comm.ExecuteReader();
            dr.Dispose();
        }
        catch (Exception ex)
        {
            txtMsg.Text = ex.ToString();
        }
        conn.Close();
        conn.Dispose();
    }
    //判斷是否有資料
    private Boolean HasId(string cId_Exist)
    {
        SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["YiChangDocumentsConnectionString"].ConnectionString);
        conn.Open();
        SqlCommand cmd = new SqlCommand("Select Count(*) From PurchaseDocuments Where p_cid=@p_cid", conn);

        cmd.Parameters.Add("@p_cid", SqlDbType.NVarChar).Value = cId_Exist;

        SqlDataReader sdr = cmd.ExecuteReader();
        DataTable table = new DataTable();
        table.Load(sdr);
        string count = table.Rows[0][0].ToString();

        sdr.Dispose();
        cmd.Dispose();
        conn.Close();
        conn.Dispose();
        if (Int32.Parse(count) != 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    //預覽列印
    protected void printOK_Click(object sender, EventArgs e)
    {
        myBMprint.Style["display"] = "block";
        myIMnoprint.Style["display"] = "none";
        myBMnoprint.Style["display"] = "none";
        txtMsg.Text = "";
        modify.Visible = true;
        printOK.Visible = true;
        listAddbtn.Enabled = true;
        GVgetData();
    }
    //取得資料表
    private void GVgetData()
    {
        SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["YiChangDocumentsConnectionString"].ConnectionString);
        conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM PurchaseDocuments WHERE p_cid=@p_cid", conn);
        cmd.Parameters.Add("@p_cid", SqlDbType.NVarChar).Value = txt_id.Text;
        SqlDataReader sdr = cmd.ExecuteReader();
        DataTable table = new DataTable();
        table.Load(sdr);

        Gridview_print.DataSource = table.AsDataView();
        Gridview_print.DataBind();

        sdr.Dispose();
        cmd.Dispose();
        conn.Close();
        conn.Dispose();
    }

    //Row 初始
    private void SetInitialRow()
    {
        DataTable dt = new DataTable("myPDT");
        DataRow dr = null;

        dt.Columns.Add(new DataColumn("Id", typeof(string)));
        dt.Columns.Add(new DataColumn("Model", typeof(string)));
        dt.Columns.Add(new DataColumn("StorageNum", typeof(string)));
        dt.Columns.Add(new DataColumn("Color", typeof(string)));
        dt.Columns.Add(new DataColumn("Lenght", typeof(string)));
        dt.Columns.Add(new DataColumn("Weight", typeof(string)));
        dt.Columns.Add(new DataColumn("Number", typeof(string)));
        dt.Columns.Add(new DataColumn("Price", typeof(string)));
        dt.Columns.Add(new DataColumn("Total", typeof(string)));

        dr = dt.NewRow();
        dr["Id"] = 1;
        dr["Model"] = string.Empty;
        dr["StorageNum"] = string.Empty;
        dr["Color"] = string.Empty;
        dr["Lenght"] = string.Empty;
        dr["Weight"] = string.Empty;
        dr["Number"] = string.Empty;
        dr["Price"] = string.Empty;
        dr["Total"] = string.Empty;

        dt.Rows.Add(dr);

        //Store the DataTable in ViewState
        ViewState["CurrentTable"] = dt;

        Gridview1.DataSource = dt;
        Gridview1.DataBind();
    }
    //Row Add function
    private void AddNewRowToGrid()
    {
        int rowIndex = 0;
        if (ViewState["CurrentTable"] != null)
        {
            DataTable dtCurrentTable = (DataTable)ViewState["CurrentTable"];
            DataRow drCurrentRow = null;
            if (dtCurrentTable.Rows.Count > 0)
            {
                for (int i = 1; i <= dtCurrentTable.Rows.Count; i++)
                {
                    //extract the TextBox values
                    Label Id = (Label)Gridview1.Rows[rowIndex].Cells[1].FindControl("Id");
                    TextBox model = (TextBox)Gridview1.Rows[rowIndex].Cells[1].FindControl("Model");
                    TextBox Storage = (TextBox)Gridview1.Rows[rowIndex].Cells[2].FindControl("StorageNum");
                    TextBox Color = (TextBox)Gridview1.Rows[rowIndex].Cells[3].FindControl("Color");
                    TextBox Lenght = (TextBox)Gridview1.Rows[rowIndex].Cells[4].FindControl("Lenght");
                    TextBox Weight = (TextBox)Gridview1.Rows[rowIndex].Cells[5].FindControl("Weight");
                    TextBox Number = (TextBox)Gridview1.Rows[rowIndex].Cells[6].FindControl("Number");
                    TextBox Price = (TextBox)Gridview1.Rows[rowIndex].Cells[7].FindControl("Price");
                    Label Total = (Label)Gridview1.Rows[rowIndex].Cells[8].FindControl("Total");

                    drCurrentRow = dtCurrentTable.NewRow();
                    drCurrentRow["Id"] = i + 1;
                    //dtCurrentTable.Rows[i - 1]["Id"] = rowIndex + 1;
                    dtCurrentTable.Rows[i - 1]["Model"] = model.Text;
                    dtCurrentTable.Rows[i - 1]["StorageNum"] = Storage.Text;
                    dtCurrentTable.Rows[i - 1]["Color"] = Color.Text;
                    dtCurrentTable.Rows[i - 1]["Lenght"] = Lenght.Text;
                    dtCurrentTable.Rows[i - 1]["Weight"] = Weight.Text;
                    dtCurrentTable.Rows[i - 1]["Number"] = Number.Text;
                    dtCurrentTable.Rows[i - 1]["Price"] = Price.Text;
                    dtCurrentTable.Rows[i - 1]["Total"] = Total.Text;

                    rowIndex++;
                }
                dtCurrentTable.Rows.Add(drCurrentRow);
                Gridview1.DataSource = dtCurrentTable;
                Gridview1.DataBind();
                

            }
            else
            {
                SetInitialRow();
            }
        }
        else
        {
            SetInitialRow();
        }
    }
    //Deleting Row Btn
    protected void Gridview1_RowDeleting(object sender, GridViewDeleteEventArgs e)
    {
        
        int rowIndex = 0;
        int rowIndex_after = 0;
        DataTable dtCurrentTable = (DataTable)ViewState["CurrentTable"];
        if (ViewState["CurrentTable"] != null)
        {
            if (dtCurrentTable.Rows.Count > 0)
            {                
                for (int i = 1; i <= dtCurrentTable.Rows.Count; i++)
                {
                    //extract the TextBox values
                    Label Id = (Label)Gridview1.Rows[rowIndex].Cells[1].FindControl("Id");
                    TextBox model = (TextBox)Gridview1.Rows[rowIndex].Cells[1].FindControl("Model");
                    TextBox Storage = (TextBox)Gridview1.Rows[rowIndex].Cells[2].FindControl("StorageNum");
                    TextBox Color = (TextBox)Gridview1.Rows[rowIndex].Cells[3].FindControl("Color");
                    TextBox Lenght = (TextBox)Gridview1.Rows[rowIndex].Cells[4].FindControl("Lenght");
                    TextBox Weight = (TextBox)Gridview1.Rows[rowIndex].Cells[5].FindControl("Weight");
                    TextBox Number = (TextBox)Gridview1.Rows[rowIndex].Cells[6].FindControl("Number");
                    TextBox Price = (TextBox)Gridview1.Rows[rowIndex].Cells[7].FindControl("Price");
                    Label Total = (Label)Gridview1.Rows[rowIndex].Cells[8].FindControl("Total");

                    dtCurrentTable.Rows[i - 1]["Id"] = rowIndex + 1;
                    dtCurrentTable.Rows[i - 1]["Model"] = model.Text;
                    dtCurrentTable.Rows[i - 1]["StorageNum"] = Storage.Text;
                    dtCurrentTable.Rows[i - 1]["Color"] = Color.Text;
                    dtCurrentTable.Rows[i - 1]["Lenght"] = Lenght.Text;
                    dtCurrentTable.Rows[i - 1]["Weight"] = Weight.Text;
                    dtCurrentTable.Rows[i - 1]["Number"] = Number.Text;
                    dtCurrentTable.Rows[i - 1]["Price"] = Price.Text;
                    dtCurrentTable.Rows[i - 1]["Total"] = Total.Text;

                    rowIndex++;
                }
                dtCurrentTable.Rows.Remove(dtCurrentTable.Rows[e.RowIndex]);
                for (int i = 1; i <= dtCurrentTable.Rows.Count; i++)
                {
                    Label Id = (Label)Gridview1.Rows[rowIndex_after].Cells[1].FindControl("Id");
                    dtCurrentTable.Rows[i - 1]["Id"] = rowIndex_after + 1;
                    rowIndex_after++;
                }
                Gridview1.DataSource = dtCurrentTable;
                Gridview1.DataBind();
            }
            else
            {
                SetInitialRow();
            }
        }
        else
        {
            SetInitialRow();
        }
    }
    //Save Row Btn
    protected void saveGvdate()
    {
        int rowIndex = 0;
        DataTable dtx = new DataTable();
        //DataTable dtx = (DataTable)ViewState["CurrentTable"];
        if (ViewState["CurrentTable"] != null)
        {
            for (int i = 1; i <= Gridview1.Rows.Count; i++)
            {
                //extract the TextBox values
                Label Id = (Label)Gridview1.Rows[rowIndex].Cells[1].FindControl("Id");
                TextBox model = (TextBox)Gridview1.Rows[rowIndex].Cells[1].FindControl("Model");
                TextBox Storage = (TextBox)Gridview1.Rows[rowIndex].Cells[2].FindControl("StorageNum");
                TextBox Color = (TextBox)Gridview1.Rows[rowIndex].Cells[3].FindControl("Color");
                TextBox Lenght = (TextBox)Gridview1.Rows[rowIndex].Cells[4].FindControl("Lenght");
                TextBox Weight = (TextBox)Gridview1.Rows[rowIndex].Cells[5].FindControl("Weight");
                TextBox Number = (TextBox)Gridview1.Rows[rowIndex].Cells[6].FindControl("Number");
                TextBox Price = (TextBox)Gridview1.Rows[rowIndex].Cells[7].FindControl("Price");
                Label Total = (Label)Gridview1.Rows[rowIndex].Cells[8].FindControl("Total");
                if (!string.IsNullOrEmpty(model.Text))
                {
                    if (string.IsNullOrEmpty(Number.Text))
                    {
                        Number.Text = "0";
                    }
                    if (string.IsNullOrEmpty(Price.Text))
                    {
                        Price.Text = "0";
                    }
                    dtx.Rows[i - 1]["Id"] = Id.Text;
                    dtx.Rows[i - 1]["Model"] = model.Text;
                    dtx.Rows[i - 1]["StorageNum"] = Storage.Text;
                    dtx.Rows[i - 1]["Color"] = Color.Text;
                    dtx.Rows[i - 1]["Lenght"] = Lenght.Text;
                    dtx.Rows[i - 1]["Weight"] = Weight.Text;
                    dtx.Rows[i - 1]["Number"] = Number.Text;
                    dtx.Rows[i - 1]["Price"] = Price.Text;
                    dtx.Rows[i - 1]["Total"] = Convert.ToInt32(Number.Text) * Convert.ToDouble(Price.Text);

                    rowIndex++;
                }
                else
                {
                    
                    rowIndex++;
                }
            }
            
            GV_Display.DataSource = dtx;
            GV_Display.DataBind();
        }
    }
    //RowDataBound
    protected void Gridview1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
           // Button delbtn = (Button)e.Row.FindControl("Delete");
            TextBox tbmodel = (TextBox)e.Row.FindControl("Model");
            TextBox tbcolor = (TextBox)e.Row.FindControl("Color");
            TextBox tblenght = (TextBox)e.Row.FindControl("Lenght");
            TextBox tbweight = (TextBox)e.Row.FindControl("Weight");
            TextBox tbnumber = (TextBox)e.Row.FindControl("Number");
            TextBox tbprice = (TextBox)e.Row.FindControl("Price");
            //delbtn.OnClientClick = "return confirm('確認要刪除嗎？');";
            //Label total = (Label)e.Row.FindControl("Total");
            
            tbmodel.Attributes.Add("onKeyDown", "if (event.keyCode==13) { document.getElementById('" + e.Row.FindControl("Color").ClientID + "').focus(); return false;}");
            tbcolor.Attributes.Add("onKeyDown", "if (event.keyCode==13) { document.getElementById('" + e.Row.FindControl("Lenght").ClientID + "').focus(); return false;}");
            tblenght.Attributes.Add("onKeyDown", "if (event.keyCode==13) { document.getElementById('" + e.Row.FindControl("Weight").ClientID + "').focus(); return false;}");
            tbweight.Attributes.Add("onKeyDown", "if (event.keyCode==13) { document.getElementById('" + e.Row.FindControl("Number").ClientID + "').focus(); return false;}");
            tbnumber.Attributes.Add("onKeyDown", "if (event.keyCode==13) { document.getElementById('" + listAddbtn.ClientID + "').focus();return true;}");
            tbprice.Attributes.Add("onKeyDown", "if (event.keyCode==13) { document.getElementById('" + listAddbtn.ClientID + "').focus();return true;}");
            var model = e.Row.FindControl("Model") as TextBox;
            var storage = e.Row.FindControl("StorageNum") as TextBox;
            var price = e.Row.FindControl("Price") as TextBox;
            var color = e.Row.FindControl("color") as TextBox;
            var weight = e.Row.FindControl("weight") as TextBox;
            var btn = e.Row.FindControl("Findbtn") as Button;
            var findcolor = e.Row.FindControl("Findcolor") as Button;

            if (model != null)
            {
                var domId = model.ClientID;
                var domstorage = storage.ClientID;
                var domcolor = color.ClientID;
                var domweight = weight.ClientID;
                var domprice = price.ClientID;
                
                if (btn != null)
                {
                    var js = string.Format(@"window.open('Subwindow_IM.aspx?id={0}&storage={1}&weight={2}&price={3}', '選擇產品', 'scrollbars=yes,top=100,left=200,width=800,height=600'); return false;", domId, domstorage, domweight, domprice);

                    btn.OnClientClick = js;
                }
            }
            if (color != null)
            {
                var domId_color = color.ClientID;

                if (findcolor != null)
                {
                    var js = string.Format(@"window.open('SelectColor.aspx?id={0}', '選擇顏色', 'scrollbars=yes,top=100,left=200,width=800,height=600'); return false;", domId_color);
                    findcolor.OnClientClick = js;
                }
            }
        }
    }

    //寫進SQL
    private void SaveToSQL()
    {
        txt_cid_print.Text = BM_Id.Text;
        txt_name_print.Text = BM_Name.Text;
        txt_Tel_print.Text = BM_Tel.Text;
        txt_cell_print.Text = BM_cell.Text;
        txt_remark_print.Text = BM_remark.Text;
        txt_addr_print.Text = BM_address.Text;
        try
        {
            int rowIndex = 0;
            SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["YiChangDocumentsConnectionString"].ConnectionString);
            conn.Open();
            if (ViewState["CurrentTable"] != null)
            {
                DataTable dt = (DataTable)ViewState["CurrentTable"];
                //for(int i = 1; i <= Gridview1.Rows.Count; i++)
                foreach (GridViewRow row in Gridview1.Rows)
                {
                    Label Id = (Label)Gridview1.Rows[rowIndex].Cells[1].FindControl("Id");
                    TextBox model = (TextBox)Gridview1.Rows[rowIndex].Cells[1].FindControl("Model");
                    TextBox Storage = (TextBox)Gridview1.Rows[rowIndex].Cells[2].FindControl("StorageNum");
                    TextBox Color = (TextBox)Gridview1.Rows[rowIndex].Cells[3].FindControl("Color");
                    TextBox Lenght = (TextBox)Gridview1.Rows[rowIndex].Cells[4].FindControl("Lenght");
                    TextBox Weight = (TextBox)Gridview1.Rows[rowIndex].Cells[5].FindControl("Weight");
                    TextBox Number = (TextBox)Gridview1.Rows[rowIndex].Cells[6].FindControl("Number");
                    TextBox Price = (TextBox)Gridview1.Rows[rowIndex].Cells[7].FindControl("Price");
                    Label Total = (Label)Gridview1.Rows[rowIndex].Cells[8].FindControl("Total");

                    if (!string.IsNullOrEmpty(model.Text))
                    {
                        getId();
                        //Set SelectCommand And Parameters
                        SqlCommand comm = new SqlCommand("Insert INTO PurchaseDocuments (id, p_id, p_cid, p_pdate, p_cno, p_cname, p_ctel, p_ccell, p_cfax, p_caddr, p_cremark, p_pmodel, p_pimage, p_pcolor, p_plen, p_pweight, p_pnum, p_pprice, p_ptotal, p_pstatus, p_ppic, p_pdriver) VALUES (@id, @p_id, @p_cid, @p_pdate, @p_cno, @p_cname, @p_ctel, @p_ccell, @p_cfax, @p_caddr, @p_cremark, @p_pmodel, @p_pimage, @p_pcolor, @p_plen, @p_pweight, @p_pnum, @p_pprice, @p_ptotal, @p_pstatus, @p_ppic, @p_pdriver)", conn);
                        //Create Select SQLcommand
                        comm.Parameters.Add("@id", SqlDbType.Int).Value = Int32.Parse(val);
                        comm.Parameters.Add("@p_id", SqlDbType.Int).Value = Id.Text;
                        comm.Parameters.Add("@p_cid", SqlDbType.NVarChar).Value = txt_id.Text;
                        comm.Parameters.Add("@p_pdate", SqlDbType.Date).Value = txt_CreateTime.Text;
                        comm.Parameters.Add("@p_cno", SqlDbType.NVarChar).Value = BM_Id.Text;
                        comm.Parameters.Add("@p_cname", SqlDbType.NVarChar).Value = BM_Name.Text;
                        comm.Parameters.Add("@p_ctel", SqlDbType.NVarChar).Value = BM_Tel.Text;
                        comm.Parameters.Add("@p_ccell", SqlDbType.NVarChar).Value = BM_cell.Text;
                        comm.Parameters.Add("@p_cfax", SqlDbType.NVarChar).Value = BM_Fax.Text;
                        comm.Parameters.Add("@p_cremark", SqlDbType.NVarChar).Value = BM_remark.Text;
                        comm.Parameters.Add("@p_caddr", SqlDbType.NVarChar).Value = BM_address.Text;
                        comm.Parameters.Add("@p_pmodel", SqlDbType.NVarChar).Value = model.Text;
                        comm.Parameters.Add("@p_pimage", SqlDbType.NVarChar).Value = "http://www.yichang-aluminium.com.tw/DB_Images/All/" + model.Text + ".png";
                        comm.Parameters.Add("@p_pcolor", SqlDbType.NVarChar).Value = Color.Text;
                        comm.Parameters.Add("@p_plen", SqlDbType.NVarChar).Value = Lenght.Text;
                        comm.Parameters.Add("@p_pweight", SqlDbType.NVarChar).Value = Weight.Text;
                        
                        //if (string.IsNullOrEmpty(Number.Text) || IsNumeric(Number.Text))
                        if (string.IsNullOrEmpty(Number.Text) || Number.Text == "&nbsp;")
                        {
                            Number.Text = "0";
                        }
                        //if (string.IsNullOrEmpty(Price.Text) || IsNumeric(Price.Text))
                        if (string.IsNullOrEmpty(Price.Text) || Price.Text == "&nbsp;")
                        {
                            Price.Text = "0";
                        }
                        Total.Text = (Int32.Parse(Number.Text) * Double.Parse(Price.Text)).ToString();
                        comm.Parameters.Add("@p_pnum", SqlDbType.NVarChar).Value = Number.Text;
                        comm.Parameters.Add("@p_pprice", SqlDbType.NVarChar).Value = Price.Text;
                        comm.Parameters.Add("@p_ptotal", SqlDbType.NVarChar).Value = Total.Text;
                        comm.Parameters.Add("@p_pstatus", SqlDbType.Bit).Value = "True";
                        comm.Parameters.Add("@p_pdriver", SqlDbType.NVarChar).Value = "0";
                        comm.Parameters.Add("@p_ppic", SqlDbType.NVarChar).Value = "Tom";
                        SqlDataReader dr = comm.ExecuteReader();
                        dr.Dispose();
                    }
                    rowIndex++;
                }
            }
            ReSerialNumber();   //資料庫重新編號
            conn.Close();
            conn.Dispose();
        }
        catch (Exception ex)
        {
            txtMsg.Text = ex.ToString();
        }
    }
    //判斷是否有資料
    private Boolean HasRowId()
    {
        SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["YiChangDocumentsConnectionString"].ConnectionString);
        conn.Open();
        SqlCommand cmd = new SqlCommand("Select Count(*) From PurchaseDocuments", conn);

        SqlDataReader sdr = cmd.ExecuteReader();
        DataTable table = new DataTable();
        table.Load(sdr);
        string count = table.Rows[0][0].ToString();

        sdr.Dispose();
        cmd.Dispose();
        conn.Close();
        conn.Dispose();
        if (count != "0")
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    //GetCount
    private void getId()
    {
        if (!HasRowId())
        {
            val = "1";
        }
        else
        {
            SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["YiChangDocumentsConnectionString"].ConnectionString);
            conn.Open();
            SqlCommand cmd = new SqlCommand("Select Count(*) From PurchaseDocuments", conn);

            SqlDataReader sdr = cmd.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(sdr);
            val = table.Rows[0][0].ToString();
            val = (Int32.Parse(val) + 1).ToString();

            sdr.Dispose();
            cmd.Dispose();
            conn.Close();
            conn.Dispose();
        }
    }
    //Reset Serial Number
    private void ReSerialNumber()
    {
        try
        {
            SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["YiChangDocumentsConnectionString"].ConnectionString);
            conn.Open();
            string selectString = "declare @num INT select @num = 0 UPDATE PurchaseDocuments SET @num = @num + 1, id = @num";
            SqlCommand cmd = new SqlCommand(selectString, conn);
            SqlDataReader sdr = cmd.ExecuteReader();

            sdr.Dispose();
            cmd.Dispose();
            conn.Close();
            conn.Dispose();
        }
        catch (Exception ex)
        {
            txtMsg.Text = "讀取資料錯誤...";
        }
    }
    /*
    //IsNumber
    public bool IsNumeric(String strNumber)
    {
        Regex NumberPattern = new Regex("[^0-9.-]");
        return !NumberPattern.IsMatch(strNumber);
    } 
    */

    protected void Model_TextChanged(object sender, EventArgs e)
    {
        Documents selectcolor = new Documents();
        //Set DataTable
        DataTable GetTableIndex = new DataTable();
        GetTableIndex = (DataTable)ViewState["CurrentTable"];
        //Get Currnet index
        TextBox row = (TextBox)sender;
        index = (row.NamingContainer as GridViewRow).RowIndex;

        TextBox model_index = (TextBox)Gridview1.Rows[index].Cells[1].FindControl("Model");
        TextBox storage_index = (TextBox)Gridview1.Rows[index].Cells[2].FindControl("StorageNum");
        TextBox color_index = (TextBox)Gridview1.Rows[index].Cells[3].FindControl("Color");
        TextBox lenght_index = (TextBox)Gridview1.Rows[index].Cells[4].FindControl("Lenght");
        TextBox weight_index = (TextBox)Gridview1.Rows[index].Cells[5].FindControl("Weight");
        TextBox price_index = (TextBox)Gridview1.Rows[index].Cells[7].FindControl("Price");

        try
        {
            txtMsg.Text = "";
            model_index.ForeColor = Color.Black;

            SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["YichangProductDBConnection"].ConnectionString);
            conn.Open();
            string selectString = "Select p_model,p_StorageNum,p_lenght,p_weight,p_m_price From ProductAll Where p_model=@p_model";
            SqlCommand cmd = new SqlCommand(selectString, conn);
            cmd.Parameters.Add("@p_model", SqlDbType.NVarChar).Value = model_index.Text;
            SqlDataReader sdr = cmd.ExecuteReader();
            DataTable result = new DataTable();
            result.Load(sdr);
            if (!string.IsNullOrEmpty(result.Rows[0][1].ToString()))
            {
                storage_index.Text = result.Rows[0][1].ToString();
            }
            else
            {
                storage_index.Text = "0";
            }
            /*if(!string.IsNullOrEmpty(result.Rows[0][2].ToString()))
            {
                color_index.Text = selectcolor.iscolor(result.Rows[0][2].ToString().ToUpper());
            }
            else
            {
                color_index.Text = "";
            } */
            /*if (!string.IsNullOrEmpty(result.Rows[0][2].ToString()))
            {
                //lenght_index.Text = result.Rows[0][2].ToString();
            }
            else
            {
                lenght_index.Text = "";
            }                      */
            lenght_index.Text = "";
            if (!string.IsNullOrEmpty(result.Rows[0][3].ToString()))
            {
                weight_index.Text = result.Rows[0][3].ToString();
            }
            else
            {
                weight_index.Text = "0";
            }
            if (!string.IsNullOrEmpty(result.Rows[0][4].ToString()))
            {
                price_index.Text = result.Rows[0][4].ToString();
            }
            else
            {
                price_index.Text = "0";
            }

            sdr.Dispose();
            cmd.Dispose();
            conn.Close();
            conn.Dispose();
            savebtn.Enabled = true;
            color_index.Focus();
        }
        catch (Exception ex)
        {
            //txtMsg.Text = ex.ToString();
            model_index.ForeColor = Color.Red;
            model_index.Text = "";
            model_index.Focus();
            txtMsg.Text = "查無此型號!";
        }
        savebtn.Enabled = true;
    }
    //Color TextChanged
    protected void Color_TextChanged(object sender, EventArgs e)
    {
        
        TextBox row = (TextBox)sender;
        index = (row.NamingContainer as GridViewRow).RowIndex;

        TextBox color_index = (TextBox)Gridview1.Rows[index].Cells[3].FindControl("Color");
        TextBox lenght_index = (TextBox)Gridview1.Rows[index].Cells[4].FindControl("Lenght");
        txtMsg.Text = "";
        
        Documents iscolor = new Documents();
        if (iscolor.iscolor(color_index.Text.ToUpper()) == "")
        {
            color_index.ForeColor = Color.Red;
            color_index.Text = "";
            color_index.Focus();
            txtMsg.Text = "查無此顏色!";
        }
        else
        {
            color_index.ForeColor = Color.Black;
            color_index.Text = iscolor.iscolor(color_index.Text.ToUpper());
            lenght_index.Focus();
        }
       
    }
    //Number TextChange
    protected void Number_TextChanged(object sender, EventArgs e)
    {
        TextBox row = (TextBox)sender;
        index = (row.NamingContainer as GridViewRow).RowIndex;

        TextBox num_index = (TextBox)Gridview1.Rows[index].Cells[6].FindControl("Number");
        TextBox price_index = (TextBox)Gridview1.Rows[index].Cells[7].FindControl("Price");
        Label total_index = (Label)Gridview1.Rows[index].Cells[8].FindControl("Total");
        if (string.IsNullOrEmpty(price_index.Text))
        {
            price_index.Text = "0";
        }
        else if (num_index.Text != "0" || !string.IsNullOrEmpty(num_index.Text))
        {
            //num_index.Text = "0";
            total_index.Text = (Int32.Parse(num_index.Text) * Double.Parse(price_index.Text)).ToString();
            if(price_index.Text == "0")
            {
                price_index.Focus();
            }
            else
            {
                num_index.Focus();
            }
            
        }
        else
        {
            txtMsg.Text = "數量不可為 0 或 空白";
            num_index.Focus();
        }
        
    }
    //Price TextChange
    protected void Price_TextChanged(object sender, EventArgs e)
    {
        TextBox row = (TextBox)sender;
        index = (row.NamingContainer as GridViewRow).RowIndex;

        TextBox num_index = (TextBox)Gridview1.Rows[index].Cells[6].FindControl("Number");
        TextBox price_index = (TextBox)Gridview1.Rows[index].Cells[7].FindControl("Price");
        Label total_index = (Label)Gridview1.Rows[index].Cells[8].FindControl("Total");
        if (string.IsNullOrEmpty(num_index.Text))
        {
            num_index.Text = "0";
        }
        else if (price_index.Text != "0" || !string.IsNullOrEmpty(price_index.Text))
        {
            total_index.Text = (Int32.Parse(num_index.Text) * Double.Parse(price_index.Text)).ToString();
            if (price_index.Text == "0")
            {
                price_index.Focus();
            }
            else
            {
                num_index.Focus();
            } 

        }
        else
        {
            txtMsg.Text = "數量不可為 0 或 空白";
            num_index.Focus();
        }
    }
    protected void Button2_Click(object sender, EventArgs e)
    {
        saveGvdate();
        //DataTable dt_display = (DataTable)ViewState["CurrentTable"];

        //GV_Display.DataSource = dt_display;
        //GV_Display.DataBind();
    }
}