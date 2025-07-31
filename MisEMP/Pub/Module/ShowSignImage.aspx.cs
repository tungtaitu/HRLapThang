using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using PccCommonForC;
using PccBsSystemForC;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;

public partial class Pub_Module_ShowSignImage : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            GenImage();
        }
    }

     private string getUserName()
		{
		//	bs_UserManager mybs = new bs_UserManager(ConfigurationSettings.AppSettings["ConnectionType"] , ConfigurationSettings.AppSettings["ConnectionServer"], ConfigurationSettings.AppSettings["ConnectionDB"], ConfigurationSettings.AppSettings["ConnectionUser"], ConfigurationSettings.AppSettings["ConnectionPwd"],Session["UserIDAndName"].ToString(),ConfigurationSettings.AppSettings["EventLogPath"]);
         bs_UserManager mybs = new bs_UserManager(ConfigurationSettings.AppSettings["ConnectionType"], ConfigurationSettings.AppSettings["ConnectionServer"], ConfigurationSettings.AppSettings["ConnectionDB"], ConfigurationSettings.AppSettings["ConnectionUser"], ConfigurationSettings.AppSettings["ConnectionPwd"],"", ConfigurationSettings.AppSettings["EventLogPath"]);
         string strXML = "<PccMsg><user_id>" + Request.QueryString["SignId"] + "</user_id></PccMsg>";
			string strReturn = mybs.DoReturnStr("GetUserData",strXML,string.Empty);
			PccMsg myMsg = new PccMsg(strReturn);

			if (myMsg.Query("Return").Equals("OK"))
				strReturn = myMsg.Query("user_desc"); 
			else
				strReturn = string.Empty;

			return strReturn;

		}

   

		//產生的字串寫入到圖檔中
		private void GenImage()
		{

			Bitmap newBitmap = null;
			Graphics g = null;
			
			try 
			{
				string signName = getUserName();
				Font fontCounter = new Font("標楷體", 18,FontStyle.Bold);
				

				// calculate size of the string.
				newBitmap = new Bitmap(1,1,PixelFormat.Format32bppArgb);				//開啟一個32bits的彩色圖
				g = Graphics.FromImage(newBitmap);
				SizeF stringSize = g.MeasureString(signName , fontCounter);	
				int nWidth = (int)stringSize.Width;										//利用輸入的字串測量圖的長寬
				int nHeight = (int)stringSize.Height;

				//設定圖形的高度
				int iSetHieght = 100;
				//設定圖形的啟始高度點 公式 x(表示字的大小)+2y(表示啟始高度) = iSetHieght;
                int iStartHieght = (int)Math.Ceiling((Double)(((100 - nHeight) / 2)));

				//nWidth = (int)Math.Ceiling(nWidth / 2);
				//nHeight = nHeight * 2;
				g.Dispose();
				newBitmap.Dispose();
      
				newBitmap = new Bitmap(nWidth,100,PixelFormat.Format32bppArgb);
				g = Graphics.FromImage(newBitmap);
				g.FillRectangle(new SolidBrush(Color.White), new Rectangle(0,0,nWidth,iSetHieght));//設定圖的顏色和長寬
			
				g.DrawString(signName , fontCounter, new SolidBrush(Color.Black), 0, iStartHieght);
				
				Response.ContentType = "Image/jpeg"; 
				newBitmap.Save(Response.OutputStream, ImageFormat.Jpeg); //將圖形存檔
				
			} 
			catch (Exception e)
			{
				e.Message.ToString();
			}
			finally 
			{
				if (null != g) g.Dispose();
				if (null != newBitmap) newBitmap.Dispose();
			}
		}

		#region Web Form 設計工具產生的程式碼
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: 此為 ASP.NET Web Form 設計工具所需的呼叫。
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// 此為設計工具支援所必須的方法 - 請勿使用程式碼編輯器修改
		/// 這個方法的內容。
		/// </summary>
		private void InitializeComponent()
		{    
			this.Load += new System.EventHandler(this.Page_Load);
		}
		#endregion
	}




