namespace PLSignWeb2.Pub.Module
{
	using System;
	using System.Data;
	using System.Drawing;
	using System.Web;
	using System.Web.UI.WebControls;
	using System.Web.UI.HtmlControls;
	using System.Drawing.Imaging;

	/// <summary>
	///		ValidNumber 的摘要描述。
	/// </summary>
	public partial  class ValidNumber : System.Web.UI.UserControl
	{

		protected void Page_Load(object sender, System.EventArgs e)
		{
			// 將使用者程式碼置於此以初始化網頁
			if (! IsPostBack)
			{
				//GenImage();
			}
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: 此呼叫為 ASP.NET Web Form 設計工具的必要項。
			//
			
			//計算這個網頁所在的層次是那裡
			int i,j = 0;
			string strPageLayer = "";
			string LocalPath = PccCommonForC.PccToolFunc.Upper(Server.MapPath("."));   

			j = LocalPath.IndexOf(PccCommonForC.PccToolFunc.Upper(Application["EDPNET"].ToString()));
		
			try
			{
				for (i = 1 ; i < LocalPath.Substring(j).Split('\\').Length ; i++)
				{
					strPageLayer += "../";
				}
				Session["PageLayer"] = strPageLayer;
			}
			catch
			{
				Session["PageLayer"] = "";
			}
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// 此為設計工具支援所必需的方法 - 請勿使用程式碼編輯器修改
		/// 這個方法的內容。
		/// </summary>
		private void InitializeComponent()
		{

		}
		#endregion

		//取得亂數值0~99999
		private Random GetRandom(int seed) 
		{
			Random r = (Random)HttpContext.Current.Cache.Get("RandomNumber");
			if (r == null)
			{
				if (seed == 0)
					r = new Random();
				else
					r = new Random(seed);
				HttpContext.Current.Cache.Insert("RandomNumber",r);
			}
			return r;
		}

		private Random GetRandom() 
		{
			return GetRandom(0);
		}

		//產生的亂數值寫入到圖檔中
		private void GenImage()
		{

			Bitmap newBitmap = null;
			Graphics g = null;
			
			try 
			{
				Random r = GetRandom(0);
	
				Session["ValidNumber"] = r.Next(0,99999).ToString("00000");			//產生亂數值
				Font fontCounter = new Font("Times New Roman", 18);

				// calculate size of the string.
				newBitmap = new Bitmap(1,1,PixelFormat.Format32bppArgb);				//開啟一個32bits的彩色圖
				g = Graphics.FromImage(newBitmap);
				SizeF stringSize = g.MeasureString(Session["ValidNumber"].ToString() , fontCounter);	
				int nWidth = (int)stringSize.Width;										//利用輸入的字串測量圖的長寬
				int nHeight = (int)stringSize.Height;
				g.Dispose();
				newBitmap.Dispose();
      
				newBitmap = new Bitmap(nWidth,nHeight,PixelFormat.Format32bppArgb);
				g = Graphics.FromImage(newBitmap);
				g.FillRectangle(new SolidBrush(Color.Silver), new Rectangle(0,0,nWidth,nHeight));//設定圖的顏色和長寬
			
				g.DrawString(Session["ValidNumber"].ToString() , fontCounter, new SolidBrush(Color.Black), 0, 0);
				
				newBitmap.Save(Server.MapPath(Session["PageLayer"] + "images/Verify.jpg"), ImageFormat.Png); //將圖形存檔
				imgVerify.ImageUrl=("../../images/Verify.jpg");										 //呼叫圖形	
				
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

		public bool IsValid()
		{
			if ( txtValidNumber.Text == Session["ValidNumber"].ToString())
			{
				return true;
			}
			else
			{
				//GenImage();
				return false;
			}
		}     
		
		public string CssClass
		{
			set
			{
				txtValidNumber.CssClass = value;
			}
		}
		

	}
}
