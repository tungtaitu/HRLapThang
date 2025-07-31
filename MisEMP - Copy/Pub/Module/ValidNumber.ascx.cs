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
	///		ValidNumber ���K�n�y�z�C
	/// </summary>
	public partial  class ValidNumber : System.Web.UI.UserControl
	{

		protected void Page_Load(object sender, System.EventArgs e)
		{
			// �N�ϥΪ̵{���X�m�󦹥H��l�ƺ���
			if (! IsPostBack)
			{
				//GenImage();
			}
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: ���I�s�� ASP.NET Web Form �]�p�u�㪺���n���C
			//
			
			//�p��o�Ӻ����Ҧb���h���O����
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
		/// �����]�p�u��䴩�ҥ��ݪ���k - �ФŨϥε{���X�s�边�ק�
		/// �o�Ӥ�k�����e�C
		/// </summary>
		private void InitializeComponent()
		{

		}
		#endregion

		//���o�üƭ�0~99999
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

		//���ͪ��üƭȼg�J����ɤ�
		private void GenImage()
		{

			Bitmap newBitmap = null;
			Graphics g = null;
			
			try 
			{
				Random r = GetRandom(0);
	
				Session["ValidNumber"] = r.Next(0,99999).ToString("00000");			//���Ͷüƭ�
				Font fontCounter = new Font("Times New Roman", 18);

				// calculate size of the string.
				newBitmap = new Bitmap(1,1,PixelFormat.Format32bppArgb);				//�}�Ҥ@��32bits���m���
				g = Graphics.FromImage(newBitmap);
				SizeF stringSize = g.MeasureString(Session["ValidNumber"].ToString() , fontCounter);	
				int nWidth = (int)stringSize.Width;										//�Q�ο�J���r����q�Ϫ����e
				int nHeight = (int)stringSize.Height;
				g.Dispose();
				newBitmap.Dispose();
      
				newBitmap = new Bitmap(nWidth,nHeight,PixelFormat.Format32bppArgb);
				g = Graphics.FromImage(newBitmap);
				g.FillRectangle(new SolidBrush(Color.Silver), new Rectangle(0,0,nWidth,nHeight));//�]�w�Ϫ��C��M���e
			
				g.DrawString(Session["ValidNumber"].ToString() , fontCounter, new SolidBrush(Color.Black), 0, 0);
				
				newBitmap.Save(Server.MapPath(Session["PageLayer"] + "images/Verify.jpg"), ImageFormat.Png); //�N�ϧΦs��
				imgVerify.ImageUrl=("../../images/Verify.jpg");										 //�I�s�ϧ�	
				
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
