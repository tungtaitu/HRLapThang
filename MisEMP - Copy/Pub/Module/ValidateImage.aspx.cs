using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Drawing.Imaging;

namespace PLSignWeb2.Pub.Module
{
	/// <summary>
	/// ValidateImage ���K�n�y�z�C
	/// </summary>
	public partial class ValidateImage : System.Web.UI.Page
	{
		protected void Page_Load(object sender, System.EventArgs e)
		{
			// �N�ϥΪ̵{���X�m�󦹥H��l�ƺ���
			if (! IsPostBack)
			{
				GenImage();
			}
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: ���I�s�� ASP.NET Web Form �]�p�u�㪺���n���C
			//
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
				
				//newBitmap.Save(Server.MapPath(Session["PageLayer"] + "images/Verify.jpg"), ImageFormat.Png); //�N�ϧΦs��
				//imgVerify.ImageUrl=("../../images/Verify.jpg");										 //�I�s�ϧ�	
				Response.ContentType = "Image/jpeg"; 
				newBitmap.Save(Response.OutputStream, ImageFormat.Jpeg); //�N�ϧΦs��
				
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

		
	}
}
