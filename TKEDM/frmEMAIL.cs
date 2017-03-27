using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using System.Collections.Generic;
using System.Text;
using System.Net.Mail;//<-基本上發mail就用這個class

namespace TKEDM
{
    public partial class frmEMAIL : Form
    {

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataTable dt = new DataTable();       
        string tablename = null;

        Thread TD;

        public frmEMAIL()
        {
            InitializeComponent();
        }
        #region FUNCTION
        public void SENDEMAIL()
        {
            try
            {
                SmtpClient sc = new SmtpClient("mail.tkfood.com.tw");//<-宣告的時候可以先給主機名稱~記住喔~這是發送端的主機名稱~
                sc.Port = 25;
                MailAddress receiverAddress = new MailAddress("tk290@tkfood.com.tw", "t1");//<-這物件只是用來設定郵件帳號而已~
                MailAddress senderAddress = new MailAddress("aurora@tkfood.com.tw", "1");
                MailMessage mail = new MailMessage(senderAddress, receiverAddress);//<-這物件是郵件訊息的部分~需設定寄件人跟收件人~可直接打郵件帳號也可以使用MailAddress物件~
                mail.Subject = "test";
                mail.Body = "<a href='http://tw.yahoo.com'>yahoo</a>";
                mail.IsBodyHtml = true;//<-如果要這封郵件吃html的話~這屬性就把他設為true~~

                //Attachment attachment = new Attachment(@"");//<-這是附件部分~先用附件的物件把路徑指定進去~
                //mail.Attachments.Add(attachment);//<-郵件訊息中加入附件

                sc.Send(mail);//<-這樣就送出去拉~

                MessageBox.Show("OK");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.ReadLine();
        }
    


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
        SENDEMAIL();
        }

        #endregion

    }
}
