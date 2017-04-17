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
using System.Net;
using System.Collections.Specialized;

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
        StringBuilder content = new StringBuilder();

        public frmEMAIL()
        {
            InitializeComponent();
            SETBODY();
        }
        #region FUNCTION
        public void SENDEMAIL()
        {
            
            try
            {

                content.Clear();
                content.AppendFormat(textBox1.Text);

                MailAddress receiverAddress = new MailAddress("tk160115@gmail.com", "hi");//<-這物件只是用來設定郵件帳號而已~
                //MailAddress receiverAddress = new MailAddress("tk290@tkfood.com.tw", "t1");//<-這物件只是用來設定郵件帳號而已~
                MailAddress senderAddress = new MailAddress("aurora@tkfood.com.tw", "老楊食品");               
                MailMessage mail = new MailMessage(senderAddress, receiverAddress);//<-這物件是郵件訊息的部分~需設定寄件人跟收件人~可直接打郵件帳號也可以使用MailAddress物件~

                mail.Priority = MailPriority.Normal;
                mail.Subject = "老楊食品";
                mail.Body = content.ToString();
                mail.IsBodyHtml = true;//<-如果要這封郵件吃html的話~這屬性就把他設為true~~

                //Attachment attachment = new Attachment(@"");//<-這是附件部分~先用附件的物件把路徑指定進去~
                //mail.Attachments.Add(attachment);//<-郵件訊息中加入附件
             


                SmtpClient MySmtp = new SmtpClient("officemail.cloudmax.com.tw", 25); //允許程式使用smtp來發mail，並設定smtp server & port
                MySmtp.EnableSsl = false; //開啟SSL連線 (gmail體系須使用SSL連線)
                MySmtp.UseDefaultCredentials = true;
                //ps=tkmail413
                MySmtp.Credentials = new NetworkCredential("aurora@tkfood.com.tw", ""); //設定帳號與密碼 需要using system.net;
                
               
                MySmtp.Send(mail);

                MySmtp = null; //將MySmtp清空
                mail.Dispose(); //釋放資源

                MessageBox.Show("OK");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.ReadLine();
        }

        public void SETBODY()
        {
            //textBox1.Text = "<img src=http://www.google-analytics.com/collect?v=1&t=event&tid=UA-92879762-1&cid=0001&ec=email&ea=open&el=recipient_id&cs=opennewsletter&cm=email&cn=TK201704>" + Environment.NewLine;
            textBox1.Text = "<html>" + Environment.NewLine;
            textBox1.Text = textBox1.Text + "<body>" + Environment.NewLine;
            textBox1.Text = textBox1.Text + "<div>Hello You." + Environment.NewLine;
            textBox1.Text = textBox1.Text + @"<img src=""http://www.google-analytics.com/collect?v=1&t=event&tid=UA-92879762-1&cid=0001&ec=email&ea=open5&el=recipient_id&cs=opennewsletter5&cm=email&cn=TK201704/>""/>" + Environment.NewLine;
            textBox1.Text = textBox1.Text + "<a href=http://new.tkfood.com.tw>老楊食品</a>" + Environment.NewLine;
            textBox1.Text = textBox1.Text + "</div><br>" + Environment.NewLine;
            textBox1.Text = textBox1.Text + "</body>" + Environment.NewLine;
            textBox1.Text = textBox1.Text + "</html>" + Environment.NewLine;
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SENDEMAIL();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            //content.AppendFormat(" <div>Hello You're from country.</div>");
            //content.AppendFormat("<a href=http://new.tkfood.com.tw'>老楊食品</a>");

            //string htmlBody = @"<img src=""http://www.google-analytics.com/collect?v=1&t=event&tid=UA-92879762-1&cid=0001&ec=email&ea=open&el=recipient_id&cs=opennewsletter&cm=email&cn=TK201704/>""/>";
            //string htmlBody = @" < html><body><img src=""http://www.google-analytics.com/collect?v=1&t=event&tid=UA-92879762-1&cid=0001&ec=email&ea=open&el=recipient_id&cs=opennewsletter&cm=email&cn=TK201704/>""/></body></html>";
            content.Clear();
            //content.AppendFormat("<img src=http://www.google-analytics.com/collect?v=1&t=event&tid=UA-92879762-1&cid=0001&ec=email&ea=open&el=recipient_id&cs=opennewsletter&cm=email&cn=TK201704/>");
            //content.AppendFormat(htmlBody.ToString());
            content.AppendFormat(textBox1.Text);
            webBrowser1.DocumentText= content.ToString();
            
        }
        #endregion


    }
}
