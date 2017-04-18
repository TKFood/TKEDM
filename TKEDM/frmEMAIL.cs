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
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataTable dt = new DataTable();       
        string tablename = null;

        Thread TD;
        StringBuilder content = new StringBuilder();

        string SENDMAIL;
        string SENDNAME;
        string SMTP;
        int SMTPPORT;
        string PASSWORD;

        string GAID;
        string GAEA;
        string GACS;
        string GACN;

        public frmEMAIL()
        {
            InitializeComponent();
            

            SERACHCONFIG();
            SERACHGACONFIG();

            SETBODY();
        }

        #region FUNCTION
        public void SERACHCONFIG()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT TOP 1 [SENDMAIL],[SENDNAME],[SMTP],[SMTPPORT],[PASSWORD] FROM [TKEDM].[dbo].[CONFIG]  ");              
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        SENDMAIL=ds1.Tables["TEMPds1"].Rows[0]["SENDMAIL"].ToString();
                        SENDNAME = ds1.Tables["TEMPds1"].Rows[0]["SENDNAME"].ToString();
                        SMTP = ds1.Tables["TEMPds1"].Rows[0]["SMTP"].ToString();
                        SMTPPORT = Convert.ToInt16(ds1.Tables["TEMPds1"].Rows[0]["SMTPPORT"].ToString());
                        PASSWORD = ds1.Tables["TEMPds1"].Rows[0]["PASSWORD"].ToString();

                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }
        }
        public void SERACHGACONFIG()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT TOP 1 [GAID],[GAEA],[GACS],[GACN]  FROM [TKEDM].[dbo].[GACONFIG]  ");
                sbSql.AppendFormat(@"  ");


                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        //GAID = ds2.Tables["TEMPds2"].Rows[0]["GAID"].ToString();
                        //GAEA = ds2.Tables["TEMPds2"].Rows[0]["GAEA"].ToString();
                        //GACS = ds2.Tables["TEMPds2"].Rows[0]["GACS"].ToString();
                        //GACN = ds2.Tables["TEMPds2"].Rows[0]["GACN"].ToString();
                        textBox2.Text = ds2.Tables["TEMPds2"].Rows[0]["GAID"].ToString();
                        textBox3.Text = ds2.Tables["TEMPds2"].Rows[0]["GAEA"].ToString();
                        textBox4.Text = ds2.Tables["TEMPds2"].Rows[0]["GACS"].ToString();
                        textBox5.Text = ds2.Tables["TEMPds2"].Rows[0]["GACN"].ToString();
                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }
        }
        public void SENDEMAIL()
        {
            
            try
            {

                content.Clear();
                content.AppendFormat(textBox1.Text);

                MailAddress receiverAddress = new MailAddress("tk160115@gmail.com", "hi");//<-這物件只是用來設定郵件帳號而已~
                //MailAddress receiverAddress = new MailAddress("tk290@tkfood.com.tw", "t1");//<-這物件只是用來設定郵件帳號而已~
                MailAddress senderAddress = new MailAddress(SENDMAIL, SENDNAME);               
                MailMessage mail = new MailMessage(senderAddress, receiverAddress);//<-這物件是郵件訊息的部分~需設定寄件人跟收件人~可直接打郵件帳號也可以使用MailAddress物件~

                mail.Priority = MailPriority.Normal;
                mail.Subject = "老楊食品";
                mail.Body = content.ToString();
                mail.IsBodyHtml = true;//<-如果要這封郵件吃html的話~這屬性就把他設為true~~

                //Attachment attachment = new Attachment(@"");//<-這是附件部分~先用附件的物件把路徑指定進去~
                //mail.Attachments.Add(attachment);//<-郵件訊息中加入附件
             


                SmtpClient MySmtp = new SmtpClient(SMTP, SMTPPORT); //允許程式使用smtp來發mail，並設定smtp server & port
                MySmtp.EnableSsl = false; //開啟SSL連線 (gmail體系須使用SSL連線)
                MySmtp.UseDefaultCredentials = true;
                //ps=tkmail413
                MySmtp.Credentials = new NetworkCredential(SENDMAIL, PASSWORD); //設定帳號與密碼 需要using system.net;
                
               
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
            //textBox1.Text = textBox1.Text + @"<img src=""http://www.google-analytics.com/collect?v=1&t=event&tid=UA-92879762-1&cid=0001&ec=email&ea=open5&el=recipient_id&cs=opennewsletter5&cm=email&cn=TK201704/>""/>" + Environment.NewLine;
            //textBox1.Text = textBox1.Text +  Environment.NewLine;
            //textBox1.Text = textBox1.Text + @"<img src=""http://www.google-analytics.com/collect?v=1&t=event&tid=" + GAID+ @"&cid=0001&ec=email&ea=" + GAEA + @"&el=recipient_id&cs=" + GACS +@"&cm=email&cn=" + GACN+@"/>""/>" + Environment.NewLine;
            textBox1.Text = textBox1.Text + @"<img src=""http://www.google-analytics.com/collect?v=1&t=event&tid=" + textBox2.Text + @"&cid=0001&ec=email&ea=" + textBox3.Text + @"&el=recipient_id&cs=" + textBox4.Text + @"&cm=email&cn=" + textBox5.Text + @"/>""/>" + Environment.NewLine;
            textBox1.Text = textBox1.Text + "<a href=http://new.tkfood.com.tw>老楊食品</a>" + Environment.NewLine;
            textBox1.Text = textBox1.Text + "</div><br>" + Environment.NewLine;
            textBox1.Text = textBox1.Text + "</body>" + Environment.NewLine;
            textBox1.Text = textBox1.Text + "</html>" + Environment.NewLine;
        }

        public void SEARCHPRIMEMEMBER()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [ID],[NAME],[SEX],[BIRTHDAT],[EMAIL]  ");
                sbSql.AppendFormat(@"  FROM [TKEDM].[dbo].[PRIMEMEMBER] ");
                sbSql.AppendFormat(@"  WHERE [SEX]='{0}'",comboBox1.Text);
                sbSql.AppendFormat(@"  AND DATEDIFF (YEAR,[BIRTHDAT],GETDATE()) >={0} AND DATEDIFF (YEAR,[BIRTHDAT],GETDATE()) <={1}",numericUpDown1.Value.ToString(),numericUpDown2.Value.ToString());
                sbSql.AppendFormat(@"  ");


                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "TEMPds3");
                sqlConn.Close();


                if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds3.Tables["TEMPds3"];
                        dataGridView1.AutoResizeColumns();
                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }
        }

 
        public void SENDPRIMEMEMBEREMAIL()
        {
            DataSet DSMAIL = ds3;
            string EMAIL;
            for (int i=0;i< DSMAIL.Tables[0].Rows.Count;i++)
            {
                EMAIL = DSMAIL.Tables[0].Rows[i]["EMAIL"].ToString();
                if (!string.IsNullOrEmpty(EMAIL))
                {
                    SETPRIMEMEMBEREBODY(DSMAIL.Tables[0].Rows[i]["NAME"].ToString());
                    //MessageBox.Show(EMAIL);
                }
            }
            
            
        }

        public void SETPRIMEMEMBEREBODY(string NAME)
        {

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SENDEMAIL();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETBODY();
            content.Clear();
            content.AppendFormat(textBox1.Text);
            webBrowser1.DocumentText= content.ToString();
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SEARCHPRIMEMEMBER();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SENDPRIMEMEMBERE();
        }
        public void SENDPRIMEMEMBERE()
        {
            content.Clear();
            content.AppendFormat(textBox1.Text);
            SENDPRIMEMEMBEREMAIL();
        }
        #endregion


    }
}
