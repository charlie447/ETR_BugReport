using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Data;

namespace Excel.Operation
{
    public class MailClass
    {
        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="data">邮件内容</param>
        public void SendStrMail(string data)
        {
            try
            {
                System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
                if (appConfig.conf.Toaddress.IndexOf(',') > -1)
                {
                    string[] mails = appConfig.conf.Toaddress.Split(',');//多个收信地址用逗号隔开
                    for (int counti = 0; counti < mails.Length; counti++)
                    {
                        if (mails[counti].Trim() != "")
                        {
                            msg.To.Add(mails[counti]);
                        }
                    }
                }
                else
                {
                    msg.To.Add(appConfig.conf.Toaddress);//添加单一收信地址
                }
                msg.To.Add(appConfig.conf.Fromaddress);

                msg.From = new System.Net.Mail.MailAddress(appConfig.conf.Fromaddress, appConfig.conf.Fromaddressname, System.Text.Encoding.UTF8);
                /* 上面3个参数分别是发件人地址（可以随便写），发件人姓名，编码*/
                string sub = appConfig.conf.MailSub;//邮件标题 
                if (appConfig.body.count == 0)
                {
                    msg.Subject = appConfig.conf.MailSub;
                }
                else
                {
                    msg.Subject = data;
                }
                msg.SubjectEncoding = System.Text.Encoding.UTF8;//邮件标题编码
                //long tol = GateCount + HandCount;
                msg.Body = data;
                msg.BodyEncoding = System.Text.Encoding.UTF8;//邮件内容编码 
                msg.IsBodyHtml = true;//是否是HTML邮件 
                msg.Priority = MailPriority.High;

                SmtpClient client = new SmtpClient();
                client.Credentials = new System.Net.NetworkCredential(appConfig.conf.Fromaddress, appConfig.conf.FromMailPassword);
                //在zj.com注册的邮箱和密码 
                client.Host = appConfig.conf.SmtpServer;//邮件发送服务器，上面对应的是该服务器上的发信帐号和密码

                object userState = msg;
                try
                {
                    client.SendAsync(msg, userState);//开始发送
                }
                catch (System.Net.Mail.SmtpException ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message.ToString());
            }
        }
    }
    //发送带表格邮件
    public class SendMail
    {
        public static MailClass Mail = new MailClass();
        public static string LargeMailBody = "";//初始化多表格邮件内容
        //单一表格邮件内容
        public static void SendMsg(DataTable data)
        {
            string MailBody = "<p style=\"font-size: 10pt\">以下内容为系统自动发送，请勿直接回复，谢谢。</p><table cellspacing=\"1\" cellpadding=\"3\" border=\"0\" bgcolor=\"000000\" style=\"font-size: 10pt;line-height: 15px;\">";
            MailBody += "<div align=\"center\">";
            MailBody += "<tr>";
            for (int hcol = 0; hcol < data.Columns.Count; hcol++)
            {
                MailBody += "<td bgcolor=\"999999\">&nbsp;&nbsp;&nbsp;";
                MailBody += data.Columns[hcol].ColumnName;
                MailBody += "&nbsp;&nbsp;&nbsp;</td>";
            }
            MailBody += "</tr>";

            for (int row = 0; row < data.Rows.Count; row++)
            {
                MailBody += "<tr>";
                for (int col = 0; col < data.Columns.Count; col++)
                {
                    MailBody += "<td bgcolor=\"dddddd\">&nbsp;&nbsp;&nbsp;";
                    MailBody += data.Rows[row][col].ToString();
                    MailBody += "&nbsp;&nbsp;&nbsp;</td>";
                }
                MailBody += "</tr>";
            }
            MailBody += "</table>";
            MailBody += "</div>";
            Mail.SendStrMail(MailBody);
        }
        //单一表格邮件内容
        public static void SendMsg(DataGridView data)
        {
            string MailBody = "<p style=\"font-size: 10pt\">以下内容为系统自动发送，请勿直接回复，谢谢。</p><table cellspacing=\"1\" cellpadding=\"3\" border=\"0\" bgcolor=\"000000\" style=\"font-size: 10pt;line-height: 15px;\">";

            MailBody += "<div align=\"center\">";
            MailBody += "<tr>";
            for (int hcol = 0; hcol < data.Columns.Count; hcol++)
            {
                MailBody += "<td bgcolor=\"999999\">&nbsp;&nbsp;&nbsp;";
                MailBody += data.Columns[hcol].HeaderText.ToString();
                MailBody += "&nbsp;&nbsp;&nbsp;</td>";
            }
            MailBody += "</tr>";

            for (int row = 0; row < data.Rows.Count; row++)
            {
                MailBody += "<tr>";
                for (int col = 0; col < data.Columns.Count; col++)
                {
                    MailBody += "<td bgcolor=\"dddddd\">&nbsp;&nbsp;&nbsp;";
                    MailBody += data.Rows[row].Cells[col].Value.ToString();
                    MailBody += "&nbsp;&nbsp;&nbsp;</td>";
                }
                MailBody += "</tr>";
            }
            MailBody += "</table>";
            MailBody += "</div>";
            Mail.SendStrMail(MailBody);
        }
        //多表格邮件内容
        public static void SendLargeMsg(DataTable data, string title = "")
        {
            if (title != "")
                LargeMailBody += "<p style=\"font-size: 10pt\">" + title + "</p>";

            LargeMailBody += "<div align=\"center\">";
            LargeMailBody += "<table cellspacing=\"1\" cellpadding=\"3\" border=\"0\" bgcolor=\"000000\" style=\"font-size: 10pt;line-height: 15px;\">";

            LargeMailBody += "<tr>";
            for (int hcol = 0; hcol < data.Columns.Count; hcol++)
            {
                LargeMailBody += "<td bgcolor=\"999999\">&nbsp;&nbsp;&nbsp;";
                LargeMailBody += data.Columns[hcol].ColumnName;
                LargeMailBody += "&nbsp;&nbsp;&nbsp;</td>";
            }
            LargeMailBody += "</tr>";

            for (int row = 0; row < data.Rows.Count; row++)
            {
                LargeMailBody += "<tr>";
                for (int col = 0; col < data.Columns.Count; col++)
                {
                    LargeMailBody += "<td bgcolor=\"dddddd\">&nbsp;&nbsp;&nbsp;";
                    LargeMailBody += data.Rows[row][col].ToString();
                    LargeMailBody += "&nbsp;&nbsp;&nbsp;</td>";
                }
                LargeMailBody += "</tr>";
            }
            LargeMailBody += "</table><br>";
            LargeMailBody += "</div>";
        }
        //多表格邮件内容
        public static void SendLargeMsg(DataGridView data, string title = "")
        {
            if (title != "")
                LargeMailBody += "<p style=\"font-size: 10pt\">" + title + "</p>";

            LargeMailBody += "<div align=\"center\">";

            LargeMailBody += "<table cellspacing=\"1\" cellpadding=\"3\" border=\"0\" bgcolor=\"000000\" style=\"font-size: 10pt;line-height: 15px;\">";

            LargeMailBody += "<tr>";
            for (int hcol = 0; hcol < data.Columns.Count; hcol++)
            {
                LargeMailBody += "<td bgcolor=\"999999\">&nbsp;&nbsp;&nbsp;";
                LargeMailBody += data.Columns[hcol].HeaderText.ToString();
                LargeMailBody += "&nbsp;&nbsp;&nbsp;</td>";
            }
            LargeMailBody += "</tr>";

            for (int row = 0; row < data.Rows.Count; row++)
            {
                LargeMailBody += "<tr>";
                for (int col = 0; col < data.Columns.Count; col++)
                {
                    LargeMailBody += "<td bgcolor=\"dddddd\">&nbsp;&nbsp;&nbsp;";
                    LargeMailBody += data.Rows[row].Cells[col].Value.ToString();
                    LargeMailBody += "&nbsp;&nbsp;&nbsp;</td>";
                }
                LargeMailBody += "</tr>";
            }
            LargeMailBody += "</table><br>";
            LargeMailBody += "</div>";
        }
        private void button1_Click(object sender, EventArgs e)
        {
            t1 = databind(qishishijian.Value, jieshushijian.Value);
            t2 = databind1(qishishijian.Value, jieshushijian.Value);
            SendMail.LargeMailBody = "";
            try
            {
                SendMail.SendLargeMsg(t1, "测试内容1");
                SendMail.SendLargeMsg(t2, "测试内容2");
                SendMail.Mail.SendStrMail("<p style=\"font-size: 10pt\">以下内容为系统自动发送，请勿直接回复，谢谢。</p>" + SendMail.LargeMailBody);
            }
            catch { MessageBox.Show("没有可以发送的内容！"); }
        }
    }

}
