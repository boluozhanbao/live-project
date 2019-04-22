using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections;
using System.Net;
using System.Security.Cryptography;

namespace VBAInExcel
{
    class Message
    {
        string Qname;//过滤助教,系统消息Ifnoassistant(Qname);
        string Qnum;//键值，活跃度//insertQnun(Qnum);
        string Qtime;//时间段//Ifactivative(Qtime);
        string Qcontent;//口令//bool Ifcorrespond(Qcontrent);
        public void SetQname(string name)
        {
            Qname = name;
            return;
        }
        public string GetQname()
        {
            return Qname;
        }
        public void SetQnum(string num)
        {
            Qnum = num;
            return;
        }
        public string GetQnum()
        {
            return Qnum;
        }
        public void SetQtime(string time)
        {
            Qtime = time;
            return;
        }
        public string GetQtime()
        {
            return Qtime;
        }
        public void SetQcontent(string content)
        {
            Qcontent = content;
            return;
        }
        public string GetQcontent()
        {
            return Qcontent;
        }
    }
    //文件处理
    class FileProcessing
    {
        string path;
        public FileProcessing(string path)
        {
            this.path = path;
        }
        public void SetPath(string path)
        {
            this.path = path;
            return;
        }
        public string GetPath()
        {
            return path;
        }
        public List<Message> messages = new List<Message>();

        public void Read()
        {
            StreamReader sr = new StreamReader(path, Encoding.Default);
            String line;
            int flag = 0;
            string time = "", name = "", num = "", message = "";
            while ((line = sr.ReadLine()) != null)
            {

                //依次匹配时间 用户名 账号 消息   
                Regex r = new Regex("(?<time>([0-9]{4})-([0-9]{2})-([0-9]{2}) ([0-9]{2}):([0-9]{2}):([0-9]{2}))(?<name>.*)[(](?<num>.*)[)]");
                Match m = r.Match(line.ToString());
                if (m.Value.ToString() != "")
                {
                    if (flag > 0)
                    {
                        Message mess = new Message();
                        mess.SetQtime(time);
                        mess.SetQname(name);
                        mess.SetQnum(num);
                        mess.SetQcontent(message);
                        mess.SetQcontent(message);
                        messages.Add(mess);
                    }
                    time = m.Result("${time}");
                    name = m.Result("${name}");
                    num = m.Result("${num}");
                    message = "";
                    flag++;
                }
                else
                {
                    message += line;
                }
            }
            Message less = new Message();
            less.SetQtime(time);
            less.SetQname(name);
            less.SetQnum(num);
            less.SetQcontent(message);
            messages.Add(less);
        }
    }
    //文件过滤
    class Pprogram
    {
        static Hashtable ht = new Hashtable();
        public static bool Ifcorrespond(string Qcontent, string password)
        {
            password = "#" + password + "#";
            Regex regex = new Regex(@"#+\w+#");//构造方法里面的参数就是匹配规则。

            //声明匹配结果集对象

            MatchCollection matchs;

            //调用匹配多个出多个结果的方法进行匹配，将返回的匹配结果集对象赋值给matchs

            matchs = regex.Matches(Qcontent);

            //输出所有匹配结果

            foreach (Match match in matchs)
            {
                if (match.ToString().Equals(password))
                    return true;
            }
            return false;
        }
        public static bool Ifactivate(string Qtime, string Begingtime, string Endtime)
        {
            if (Qtime.CompareTo(Begingtime) >= 0 && Qtime.CompareTo(Endtime) <= 0)
                return true;
            return false;
        }
        //过滤助教和系统信息
        public static bool Ifnoassistant(string Qname)
        {
            Regex reg = new Regex("助教_");
            Match m = reg.Match(Qname);
            if (m.Value.ToString().Equals("助教_"))
            {
                return false;
            }
            else
            {
                reg = new Regex("系统消息");
                m = reg.Match(Qname);
                if (m.Value.ToString().Equals("系统消息"))
                {
                    return false;
                }
                else return true;
            }
        }
        //向哈希表中插入数据
        public static void InsertQnum(string Qnum, Hashtable ht)
        {
            if (!ht.ContainsKey(Qnum))
            {
                ht.Add(Qnum, 1);
            }
            else
            {
                int i = (int)ht[Qnum];
                ht.Remove(Qnum);
                i = i + 1;
                ht.Add(Qnum, i);
            }
        }
        //给予客户端调用获得参与获奖的名单
        public static string[] GetArray()
        {
            int i = 0;
            string[] result = new string[ht.Count];
            foreach (String key in ht.Keys)
            {

                result[i] = key;
                i++;
            }
            return result;
        }
        //给予抽奖算法调用的可参与的人数
        public static int Getnum()
        {
            return ht.Count - 1;
        }
        //过滤信息并添加可用数据,客户端先行调用
        public static void filter(List<Message> messages, string Begingtime, string Endtime, string password)
        {

            foreach (var message in messages)
            {
                if (Ifnoassistant(message.GetQname()))
                {
                    if (Ifactivate(message.GetQtime(), Begingtime, Endtime))
                    {
                        if (Ifcorrespond(message.GetQcontent(), password))
                        {
                            InsertQnum(message.GetQnum(), ht);
                        }
                    }
                }
            }
        }

        public static string Sha256(string data)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(data);
            byte[] hash = SHA256Managed.Create().ComputeHash(bytes);
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < hash.Length; i++)
            {
                builder.Append(hash[i].ToString("X2"));
            }
            return builder.ToString();
        }
        public static List<long> GetWinners(int start, int end, int num_win, string key)
        {
            List<long> winners = new List<long>();
            string res = key;
            while (winners.Count < num_win)
            {
                res = Sha256(res);
                res = res.Substring(res.Length - 3);
                long r = (Convert.ToInt64(res, 16)) % (end - start) + start;
                if (winners.Contains(r)) continue;
                winners.Add(r);
            }
            return winners;
        }
        static public string GetBTC()
        {
            HttpWebRequest httpReq;
            HttpWebResponse httpResp;

            string strBuff = "";
            char[] cbuffer = new char[256];
            int byteRead = 0;

            Uri httpURL = new Uri("https://btc.com");
            httpReq = (HttpWebRequest)WebRequest.Create(httpURL);
            httpResp = (HttpWebResponse)httpReq.GetResponse();
            Stream respStream = httpResp.GetResponseStream();
            StreamReader respStreamReader = new StreamReader(respStream, Encoding.UTF8);
            byteRead = respStreamReader.Read(cbuffer, 0, 256);
            while (byteRead != 0)
            {
                string strResp = new string(cbuffer, 0, byteRead);
                strBuff = strBuff + strResp;
                byteRead = respStreamReader.Read(cbuffer, 0, 256);
            }
            respStream.Close();

            Regex No_a = new Regex("<a href=\"https://btc.com/[a-z0-9]{64}\"");
            string test = strBuff;
            foreach (Match m in No_a.Matches(test))
            {

                Regex No_b = new Regex("[a-z0-9]{64}");
                foreach (Match mm in No_b.Matches(m.Groups[0].Value))
                {
                    strBuff = mm.Groups[0].Value;
                }
                break;
            }
            return strBuff;
        }
    }
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
