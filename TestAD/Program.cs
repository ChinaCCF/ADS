using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
//using System.ServiceProcess;
using System.Text;
using System.Xml;
using System.Web.Script.Serialization;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Net;
using System.IO;
using System.DirectoryServices;

namespace ADS
{
    class Config
    {
        static Config cfg_ = null;

        public string adhost;
        public string adname;
        public string adpwd;
        public string port;
        public static Config sharedInstance()
        {
            if (cfg_ == null)
            {
                cfg_ = new Config();
                string path = AppDomain.CurrentDomain.BaseDirectory;

                XmlDocument xml = new XmlDocument();
                xml.Load(path + "cfg.xml");
                XmlNode root = xml.SelectSingleNode("root");
                XmlElement node = root["ADHost"];
                cfg_.adhost = node.InnerText;
                node = root["ADName"];
                cfg_.adname = node.InnerText;
                node = root["ADPwd"];
                cfg_.adpwd = node.InnerText;
                node = root["Port"];
                cfg_.port = node.InnerText;
            }
            return cfg_;
        }
    }
    class Json
    {
        static public Dictionary<string, object> toObject(string str)
        {
            JavaScriptSerializer js = new JavaScriptSerializer();
            try
            {
                return js.Deserialize<Dictionary<string, object>>(str);
            }
            catch (Exception)
            {
                return null;
            }
        }

        static public string toString(Dictionary<string, object> dic)
        {
            JavaScriptSerializer js = new JavaScriptSerializer();
            try
            {
                return js.Serialize(dic);
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
    class IPTool
    {
        static public List<string> getIPs()
        {
            List<string> re = new List<string>();
            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
            foreach (NetworkInterface adapter in adapters)
            {
                if (adapter.NetworkInterfaceType == NetworkInterfaceType.Ethernet)
                {
                    IPInterfaceProperties ips = adapter.GetIPProperties();
                    UnicastIPAddressInformationCollection ipCollection = ips.UnicastAddresses;
                    foreach (UnicastIPAddressInformation ip in ipCollection)
                    {
                        if (ip.Address.AddressFamily == AddressFamily.InterNetwork)
                            re.Add(ip.Address.ToString());
                    }
                }
            }
            return re;
        }
    }
    class HTTP
    {
        HttpListener listerner = null;
        bool isRun = false;
        public HTTP()
        {
            listerner = new HttpListener();
        }
        public void setPort(int port)
        {
            listerner.Prefixes.Clear();

            List<string> ips = IPTool.getIPs();
            ips.Add("localhost");
            ips.Add("127.0.0.1");
            foreach (string ip in ips)
                listerner.Prefixes.Add(string.Format("http://{0}:{1}/", ip, port.ToString()));
        }
        public void start()
        {
            listerner.Start();
            isRun = true;
            listerner.BeginGetContext(new AsyncCallback(worker), null);
        }

        public void stop()
        {
            isRun = false;
            listerner.Stop();
        }

        public void worker(IAsyncResult para)
        {
            HttpListenerContext ctx = listerner.EndGetContext(para);
            if (isRun)
                listerner.BeginGetContext(new AsyncCallback(worker), null);

            Worker w = new Worker(ctx.Request.InputStream, ctx.Response.OutputStream);
            ctx.Response.StatusCode = w.doWork();
            w.close();
            ctx.Response.Close();
        }
    }
    enum ADType
    {
        User,
        OU,
        Unknow
    }
    class ADItem
    {
        static string ADUser = "user";
        static string ADOU = "organizationalUnit";
        static string ADLoginName = "sAMAccountName";
        static string ADUserCTRL = "userAccountControl";
        static public string itemString(ADType t)
        {
            if (t == ADType.OU)
                return ADOU;
            else
            {
                if (t == ADType.User)
                    return ADUser;
                return "";
            }
        }
        static public string itemPrefix(ADType t)
        {
            if (t == ADType.OU)
                return "OU";
            else
            {
                if (t == ADType.User)
                    return "CN";
                return "";
            }
        }
        static public ADType itemType(string str)
        {
            str = str.Trim();
            if (string.Compare(str, ADUser) == 0)
                return ADType.User;
            if (string.Compare(str, ADOU) == 0)
                return ADType.OU;

            return ADType.Unknow;
        }



        DirectoryEntry de_;
        public ADItem(DirectoryEntry de)
        {
            de_ = de;
        }
        public ADItem parent()
        {
            return new ADItem(de_.Parent);
        }
        public ADType type()
        {
            Dictionary<string, string> objClass = this.getProperty("objectClass");
            foreach (string key in objClass.Keys)
            {
                ADType t = itemType(key);
                if (t != ADType.Unknow)
                    return t;
            }
            return ADType.Unknow;
        }
        public string name()
        {
            string subStr = null;
            if (type() == ADType.OU)
                subStr = "OU=";
            else
                subStr = "CN=";
            return de_.Name.Substring(subStr.Length, de_.Name.Length - subStr.Length);
        }
        public ADItem createChild(string name, ADType t)
        {
            name = name.Trim();
            ADItem item = this.getChild(name);
            if (item != null)
                return null;
            DirectoryEntry de = de_.Children.Add(itemPrefix(t) + "=" + name, itemString(t));
            if (de == null)
                return null;
            try
            {
                de.CommitChanges();
            }
            catch (Exception)
            {
                return null;
            }
            return new ADItem(de);
        }
        public bool deleteSelf()
        {
            if (type() == ADType.OU)
            {
                foreach (DirectoryEntry de in de_.Children)
                {
                    ADItem item = new ADItem(de);
                    if (!item.deleteSelf())
                        return false;
                }
            }

            try
            {
                DirectoryEntry p = de_.Parent;
                p.Children.Remove(de_);
                p.CommitChanges();
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
        public ADItem getChild(string name)
        {
            name = name.Trim();
            foreach (DirectoryEntry de in de_.Children)
            {
                if (string.Compare(new ADItem(de).name(), name) == 0)
                    return new ADItem(de);
            }
            return null;
        }
        public Dictionary<string, string> getProperty(string key)
        {
            key = key.Trim();
            Dictionary<string, string> vals = new Dictionary<string, string>();
            for (int i = 0; i < de_.Properties[key].Count; ++i)
            {
                object obj = de_.Properties[key][i];
                try
                {
                    if (obj.GetType() == typeof(string))
                        vals.Add((string)obj, "string");
                    else
                        vals.Add(obj.ToString(), obj.GetType().ToString());
                }
                catch (Exception)
                {
                    return null;
                }
            }
            return vals;
        }
        public Dictionary<string, object> enumProperty()
        {
            Dictionary<string, object> di = new Dictionary<string, object>();
            foreach (string key in de_.Properties.PropertyNames)
                di.Add(key, getProperty(key));
            
            return di;
        }
        public List<object> enumItem(bool recursion = false)
        {
            List<object> list = new List<object>();
            foreach (DirectoryEntry de in de_.Children)
            {
                ADItem item = new ADItem(de);
                if (recursion)
                {
                    if (item.type() == ADType.OU)
                    {
                        Dictionary<string, object> dic = new Dictionary<string, object>();
                        dic.Add(item.name(), item.enumItem(recursion));
                        list.Add(dic);
                    }
                    else
                        list.Add(item.name());
                }
                else
                    list.Add(item.name());
            }
            return list;
        }
        public bool setProperty(string p, string val)
        {
            p = p.Trim();
            val = val.Trim();

            if (string.IsNullOrWhiteSpace(p) || string.IsNullOrWhiteSpace(val))
                return false;

            try
            {
                de_.Properties[p].Add(val);
                de_.CommitChanges();
            }
            catch (Exception)
            {
                try
                {
                    de_.Properties[p].Clear();
                    de_.Properties[p].Add(val);
                    de_.CommitChanges();
                }
                catch (Exception)
                {
                    return false;
                }
            }
            return true;
        }
        public string changeUserPwd(string pwd)
        {
            pwd = pwd.Trim();
            if (string.IsNullOrEmpty(pwd))
                return "change user pwd : password can't be empty!";

            ADType adt = type();
            if (adt != ADType.User)
                return "change user pwd : no user item!";
            try
            {
                de_.Invoke("SetPassword", new object[] { pwd });
                de_.CommitChanges();
            }
            catch (Exception)
            {
                return "change user pwd : fail!";
            }
            return null;
        }
        public string changeUserLoginName(string name)
        {
            name = name.Trim();
            if (string.IsNullOrEmpty(name))
                return "change user login name : name can't be empty!";

            ADType adt = type();
            if (adt != ADType.User)
                return "change user login name : no user item!";
            if (setProperty(ADLoginName, name))
                return null;
            return "change user login name : fail!";
        }
        public string enableUser(bool isEnable)
        {
            int val = (int)de_.Properties[ADUserCTRL].Value;
            //val |= 0x1000;//密码永不过期
            //de_.CommitChanges();

            if (isEnable)
                val &= ~0x2;//启用账号            
            else
                val |= 0x2;//启用账号            

            try
            {
                de_.Properties[ADUserCTRL].Value = val;
                de_.CommitChanges();
            }
            catch (Exception)
            {
                return "enable user : fail1";
            }
            return null;
        }
    }
    class AD
    {
        static AD ad_ = null;

        DirectoryEntry root_ = null;
        public static AD sharedInstance()
        {
            if (ad_ == null)
            {
                Config cfg = Config.sharedInstance();
                ad_ = new AD();
                ad_.connect(cfg.adhost.Trim(), cfg.adname.Trim(), cfg.adpwd.Trim());
            }
            return ad_;
        }
        public ADItem root()
        {
            return new ADItem(root_);
        }
        public bool connect(string host, string user, string pwd)
        {
            try
            {
                DirectoryEntry de = new DirectoryEntry("LDAP://" + host, user, pwd);
                DirectorySearcher se = new DirectorySearcher(de);
                //1:字符串必须在括号内,例如:(objectClass=name)
                //2:运算符为 <, <=, =, =>, >
                //3:符合表达式前面需要带&或者|
                se.Filter = "(&(objectClass=user)(cn=Guest))";
                SearchResult re = se.FindOne();
                if (re != null)
                    root_ = re.GetDirectoryEntry().Parent.Parent;
            }
            finally
            { }
            return root_ != null;
        }
        public ADItem lastItem(string path)
        {
            path = path.Trim();
            if (path[0] == '/')
                path = path.Substring(1, path.Length - 1);
            string[] arr = path.Split('/');
            int cnt = arr.Length;
            if (cnt == 1)
            {
                string str = arr[0].Trim();
                if (string.IsNullOrWhiteSpace(str))
                    return root();
            }

            int index = 0;
            ADItem item = root();
            while (index < cnt && item != null)
                item = item.getChild(arr[index++].Trim());

            return item;
        }
        public List<object> enumItem(string path, bool recursion = false)
        {
            ADItem item = lastItem(path);
            if (item == null)
                return null;
            return item.enumItem(recursion);
        }
        public string deleteItem(string path)
        {
            ADItem item = lastItem(path);
            if (item == null)
                return "delete item : path is error!";

            if (item.deleteSelf())
                return null;
            return "delete item : unknow error!";
        }
        public string createUser(string path, string name, string displayName)
        {
            ADItem parent = lastItem(path);
            if (parent == null)
                return "create user : path is error!";
            ADItem item = parent.createChild(name, ADType.User);
            if (item == null)
                return "create user : unknow error!";

            return item.changeUserLoginName(displayName);
        }
        public string createOU(string path, string name)
        {
            ADItem parent = lastItem(path);
            if (parent == null)
                return "create ou : path is error!";
            ADItem item = parent.createChild(name, ADType.OU);
            if (item == null)
                return "create ou : unknow error!";
            return null;
        }
        public string changeUserPwd(string path, string pwd)
        {
            ADItem item = lastItem(path);
            if (item == null)
                return "change user password : path is error!";
            return item.changeUserPwd(pwd);
        }
        public string enableUser(string path, string val)
        {
            ADItem item = lastItem(path);
            if (item == null)
                return "enable user : path is error!";

            bool v = true;
            if (string.Compare(val.Trim(), "0") == 0)
                v = false;
            return item.enableUser(v);
        }
        public object itemInfo(string path)
        {
            ADItem item = lastItem(path);
            if (item == null)
                return "item info : path is error!";

            return item.enumProperty();
        }
    }
    class Worker
    {
        StreamReader r_;
        StreamWriter w_;

        public Worker(Stream r, Stream w)
        {
            w_ = new StreamWriter(w);
            r_ = new StreamReader(r);
        }

        public int doWork()
        {
            if (r_.EndOfStream)
                return result(400, "content is empty!");

            string body = r_.ReadToEnd();

            Dictionary<string, object> req = Json.toObject(body);
            if (req == null)
                return result(400, "error content!");

            if (!req.ContainsKey("api"))
                return result(400, "miss api member!");
            string api = (string)req["api"];
            api = api.Trim();

            if (!req.ContainsKey("paras"))
                return result(400, "miss paras member!");
            Dictionary<string, object> paras = (Dictionary<string, object>)req["paras"];

            if (!paras.ContainsKey("path"))
                return result(400, "miss path member!");
            string path = (string)paras["path"];
            path = path.Trim();

            if (string.Compare(api, "enumItem") == 0)
            {
                bool recursion = true;
                if (paras.ContainsKey("recursion"))
                {
                    string r = (string)paras["recursion"];
                    if (string.Compare(r.Trim(), "0") == 0)
                        recursion = false;
                }

                AD ad = AD.sharedInstance(); 
                return result(200, "success!", ad.enumItem(path, recursion));
            }

            if (string.Compare(api, "deleteItem") == 0)
            {
                AD ad = AD.sharedInstance();
                string str = ad.deleteItem(path);
                if (str == null)
                    return result(200, "success!");
                return result(500, str);
            }

            if (string.Compare(api, "createUser") == 0)
            {
                if (!paras.ContainsKey("name"))
                    return result(400, "miss name member!");
                if (!paras.ContainsKey("displayName"))
                    return result(400, "miss displayName member!");
                string name = (string)paras["name"];
                string displayName = (string)paras["displayName"];
                AD ad = AD.sharedInstance();
                string str = ad.createUser(path.Trim(), name.Trim(), displayName.Trim());
                if (str == null)
                    return result(200, "success!");
                return result(500, str);
            }

            if (string.Compare(api, "createOU") == 0)
            {
                if (!paras.ContainsKey("name"))
                    return result(400, "miss name member!");

                string name = (string)paras["name"];

                AD ad = AD.sharedInstance();
                string str = ad.createOU(path.Trim(), name.Trim());
                if (str == null)
                    return result(200, "success!");
                return result(500, str);
            }

            if (string.Compare(api, "changeUserPwd") == 0)
            {
                if (!paras.ContainsKey("pwd"))
                    return result(400, "miss pwd member!");

                string pwd = (string)paras["pwd"];

                AD ad = AD.sharedInstance();
                string str = ad.changeUserPwd(path.Trim(), pwd.Trim());
                if (str == null)
                    return result(200, "success!");
                return result(500, str);
            }

            if (string.Compare(api, "enableUser") == 0)
            {
                if (!paras.ContainsKey("enable"))
                    return result(400, "miss enable member!");

                string enable = (string)paras["enable"];

                AD ad = AD.sharedInstance();
                string str = ad.enableUser(path.Trim(), enable.Trim());
                if (str == null)
                    return result(200, "success!");
                return result(500, str);
            }

            if (string.Compare(api, "itemInfo") == 0)
            {
                AD ad = AD.sharedInstance();
                return result(200, "success!", ad.itemInfo(path));
            }

            return result(400, "error api request!");
        }
        public int result(int code, string msg, object obj = null)
        {
            Dictionary<string, object> dic = new Dictionary<string, object>();
            dic.Add("code", code);
            dic.Add("msg", msg);
            dic.Add("result", obj);
            w_.WriteLine(Json.toString(dic));
            return code;
        }
        public void close()
        {
            r_.Close();
            w_.Close();
        }
    }
 
    class Program
    {
        static Config cfg = null;
        static HTTP http = null;
        static void Main(string[] args)
        {
            cfg = Config.sharedInstance();
            http = new HTTP();
            http.setPort(int.Parse(cfg.port));
            http.start();

            while(true){}
        }
    }
}

 