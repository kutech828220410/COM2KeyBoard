using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Threading;
using Basic;
using System.Windows.Forms;
namespace COM2KeyBoard
{
   
    class Program
    {
        #region MyConfigClass
        public static string currentDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        private static string MyConfigFileName = $@"{currentDirectory}\MyConfig.txt";
        static public MyConfigClass myConfigClass = new MyConfigClass();
        public class MyConfigClass
        {
            private string scanner01_COMPort = "COM2";
            public string Scanner01_COMPort { get => scanner01_COMPort; set => scanner01_COMPort = value; }

        }
        static private void LoadMyConfig()
        {
            string jsonstr = MyFileStream.LoadFileAllText($"{MyConfigFileName}");
            if (jsonstr.StringIsEmpty())
            {
                jsonstr = Basic.Net.JsonSerializationt<MyConfigClass>(new MyConfigClass(), true);
                List<string> list_jsonstring = new List<string>();
                list_jsonstring.Add(jsonstr);
                if (!MyFileStream.SaveFile($"{MyConfigFileName}", list_jsonstring))
                {
                    Console.WriteLine($"建立{MyConfigFileName}檔案失敗!");
                }
                Console.WriteLine($"未建立參數文件!請至子目錄設定{MyConfigFileName}");
                System.Threading.Thread.Sleep(2000);
                Environment.Exit(0);
            }
            else
            {
                myConfigClass = Basic.Net.JsonDeserializet<MyConfigClass>(jsonstr);

                jsonstr = Basic.Net.JsonSerializationt<MyConfigClass>(myConfigClass, true);
                List<string> list_jsonstring = new List<string>();
                list_jsonstring.Add(jsonstr);
                if (!MyFileStream.SaveFile($"{MyConfigFileName}", list_jsonstring))
                {
                    Console.WriteLine($"建立{MyConfigFileName}檔案失敗!");
                }

            }

        }
        #endregion
        [STAThread]
        static void Main(string[] args)
        {
            bool createdNew;
            Mutex mutex = new Mutex(true, "COM2KeyBoard", out createdNew);

            if (!createdNew)
            {
                Console.WriteLine("程序已執行，無法重複執行。");
                return;
            }
            LoadMyConfig();
            MySerialPort mySerialPort = new MySerialPort();
            mySerialPort.Init(myConfigClass.Scanner01_COMPort, 9600, 8, System.IO.Ports.Parity.None, System.IO.Ports.StopBits.One);
            mySerialPort.SerialPortOpen();
            if(mySerialPort.IsConnected == false)
            {
                Console.WriteLine($"{myConfigClass.Scanner01_COMPort} 開啟失敗!!");
                System.Threading.Thread.Sleep(2000);
                Environment.Exit(0);
            }
            Console.WriteLine($"{myConfigClass.Scanner01_COMPort} 開啟成功!!");
            while (true)
            {
                string text = mySerialPort.ReadString();
                if(text.StringIsEmpty() == false)
                {
                    try
                    {
                        System.Threading.Thread.Sleep(100);
                        text = mySerialPort.ReadString();
                        mySerialPort.ClearReadByte();
                        text = text.Replace("\0", "");
                        Console.WriteLine($"接收到字串: {text}");
                        try
                        {
                            //SendKeys.SendWait(text);
                            Clipboard.SetText(text);
                            SendKeys.SendWait("^V");
                            Console.WriteLine($"發送鍵盤模擬: {text}");
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"異常: {e.Message}");
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"異常: {e.Message}");
                    }
                  
     

                }
                System.Threading.Thread.Sleep(10);
            }
        }
    }
}
