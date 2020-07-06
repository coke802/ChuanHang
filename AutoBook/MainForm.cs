using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Threading;
using System.Diagnostics;
using System.IO;

namespace AutoBook
{
    public partial class MainForm : Form
    {

        [DllImport("urlmon.dll", CharSet = CharSet.Ansi)]
        private static extern int UrlMkSetSessionOption(int dwOption, string pBuffer, int dwBufferLength, int dwReserved);

        [DllImport("Kernel32.dll")]
        public static extern bool Beep(int frequency, int duration);

        const int URLMON_OPTION_USERAGENT = 0x10000001;

        bool stopped = true;
        bool injected = false;
        int step = 0;
        int sleep = 0;
        int beep_count = 0;

        public MainForm()
        {
            InitializeComponent();

            comboBox1.Text = "";
        }

        public static void ChangeUserAgent(string userAgent)
        {
            UrlMkSetSessionOption(URLMON_OPTION_USERAGENT, userAgent, userAgent.Length, 0);
        }

        static void SetWebBrowserFeatures()
        {
            // don't change the registry if running in-proc inside Visual Studio
            if (LicenseManager.UsageMode != LicenseUsageMode.Runtime)
                return;

            var appName = System.IO.Path.GetFileName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);

            var featureControlRegKey = @"HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\";

            Registry.SetValue(featureControlRegKey + "FEATURE_BROWSER_EMULATION",
                appName, GetBrowserEmulationMode(), RegistryValueKind.DWord);

            // enable the features which are "On" for the full Internet Explorer browser

            Registry.SetValue(featureControlRegKey + "FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_AJAX_CONNECTIONEVENTS",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_GPU_RENDERING",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_WEBOC_DOCUMENT_ZOOM",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_NINPUT_LEGACYMODE",
                appName, 0, RegistryValueKind.DWord);
        }

        static UInt32 GetBrowserEmulationMode()
        {
            int browserVersion = 0;
            using (var ieKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Internet Explorer",
                RegistryKeyPermissionCheck.ReadSubTree,
                System.Security.AccessControl.RegistryRights.QueryValues))
            {
                var version = ieKey.GetValue("svcVersion");
                if (null == version)
                {
                    version = ieKey.GetValue("Version");
                    if (null == version)
                        throw new ApplicationException("Microsoft Internet Explorer is required!");
                }
                int.TryParse(version.ToString().Split('.')[0], out browserVersion);
            }

            if (browserVersion < 7)
            {
                throw new ApplicationException("Unsupported version of Microsoft Internet Explorer!");
            }

            UInt32 mode = 11000; // Internet Explorer 11. Webpages containing standards-based !DOCTYPE directives are displayed in IE11 Standards mode. 

            switch (browserVersion)
            {
                case 7:
                    mode = 7000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE7 Standards mode. 
                    break;
                case 8:
                    mode = 8000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE8 mode. 
                    break;
                case 9:
                    mode = 9000; // Internet Explorer 9. Webpages containing standards-based !DOCTYPE directives are displayed in IE9 mode.                    
                    break;
                case 10:
                    mode = 10000; // Internet Explorer 10.
                    break;
            }

            return mode;
        }

        private void OnLoad(object sender, EventArgs e)
        {
            SetWebBrowserFeatures();
            //ChangeUserAgent("Dalvik/2.1.0 (Linux; U; Android 9; MI 8 Lite MIUI/V10.3.2.0.PDTCNXM)");
            //ChangeUserAgent("Mozilla/5.0 (iPhone; CPU iPhone OS 11_4_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/11.0 Mobile/15E148 Safari/604.1");
            webBrowser1.ScriptErrorsSuppressed = true;

            webBrowser1.Navigate("https://m.sichuanair.com/touch-webapp/home");
        }

        private void DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string url = e.Url.AbsoluteUri.ToString();
            UtilsLog.Log("DocumentCompleted url={0}", url);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Length == 0)
            {
                Form1 form1 = new Form1();
                form1.ShowDialog();
                return;
            }

            Run();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Stop();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://m.sichuanair.com/touch-webapp/home");
        }

        public void Run()
        {
            if (stopped)
            {
                stopped = false;
                Thread t = new Thread(RunLoop);
                t.Start();
            }
        }

        public void Stop()
        {
            if (!stopped)
            {
                stopped = true;
            }
        }

        public void RunLoopSleep(int timeout)
        {
            UtilsLog.Log("RunLoopSleep ENTER {0}", timeout);
            sleep += timeout;
        }

        public void RunLoopBegin()
        {
            step = 0;
            sleep = 0;
        }

        public void RunLoopEnd()
        {
        }

        public delegate void InvokeWebBrowser();
        public void RunLoop()
        {
            RunLoopBegin();

            while (!stopped)
            {
                if (sleep > 0)
                {
                    if (sleep > 1000)
                    {
                        sleep -= 1000;
                    }
                    else
                    {
                        sleep = 0;
                    }
                }
                else
                {
                    webBrowser1.BeginInvoke(new InvokeWebBrowser(RunLoopMainEx));
                }

                Thread.Sleep(1000);
            }

            RunLoopEnd();
        }

        private void RunLoopMainEx()
        {
            try
            {
                RunLoopMain();
            }
            catch (Exception e)
            {
                UtilsLog.Log("RunLoopMainEx EXCEPTION" + e.Message);
            }
        }

        private void RunLoopMain()
        {
            switch (step)
            {
                case 0:
                    RunLoopMainStep0();
                    break;
                case 1:
                    RunLoopMainStep1();
                    break;
                case 2:
                    RunLoopMainStep2();
                    break;
                case 3:
                    RunLoopMainStep3();
                    break;
                case 4:
                    RunLoopMainStep4();
                    break;
                case 5:
                    RunLoopMainStep5();
                    break;
                case 6:
                    RunLoopMainStep6();
                    break;
                case 7:
                    RunLoopMainStep7();
                    break;
                case 8:
                    RunLoopMainStep8();
                    break;
                case 9:
                    RunLoopMainStep9();
                    break;
                case 10:
                    RunLoopMainStep10();
                    break;
            }
        }


        private void ExecJsScript()
        {
            if (!injected)
            {
                string funcName = "testWebpack";
                string funcBody = "";

                funcBody += "var t = webpackJsonp([],{},[\"abGD\"]);";
                funcBody += "t.a.queryFlightCheckOrig = t.a.queryFlightCheck;";
                funcBody += "t.a.queryFlightCheck = function(e) {";
                funcBody += "return this.queryFlightCheckOrig({";
                funcBody += "flightSearchRequest: e";
                funcBody += "}).then(function (e) {";
                funcBody += "e.body.message = {};";
                funcBody += "e.body.message.keyCode = 0;";
                funcBody += "e.body.message.value = \"\";";
                funcBody += "return e";
                funcBody += "}).catch(function (e) {";
                funcBody += "})";
                funcBody += "}";

                ExecJSFunc(funcName, funcBody);

                injected = true;
            }
        }

        private void RunLoopMainStep0()
        {
            UtilsLog.Log("RunLoopMainStep0 ENTER");

            ExecJsScript();
            step++;
        }

        private void RunLoopMainStep1()
        {
            UtilsLog.Log("RunLoopMainStep1 ENTER");

            HtmlElement elemsearchFlightBox = webBrowser1.Document.GetElementById("searchFlightBox");
            if (elemsearchFlightBox == null)
            {
                return;
            }

            List<HtmlElement> elemsearchFlightBoxItems = new List<HtmlElement>();
            UtilsHtml.FindElementHtmlElements(elemsearchFlightBox, ":div/:div", elemsearchFlightBoxItems);
            foreach (HtmlElement elemsearchFlightBoxItem in elemsearchFlightBoxItems)
            {
                if (elemsearchFlightBoxItem.InnerText == "搜索机票")
                {
                    elemsearchFlightBoxItem.InvokeMember("click");
                    RunLoopSleep(2000);
                    step++;
                    break;
                }
            }
        }

        private void RunLoopMainStep2()
        {
            UtilsLog.Log("RunLoopMainStep2 ENTER");

            HtmlElement elemapp = webBrowser1.Document.GetElementById("app");
            if (elemapp != null)
            {
                List<HtmlElement> elemappItems = new List<HtmlElement>();
                UtilsHtml.FindElementHtmlElements(elemapp, ":div/:div", elemappItems);
                foreach (HtmlElement elemappItem in elemappItems)
                {
                    string elemappItemClassName = elemappItem.GetAttribute("className");
                    if (elemappItemClassName != null && elemappItemClassName.Contains("flight-content"))
                    {
                        step++;
                        break;
                    }
                }
            }

            /*
            List<HtmlElement> elemvuxconfirmItems = new List<HtmlElement>();
            UtilsHtml.FindElementHtmlElements(webBrowser1.Document.Body, ":div/:div/:div/:div/:a", elemvuxconfirmItems);
            foreach (HtmlElement elemvuxconfirmItem in elemvuxconfirmItems)
            {
                string elemvuxconfirmItemText = elemvuxconfirmItem.InnerText;
                UtilsLog.Log("RunLoopMainStep2 elemvuxconfirmItemText={0}", elemvuxconfirmItemText);
                if (elemvuxconfirmItemText == "确定")
                {
                    elemvuxconfirmItem.InvokeMember("click");
                    RunLoopSleep(2000);
                    break;
                }
            }
            */
        }

        private void RunLoopMainStep3()
        {
            UtilsLog.Log("RunLoopMainStep3 ENTER");

            HtmlElement elemapp = webBrowser1.Document.GetElementById("app");
            if (elemapp == null)
            {
                return;
            }

            HtmlElement elemFlightContent = FindFlightContent();
            if (elemFlightContent != null)
            {
                List<HtmlElement> elemFlightContentItems1 = new List<HtmlElement>();
                UtilsHtml.FindElementHtmlElements(elemFlightContent, ":ul/:li/:div/class:ticket-item", elemFlightContentItems1);
                if (elemFlightContentItems1.Count > 0)
                {
                    //成功
                    HtmlElement elemFlightContentItem1 = elemFlightContentItems1[0];
                    elemFlightContentItem1.Children[0].InvokeMember("click");
                    RunLoopSleep(2000);
                    step++;
                }
                else
                {
                    List<HtmlElement> elemFlightContentItems2 = new List<HtmlElement>();
                    UtilsHtml.FindElementHtmlElements(elemFlightContent, ":div/:div/:div/:div", elemFlightContentItems2);
                    foreach (HtmlElement elemFlightContentItem2 in elemFlightContentItems2)
                    {
                        if (elemFlightContentItem2.InnerText == "重新查询")
                        {
                            //失败
                            elemFlightContentItem2.InvokeMember("click");
                            RunLoopSleep(2000);
                            step = 1;
                            break;
                        }
                    }
                }
            }
        }

        private void RunLoopMainStep4()
        {
            UtilsLog.Log("RunLoopMainStep4 ENTER");

            HtmlElement elemapp = webBrowser1.Document.GetElementById("app");
            if (elemapp == null)
            {
                return;
            }

            List<HtmlElement> elemappItems = new List<HtmlElement>();
            UtilsHtml.FindElementHtmlElements(elemapp, ":div/:ol/:li/class:child-list-item/class:child-item", elemappItems);
            UtilsLog.Log("RunLoopMainStep4 elemappItems.Count={0}", elemappItems.Count);
            if (elemappItems.Count > 0)
            {
                foreach (HtmlElement elemappItem in elemappItems)
                {
                    if (elemappItem.Children.Count >= 2)
                    {
                        HtmlElement elemappItemChildItemLeft = elemappItem.Children[0];
                        HtmlElement elemappItemChildItemRight = elemappItem.Children[1];

                        string elemappItemChildItemLeftText = elemappItemChildItemLeft.InnerText;
                        UtilsLog.Log("RunLoopMainStep4 elemappItemChildItemLeftText={0}", elemappItemChildItemLeftText);
                        if (elemappItemChildItemLeftText.Contains(comboBox1.Text))
                        {
                            elemappItemChildItemRight.InvokeMember("click");
                            RunLoopSleep(2000);
                            step++;
                            break;
                        }
                    }
                }
            }
        }

        private void RunLoopMainStep5()
        {
            UtilsLog.Log("RunLoopMainStep5 ENTER");

            List<HtmlElement> elems = new List<HtmlElement>();
            UtilsHtml.FindWindowHtmlElementsByTagAndText(webBrowser1.Document.Window, "p", "确认并预订", elems);
            if (elems.Count > 0)
            {
                elems[0].InvokeMember("click");
                RunLoopSleep(2000);
                step++;
            }
        }

        private void RunLoopMainStep6()
        {
            UtilsLog.Log("RunLoopMainStep6 ENTER");

            HtmlElement elemapp = webBrowser1.Document.GetElementById("app");
            if (elemapp == null)
            {
                return;
            }

            List<HtmlElement> elemappItems = new List<HtmlElement>();
            UtilsHtml.FindElementHtmlElements(elemapp, ":div/:div/:div/:div/class:passenger-list/:li", elemappItems);
            UtilsLog.Log("RunLoopMainStep6 elemappItems.Count={0}", elemappItems.Count);
            if (elemappItems.Count > 0)
            {
                elemappItems[0].Focus();
                elemappItems[0].InvokeMember("click");
                RunLoopSleep(2000);
                step++;
            }
        }

        private void RunLoopMainStep7()
        {
            UtilsLog.Log("RunLoopMainStep7 ENTER");

            HtmlElement elemapp = webBrowser1.Document.GetElementById("app");
            if (elemapp == null)
            {
                return;
            }

            List<HtmlElement> elemappItems = new List<HtmlElement>();
            UtilsHtml.FindElementHtmlElements(elemapp, ":div/:div/:div/:div/:div/class:switch-btn", elemappItems);
            UtilsLog.Log("RunLoopMainStep7 elemappItems.Count={0}", elemappItems.Count);
            if (elemappItems.Count > 0)
            {
                elemappItems[0].Focus();
                elemappItems[0].InvokeMember("click");
                RunLoopSleep(2000);
                step++;
            }
        }

        private void RunLoopMainStep8()
        {
            UtilsLog.Log("RunLoopMainStep8 ENTER");

            List<HtmlElement> elems = new List<HtmlElement>();
            UtilsHtml.FindWindowHtmlElementsByTagAndText(webBrowser1.Document.Window, "div", "提交订单", elems);
            if (elems.Count > 0)
            {
                elems[0].Focus();
                elems[0].InvokeMember("click");
                RunLoopSleep(2000);
                beep_count = 5;
                step++;
            }
        }

        private void RunLoopMainStep9()
        {
            UtilsLog.Log("RunLoopMainStep9 ENTER");
            if (beep_count == 0)
            {
                step++;
                return;
            }

            Beep(500, 700);
            beep_count--;
        }

        private void RunLoopMainStep10()
        {
            UtilsLog.Log("RunLoopMainStep10 ENTER");
        }


        private HtmlElement FindFlightContent()
        {
            HtmlElement elemapp = webBrowser1.Document.GetElementById("app");
            if (elemapp != null)
            {
                List<HtmlElement> elemappItems = new List<HtmlElement>();
                UtilsHtml.FindElementHtmlElements(elemapp, ":div/:div", elemappItems);
                foreach (HtmlElement elemappItem in elemappItems)
                {
                    string elemappItemClassName = elemappItem.GetAttribute("className");
                    if (elemappItemClassName != null && elemappItemClassName.Contains("flight-content"))
                    {
                        return elemappItem;
                    }
                }
            }

            return null;
        }

        public void ExecJSFunc(string funcName, string funcBody)
        {
            UtilsLog.Log("ExecJSFunc ENTER {0}", funcName);

            HtmlDocument document = webBrowser1.Document;
            string funcText = "function " + funcName + "(){" + funcBody + "}";
            HtmlElement elementScript = document.CreateElement("script");
            elementScript.SetAttribute("type", "text/javascript");
            elementScript.SetAttribute("text", funcText);
            HtmlElement elementHead = document.GetElementsByTagName("head")[0];
            elementHead.AppendChild(elementScript);
            document.InvokeScript(funcName);
        }
    }
}
