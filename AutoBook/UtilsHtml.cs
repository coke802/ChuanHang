using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using mshtml;

namespace AutoBook
{
    class UtilsHtml
    {
        public static void PrintFrameWindowURLFrame(int index, HtmlWindow frame)
        {
            try
            {
                UtilsLog.Log("PrintFrameWindowURLFrame [" + index.ToString() + "]=" + frame.Document.Url.ToString());
            }
            catch (Exception e)
            {
                //UtilsLog.Log("PrintFrameWindowURLFrame EXCEPTION " + e.Message);
            }
        }

        public static void PrintFrameWindowURL(WebBrowser web)
        {
            HtmlWindow webWindow = web.Document.Window;
            UtilsLog.Log("PrintFrameWindowURL ENTER Count= " + webWindow.Frames.Count);
            for (int i = 0; i < webWindow.Frames.Count; i++)
            {
                PrintFrameWindowURLFrame(i, webWindow.Frames[i]);
            }
        }

        public static void PrintFrameWindowNameFrame(int index, HtmlWindow frame)
        {
            try
            {
                UtilsLog.Log("PrintFrameWindowNameFrame [" + index.ToString() +  "]=" + frame.Name);
            }
            catch (Exception e)
            {
                //UtilsLog.Log("PrintFrameWindowNameFrame EXCEPTION " + e.Message);
            }
        }

        public static void PrintFrameWindowName(WebBrowser web)
        {
            HtmlWindow webWindow = web.Document.Window;
            UtilsLog.Log("PrintFrameWindowName ENTER Count= " + webWindow.Frames.Count);
            for (int i = 0; i < webWindow.Frames.Count; i++)
            {
                PrintFrameWindowNameFrame(i, webWindow.Frames[i]);
            }
        }

        public static bool FindFrameWindowFrame(HtmlWindow frame, string name)
        {
            bool result = false;
            try
            {
                if (frame.Name == name)
                {
                    result = true;
                }
 
            }
            catch (Exception e)
            {
                //UtilsLog.Log("FindFrameWindowFrame EXCEPTION " + e.Message);
            }

            return result;
        }

        public static HtmlWindow FindFrameWindow(WebBrowser web, string name)
        {
            //PrintFrameWindowURL(web);
            //PrintFrameWindowName(web);

            HtmlWindow result = null;
            HtmlWindow webWindow = web.Document.Window;
            foreach (HtmlWindow frame in webWindow.Frames)
            {
                if (FindFrameWindowFrame(frame, name))
                {
                    result = frame;
                    break;
                }
            }

            return result;
        }

        public static HtmlWindow FindFrameWindow2(WebBrowser web, string name1, string name2)
        {
            HtmlWindow result = null;
            HtmlWindow frame1 = FindFrameWindow(web, name1);
            if (frame1 != null)
            {
                foreach (HtmlWindow frame2 in frame1.Frames)
                {
                    if (FindFrameWindowFrame(frame2, name2))
                    {
                        result = frame2;
                        break;
                    }
                }
            }

            return result;
        }

        public static HtmlWindow FindFrameWindow3(WebBrowser web, string name1, string name2, string name3)
        {
            HtmlWindow result = null;
            HtmlWindow frame1 = FindFrameWindow2(web, name1, name2);
            if (frame1 != null)
            {
                foreach (HtmlWindow frame2 in frame1.Frames)
                {
                    if (FindFrameWindowFrame(frame2, name3))
                    {
                        result = frame2;
                        break;
                    }
                }
            }

            return result;
        }

        public static string GB2312ToUtf8(string gb2312String)
        {
            Encoding fromEncoding = Encoding.GetEncoding("gb2312");
            Encoding toEncoding = Encoding.UTF8;
            return EncodingConvert(gb2312String, fromEncoding, toEncoding);
        }

        public static string Utf8ToGB2312(string utf8String)
        {
            Encoding fromEncoding = Encoding.UTF8;
            Encoding toEncoding = Encoding.GetEncoding("gb2312");
            return EncodingConvert(utf8String, fromEncoding, toEncoding);
        }

        public static string EncodingConvert(string fromString, Encoding fromEncoding, Encoding toEncoding)
        {
            byte[] fromBytes = fromEncoding.GetBytes(fromString);
            byte[] toBytes = Encoding.Convert(fromEncoding, toEncoding, fromBytes);
            string toString = toEncoding.GetString(toBytes);
            return toString;
        }

        public static void SaveHtmlWindowHtmlSource(HtmlWindow window, string fileName)
        {
            UtilsLog.Log("SaveHtmlWindowHtmlSource ENTER " + fileName);
            try
            {
                FileStream fs = new FileStream(fileName, FileMode.Create);
                string source = Utf8ToGB2312(window.Document.Body.OuterHtml);
                byte[] bytes = Encoding.UTF8.GetBytes(source);
                fs.Write(bytes, 0, bytes.Length);
                fs.Close();
            }
            catch (Exception e)
            {
                UtilsLog.Log("SaveHtmlWindowHtmlSource EXCEPTION " + e.Message);
            }
        }

        public static void SaveWebBrowserHtmlSource(WebBrowser web)
        {
            string[] urlNameLists = web.Document.Url.AbsoluteUri.Split('/');
            string urlLastName = urlNameLists[urlNameLists.Length - 1];

            string[] urlTitleLists = urlLastName.Split('.');
            string urlTitleName = urlTitleLists[0];

            DateTime dateTime = DateTime.Now;
            string fileDir = System.Environment.CurrentDirectory;
            UtilsLog.Log("SaveWebBrowserHtmlSource fileDir=" + fileDir);

            string saveDir = fileDir + "\\html\\" + dateTime.Year.ToString("D4") + dateTime.Month.ToString("D2") + dateTime.Day.ToString("D2")
                + "_" + dateTime.Hour.ToString("D2") + dateTime.Minute.ToString("D2") + dateTime.Second.ToString("D2");
            Directory.CreateDirectory(saveDir);
            UtilsLog.Log("SaveWebBrowserHtmlSource saveDir=" + saveDir);

            string fileName = saveDir + "\\" +  urlTitleName + ".txt";
            SaveHtmlWindowHtmlSource(web.Document.Window, fileName);

            HtmlWindow webWindow = web.Document.Window;
            foreach (HtmlWindow frame in webWindow.Frames)
            {
                try
                {
                    string frameFileName = saveDir + "\\Frame_" + frame.Name + ".txt";
                    SaveHtmlWindowHtmlSource(frame, frameFileName);
                }
                catch (Exception e)
                {
                    UtilsLog.Log("SaveWebBrowserHtmlSource EXCEPTION " + e.Message);
                }
            }
        }

        public static void PrintWebBrowserHtmlTitle(WebBrowser web)
        {
            UtilsLog.Log("PrintWebBrowserHtmlTitle ReadyState={0} StatusText={1} DocumentTitle={2}" + web.ReadyState + web.StatusText + web.DocumentTitle);
        }

        public static HtmlElement FindChildElementByText(HtmlElement parent, string text)
        {
            foreach (HtmlElement elementChild in parent.Children)
            {
                if (elementChild.InnerText.Contains(text))
                {
                    return elementChild;
                }
            }

            return null;
        }

        public static HtmlElement FindChildElementByText2(HtmlElement parent, string text)
        {
            foreach (HtmlElement elementChild in parent.Children)
            {
                foreach (HtmlElement elementChild2 in elementChild.Children)
                {
                    if (elementChild2.InnerText.Contains(text))
                    {
                        return elementChild2;
                    }
                }
            }

            return null;
        }

        public static HtmlElement FindChildElementByText3(HtmlElement parent, string text)
        {
            foreach (HtmlElement elementChild in parent.Children)
            {
                foreach (HtmlElement elementChild2 in elementChild.Children)
                {
                    foreach (HtmlElement elementChild3 in elementChild2.Children)
                    {
                        if (elementChild3.InnerText.Contains(text))
                        {
                            return elementChild3;
                        }
                    }
                }
            }

            return null;
        }

        public static bool CompareElement(HtmlElement element, string label)
        {
            bool result = false;
            try
            {
                string[] keyValue = label.Split(':');

                string key = keyValue[0];
                string value = keyValue[1];
                //UtilsLog.Log("CompareElement key=" + key + " value=" + value + " TagName=" + element.TagName);

                if (key.Length == 0)
                {
                    if (string.Compare(element.TagName, value, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        result = true;
                    }
                }
                else
                {
                    string attribute = "";
                    if (key == "class")
                    {
                        attribute = element.GetAttribute("className");
                    }
                    else
                    {
                        attribute = element.GetAttribute(key);
                    }

                    if (attribute == value)
                    {
                        result = true;
                    }
                }
            }
            catch (Exception e)
            {
                result = false;
            }

            return result;
        }

        public static bool CompareElementText(HtmlElement element, string text)
        {
            bool result = false;
            try
            {
                if (text == "*" || element.InnerText.Contains(text))
                {
                    result = true;
                }
            }
            catch (Exception e)
            {
                result = false;
            }

            return result;
        }

        public static void FindElementHtmlElements(HtmlElement parent, string path, List<HtmlElement> elements)
        {
            //UtilsLog.Log("FindElementHtmlElements ENTER {0}", path);

            string currLabel = null;
            string nextLabel = null;

            int pos = path.IndexOf('/', 0, path.Length);
            if (pos == -1)
            {
                currLabel = path;
                foreach (HtmlElement elementChild1 in parent.Children)
                {
                    if (CompareElement(elementChild1, currLabel))
                    {
                        elements.Add(elementChild1);
                    }
                }

                return;
            }

            currLabel = path.Substring(0, pos);
            pos++;
            nextLabel = path.Substring(pos, path.Length - pos);
            foreach (HtmlElement elementChild2 in parent.Children)
            {
                if (CompareElement(elementChild2, currLabel))
                {
                    FindElementHtmlElements(elementChild2, nextLabel, elements);
                }
            }
        }

        public static void FindWindowHtmlElements(HtmlWindow window, string path, List<HtmlElement> elements)
        {
            FindElementHtmlElements(window.Document.Body, path, elements);
        }

        public static void FindWindowHtmlElementsByTagAndText(HtmlWindow window, string tagName, string text, List<HtmlElement> elements)
        {
            foreach (HtmlElement elementChild in window.Document.All)
            {
                if (CompareElement(elementChild,  ":" + tagName))
                {
                    if (CompareElementText(elementChild, text))
                    {
                        elements.Add(elementChild);
                    }
                }
            }
        }

        public static void FindWindowHtmlElementsByTagAndName(HtmlWindow window, string tagName, string name, List<HtmlElement> elements)
        {
            foreach (HtmlElement elementChild in window.Document.All)
            {
                if (CompareElement(elementChild, ":" + tagName))
                {
                    if (CompareElement(elementChild, "name:" + name))
                    {
                        elements.Add(elementChild);
                    }
                }
            }
        }

        public static void FindWindowHtmlElementsByClassAndText(HtmlWindow window, string className, string text, List<HtmlElement> elements)
        {
            foreach (HtmlElement elementChild in window.Document.All)
            {
                if (CompareElement(elementChild, "class:" + className))
                {
                    if (CompareElementText(elementChild, text))
                    {
                        elements.Add(elementChild);
                    }
                }
            }
        }

        public static void FindWindowHtmlElementsByClassAndName(HtmlWindow window, string className, string name, List<HtmlElement> elements)
        {
            foreach (HtmlElement elementChild in window.Document.All)
            {
                if (CompareElement(elementChild, "class:" + className))
                {
                    if (CompareElement(elementChild, "name:" + name))
                    {
                        elements.Add(elementChild);
                    }
                }
            }
        }

        public static void FindWindowHtmlElementsByIdText(HtmlWindow window, string idName, List<HtmlElement> elements)
        {
            foreach (HtmlElement elementChild in window.Document.All)
            {
                if (elementChild.Id != null)
                {
                    if (elementChild.Id.Contains(idName))
                    {
                        elements.Add(elementChild);
                    }
                }
            }
        }

        public static void FindWindowHtmlElementsByValueText(HtmlWindow window, string text, List<HtmlElement> elements)
        {
            foreach (HtmlElement elementChild in window.Document.All)
            {
                if (CompareElement(elementChild, "value:" + text))
                {
                    elements.Add(elementChild);
                }
            }
        }

        public static void ParseElementStyle(HtmlElement element, Dictionary<string, string> elementStyles)
        {
            if (element.Style != null)
            {
                string[] styleItems1 = element.Style.Split(';');
                foreach (string styleItem1 in styleItems1)
                {
                    string[] styleItems2 = styleItem1.Split(':');
                    if (styleItems2.Length != 2)
                    {
                        continue;
                    }

                    string attrName = styleItems2[0].Replace(" ", "").ToLower();
                    string attrValue = styleItems2[1].Replace(" ", "").ToLower();
                    //UtilsLog.Log("ParseElementStyle attrName={0} attrValue={1}", attrName, attrValue);

                    elementStyles[attrName] = attrValue;
                }
            }
        }

        public static string GetElementStyleByKey(HtmlElement element, string key)
        {
            if (element != null)
            {
                Dictionary<string, string> elementStyles = new Dictionary<string, string>();
                ParseElementStyle(element, elementStyles);

                if (elementStyles.ContainsKey(key))
                {
                    return elementStyles[key];
                }
            }

            return "";
        }

        public static Point GetHtmlElementClientPosition(HtmlElement element)
        {
            Point p = new Point();
            p.X = element.ClientRectangle.Left + element.ClientRectangle.Width / 2;
            p.Y = element.ClientRectangle.Right + element.ClientRectangle.Height / 2;
            return p;
        }

        public static Point GetHtmlElementScreenPosition(HtmlElement element)
        {
            Point p = new Point();

            p.X = element.OffsetRectangle.Width / 2;
            p.Y = element.OffsetRectangle.Height / 2;
            UtilsLog.Log("GetHtmlElementScreenPosition p=" + p.X.ToString() + " " + p.Y.ToString());

            HtmlElement nextElement = element;
            while (nextElement.TagName.ToLower() != "body")
            {
                p.X += nextElement.OffsetRectangle.Left;
                p.Y += nextElement.OffsetRectangle.Top;
                nextElement = nextElement.Parent;
                UtilsLog.Log("GetHtmlElementScreenPosition p=" + p.X.ToString() + " " + p.Y.ToString());
            }

            p.X -= element.Document.Body.ScrollLeft;
            p.Y -= element.Document.Body.ScrollTop;

            HtmlWindow window = element.Document.Window;
            p.X += window.Position.X;
            p.Y += window.Position.Y;
            UtilsLog.Log("GetHtmlElementScreenPosition p=" + p.X.ToString() + " " + p.Y.ToString());

            return p;
        }

        public static void SaveImageElementToFile(HtmlWindow frameWindow, HtmlElement imageElement, string imageFileName, ImageFormat imageFormat)
        {
            HTMLDocument doc = (HTMLDocument)frameWindow.Document.DomDocument;
            HTMLBody body = (HTMLBody)doc.body;
            IHTMLControlRange rang = (IHTMLControlRange)body.createControlRange();
            IHTMLControlElement Img = (IHTMLControlElement)imageElement.DomElement;

            rang.add(Img);
            rang.execCommand("Copy", false, null);
            Image newImage = Clipboard.GetImage();
            newImage.Save(imageFileName, imageFormat);
            Clipboard.Clear();
        }

    }
}
