using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing.Printing;
using System.IO;
using System.Drawing;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Data;
using System.Linq;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;

namespace ActiveXTest1
{
    [Guid("073A987E-2A7C-4874-8BEE-321E04F4E84E")]
    public partial class Uc : UserControl, IObjectSafety
    {
        /// <summary>
        /// 打印文档的对象
        /// </summary>
        private PrintDocument m_printDoc = null;
        private JObject jsondata = null;
        private JObject jsonset = null;
        public Uc()
        {
            

            InitializeComponent();
            

            //打印事件
            m_printDoc.PrintPage += new PrintPageEventHandler(m_printDoc_PrintPage);
            m_printDoc.EndPrint += new PrintEventHandler(m_printDoc_EndPrintPage);
            m_printDoc.BeginPrint += M_printDoc_BeginPrint;
    }



        private void M_printDoc_BeginPrint(object sender, PrintEventArgs e)
        {
            string jsonfile = System.Environment.CurrentDirectory + "//BarCodeModule.json";
            using (System.IO.StreamReader file = System.IO.File.OpenText(jsonfile))
            {

                using (JsonTextReader reader = new JsonTextReader(file))
                {

                    jsonset = (JObject)JToken.ReadFrom(reader);

                }
            }

            m_printDoc = new PrintDocument();//实例打印文档对象
            //自定义纸张大小
            m_printDoc.DefaultPageSettings.PaperSize = new PaperSize(jsonset["pagename"].ToString()
           , (int)(jsonset["width"].Toint() / 25.4 * 100)
           , (int)(jsonset["height"].Toint() / 25.4 * 100));
            m_printDoc.PrinterSettings.PrinterName = jsonset["printname"].ToString();
            //自定义图片内容整体上间距/左间距
            m_printDoc.OriginAtMargins = true;
            m_printDoc.DefaultPageSettings.Margins.Top = 0;
            m_printDoc.DefaultPageSettings.Margins.Left = 0;
        }

        private void m_printDoc_EndPrintPage(object sender, PrintEventArgs e)
        {
            jsondata = null;           
        }


        /// <summary>
        /// 绘制需要打印的内容
        /// </summary>
        private void m_printDoc_PrintPage(object sender, PrintPageEventArgs e)
        {


            JObject jobject = jsonset;
            foreach (var item in jsondata)
            {
                try
                {
                    JToken jo2 = jobject["parameter"][item.Key];
                    if (jo2 != null)
                    {
                        e.Graphics.DrawString(item.Value.ToString(), new Font(jo2["fontname"].ToString(), Convert.ToInt16(jo2["fontsize"]), (FontStyle)Convert.ToUInt16(jo2["fontstyle"])), Brushes.Black, Convert.ToUInt16(jo2["left"]), Convert.ToUInt16(jo2["top"]));
                    }
                }
                catch (Exception ex)
                {

                    continue;
                }

            }
            JToken jsonimg = jobject["barcodeimg"];
            C_Code128 _Code = new C_Code128();
            _Code.ValueFont = new Font(jsonimg["fontname"].ToString(), Convert.ToInt16(jsonimg["fontsize"]));
            e.Graphics.DrawImage(_Code.GetCodeImage(jsondata["Barcode"].ToString(), C_Code128.Encode.Code128C), new Point() { X = jsonimg["left"].Toint(), Y = jsonimg["top"].Toint() });
        }

     
        #region IObjectSafety 成员
        private const string _IID_IDispatch = "{00020400-0000-0000-C000-000000000046}";
        private const string _IID_IDispatchEx = "{a6ef9860-c720-11d0-9337-00a0c90dcaa9}";
        private const string _IID_IPersistStorage = "{0000010A-0000-0000-C000-000000000046}";
        private const string _IID_IPersistStream = "{00000109-0000-0000-C000-000000000046}";
        private const string _IID_IPersistPropertyBag = "{37D84F60-42CB-11CE-8135-00AA004BB851}";

        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x00000002;
        private const int S_OK = 0;
        private const int E_FAIL = unchecked((int)0x80004005);
        private const int E_NOINTERFACE = unchecked((int)0x80004002);

        private bool _fSafeForScripting = true;
        private bool _fSafeForInitializing = true;

        public int GetInterfaceSafetyOptions(ref Guid riid, ref int pdwSupportedOptions, ref int pdwEnabledOptions)
        {
            int Rslt = E_FAIL;

            string strGUID = riid.ToString("B");
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForScripting == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForInitializing == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_DATA;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }

        public int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions)
        {
            int Rslt = E_FAIL;
            string strGUID = riid.ToString("B");
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_CALLER) && (_fSafeForScripting == true))
                        Rslt = S_OK;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_DATA) && (_fSafeForInitializing == true))
                        Rslt = S_OK;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }

        #endregion


       
        public void GetPrintData(string barcodedata)
        {
            jsondata= JObject.Parse(barcodedata);            
            m_printDoc.Print();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
       
    }
}