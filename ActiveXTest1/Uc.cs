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

namespace ActiveXTest1
{
    [Guid("073A987E-2A7C-4874-8BEE-321E04F4E84E")]
    public partial class Uc : UserControl, IObjectSafety
    {
        /// <summary>
        /// 打印文档的对象
        /// </summary>
        private PrintDocument m_printDoc = null;
        private float m_pageWidth = 50F;//纸张宽度 mm单位
        private float m_pageHeight = 30F;//纸张高度 mm单位
        private BarCodeInfo SaveData = null;//打印的数据
        public Uc()
        {
            XDocument document = XDocument.Load(@"D:\nykjprint\PrintSet.xml");
            XElement root = document.Root;
            InitializeComponent();
            
            m_printDoc = new PrintDocument();//实例打印文档对象
            //自定义纸张大小
            m_printDoc.DefaultPageSettings.PaperSize = new PaperSize("newPage70X40"
           , (int)(m_pageWidth / 25.4 * 100)
           , (int)(m_pageHeight / 25.4 * 100));
            MessageBox.Show(root.Value);
            m_printDoc.PrinterSettings.PrinterName = root.Value;
            //自定义图片内容整体上间距/左间距
            m_printDoc.OriginAtMargins = true;
            m_printDoc.DefaultPageSettings.Margins.Top = (int)(2 / 25.4 * 100);
            m_printDoc.DefaultPageSettings.Margins.Left = (int)(2 / 25.4 * 100);
            //打印事件
            m_printDoc.PrintPage += new PrintPageEventHandler(m_printDoc_PrintPage);
            m_printDoc.EndPrint += new PrintEventHandler(m_printDoc_EndPrintPage);

           
    }

        private void m_printDoc_EndPrintPage(object sender, PrintEventArgs e)
        {
             SaveData = null;           
        }



        /// <summary>
        /// 绘制需要打印的内容
        /// </summary>
        private void m_printDoc_PrintPage(object sender, PrintPageEventArgs e)
        {
            var MyBarcode = new BarcodeOperate();
            if (SaveData == null) return;
            BarCodeInfo barcode = SaveData;
            //条码第一行文本信息
            StringBuilder sb = new StringBuilder();
            sb.Append(barcode.PATNAME + "  " + barcode.PATSEX + "  " + barcode.PATAGE + "  " + barcode.SPLITTYPENAME);
            //创建第一行文本信息
            e.Graphics.DrawString(sb.ToString(), new Font("宋体", 10), Brushes.Black, 2, 2);
            //绘制条码图案
            C_Code128 _Code = new C_Code128();
            _Code.ValueFont = new Font("宋体", 10);
            e.Graphics.DrawImage(_Code.GetCodeImage(barcode.Barcode, C_Code128.Encode.Code128C), new Point() { X = 38, Y = 18 });

            //条码左侧和右侧的文字信息
            string barcodeimage_lefttext = "";
         
            if (!string.IsNullOrWhiteSpace(barcode.ISREPRINT)&& barcode.ISREPRINT=="1")
            {                
                barcodeimage_lefttext=  DateTime.Now.Month + "/" + DateTime.Now.Day + "\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute+ "\r\n" +" 重";
            }
            else
            {
                barcodeimage_lefttext= DateTime.Now.Year + "\r\n" + DateTime.Now.Month + "/" + DateTime.Now.Day + "\r\n" + DateTime.Now.Hour + ":" + DateTime.Now.Minute;
            }
            string barcodeimage_righttext = barcode.SPECIMENTYPECODE + "\r\n" + barcode.DEPTNAME;
            e.Graphics.DrawString(barcodeimage_lefttext, new Font("宋体", 8), Brushes.Black, 0, 18);
            e.Graphics.DrawString(barcodeimage_righttext, new Font("宋体", 8), Brushes.Black, 150, 18);


            //添加病人号

            e.Graphics.DrawString(barcode.PATCODE, new Font("宋体", 10, FontStyle.Bold), Brushes.Black, 48, 58);

            //创建项目文本信息
            char[] sarray = barcode.Items.ToArray();
            string itemdata = "";
            string itemdataprint = "";
            for (int i = 0; i < barcode.Items.Count(); i++)
            {
                if ((itemdataprint.Length + 1) % 15 == 0 && i > 0)
                {
                    itemdataprint += sarray[i] + "\r\n";
                    itemdata += sarray[i];
                }
                else
                {
                    itemdata += sarray[i];
                    itemdataprint += sarray[i];
                }

            }
            e.Graphics.DrawString(itemdataprint, new Font("宋体", 8), Brushes.Black, 0, 78);

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
            SaveData = JsonConvert.DeserializeObject<BarCodeInfo>(barcodedata);
            m_printDoc.Print();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
       
    }
}