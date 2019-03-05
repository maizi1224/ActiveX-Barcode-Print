using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiveXTest1
{
    public class BarCodeInfo
    {
        /// <summary>
        /// 病人姓名
        /// </summary>
        public string PATNAME { get; set; }

        /// <summary>
        /// 性别
        /// </summary>
        public string PATSEX { get; set; }

        /// <summary>
        /// 年龄
        /// </summary>
        public string PATAGE { get; set; }

        /// <summary>
        /// 试管类型
        /// </summary>
        public string SPLITTYPENAME { get; set; }

        /// <summary>
        /// 打印日期
        /// </summary>
        public DateTime PRINTDATE { get; set; }

        /// <summary>
        /// 默认标本类型
        /// </summary>
        public string SPECIMENTYPECODE { get; set; }

        /// <summary>
        /// 申请科室
        /// </summary>
        public string DEPTNAME { get; set; }

        /// <summary>
        /// 申请项目
        /// </summary>
        public string Items { get; set; }

        /// <summary>
        /// 打印条码
        /// </summary>
        public string Barcode { get; set; }


        /// <summary>
        /// 病人号
        /// </summary>
        public string PATCODE { get; set; }

        /// <summary>
        /// 标志是否是重打 重打为1非重打为0
        /// </summary>
        public string ISREPRINT { get; set; }

    }
}
