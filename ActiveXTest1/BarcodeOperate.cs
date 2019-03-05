using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ActiveXTest1
{
    class BarcodeOperate
    {
        private readonly Dictionary<char, string> _code39Dictionary;

        public BarcodeOperate()
        {
            _code39Dictionary = new Dictionary<char, string>();
            _code39Dictionary['0'] = "101001101101";
            _code39Dictionary['1'] = "110100101011";
            _code39Dictionary['2'] = "101100101011";
            _code39Dictionary['3'] = "110110010101";
            _code39Dictionary['4'] = "101001101011";
            _code39Dictionary['5'] = "110100110101";
            _code39Dictionary['6'] = "101100110101";
            _code39Dictionary['7'] = "101001011011";
            _code39Dictionary['8'] = "110100101101";
            _code39Dictionary['9'] = "101100101101";
            _code39Dictionary['A'] = "110101001011";
            _code39Dictionary['B'] = "101101001011";
            _code39Dictionary['C'] = "110110100101";
            _code39Dictionary['D'] = "101011001011";
            _code39Dictionary['E'] = "110101100101";
            _code39Dictionary['F'] = "101101100101";
            _code39Dictionary['G'] = "101010011011";
            _code39Dictionary['H'] = "110101001101";
            _code39Dictionary['I'] = "101101001101";
            _code39Dictionary['J'] = "101011001101";
            _code39Dictionary['K'] = "110101010011";
            _code39Dictionary['L'] = "101101010011";
            _code39Dictionary['M'] = "110110101001";
            _code39Dictionary['N'] = "101011010011";
            _code39Dictionary['O'] = "110101101001";
            _code39Dictionary['P'] = "101101101001";
            _code39Dictionary['Q'] = "101010110011";
            _code39Dictionary['R'] = "110101011001";
            _code39Dictionary['S'] = "101101011001";
            _code39Dictionary['T'] = "101011011001";
            _code39Dictionary['U'] = "110010101011";
            _code39Dictionary['V'] = "100110101011";
            _code39Dictionary['W'] = "110011010101";
            _code39Dictionary['X'] = "100101101011";
            _code39Dictionary['Y'] = "110010110101";
            _code39Dictionary['Z'] = "100110110101";
            _code39Dictionary['-'] = "100101011011";
            _code39Dictionary['.'] = "110010101101";
            _code39Dictionary[' '] = "100110101101";
            _code39Dictionary['$'] = "100100100101";
            _code39Dictionary['/'] = "100100101001";
            _code39Dictionary['+'] = "100101001001";
            _code39Dictionary['%'] = "101001001001";
            _code39Dictionary['*'] = "100101101101";
        }

        public Bitmap Create(int width, int height, int paddingLeft, string content)
        {

            var encodedBits = new StringBuilder();
            encodedBits.Append("1001011011010"); // Code39 prefix + 0 separator at the end
            foreach (char c in content.ToUpper())
            {
                string encoded;
                if (!_code39Dictionary.TryGetValue(c, out encoded))
                    throw new ArgumentOutOfRangeException("content", "Characher '" + c + "' is not compatible with this Code39 implementation!");
                encodedBits.Append(encoded + '0');
            }
            encodedBits.Append("100101101101");  // Code39 suffix

            // We initially set the X coordinate to match the padding so that
            // drawing begins from the padding
            int offsetLeft = paddingLeft;


            //Set Width
            width = encodedBits.Length + (offsetLeft * 2);

            // Create bitmap canvas object with specified size
            var canvas = new Bitmap(width, height);
            // Create graphics to draw on
            var graphics = Graphics.FromImage(canvas);
            // Fill with background color
            
            graphics.FillRectangle(Brushes.White, new Rectangle(0, 0, width, height));


            // Create Label Text Layer

            var TextLayerSize = graphics.MeasureString(content, new Font("Tahoma", 8));
            int TextLayerSizeHeight = (int)Math.Ceiling(TextLayerSize.Height);
            int TextLayerSizeWidth = (int)Math.Ceiling(TextLayerSize.Width);
            graphics.DrawString(content, new Font("Tahoma", 8), Brushes.Black, new RectangleF((width / 2) - (TextLayerSizeWidth / 2), height - TextLayerSizeHeight, TextLayerSizeWidth, TextLayerSizeHeight));
            graphics.DrawString("张三", new Font("Tahoma", 8), Brushes.Black, new RectangleF(0, 0, TextLayerSizeWidth, TextLayerSizeHeight));
            // End .


            foreach (char c in encodedBits.ToString())
            {
                var rectangle = new Rectangle(offsetLeft++, 0, 1, height - TextLayerSizeHeight);
                graphics.FillRectangle(c == '0' ? Brushes.White : Brushes.Black, rectangle);
            }
            
            return canvas;
        }


        public Bitmap CreateCode128()
        {
            Bitmap b = new Bitmap(100,100);
            return b;
        }
    }
}

