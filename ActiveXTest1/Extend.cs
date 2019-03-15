using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ActiveXTest1
{
   public static class Extend
    {
        public static int Toint(this JToken _JToken)
        {
            return Convert.ToInt16(_JToken);
        }
    }
}
