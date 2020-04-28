using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Configuration;

namespace nndk
{
    public static class Common
    {
        public static string GetCodes()
        {
            return ConfigurationSettings.AppSettings["codes"].ToString();
        }
    }
}
