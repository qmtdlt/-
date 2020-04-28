using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace nndk
{
    public static class DateTimeExtension
    {
        /// <summary>
        /// 可空日期转短日期字符串
        /// </summary>
        /// <returns></returns>
        public static string ToShortTimeString(this DateTime? date)
        {
            if (date == null)
            {
                return "";
            }
            else
            {
                return date.Value.ToShortDateString();
            }
        }
    }
    public static class StringExtension
    {
        /// <summary>
        /// 字符串转时间
        /// </summary>
        /// <param name="src"></param>
        /// <returns></returns>
        public static DateTime? ToDateTime(this string src)
        {
            if (DateTime.TryParse(src, out DateTime date))
            {
                return date;
            }
            else
            {
                return null;
            }
        }
    }
    public static class ObjectExtension
    {
        /// <summary>
        /// obj转可空int类型，obj为空，返回空，obj转int成功，正常返回，否则返回null
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static int? objToInt(this object obj)
        {
            if (obj == null)
            {
                return null;
            }
            else
            {
                if (int.TryParse(obj.ToString(), out int res))
                {
                    return res;
                }
                else
                {
                    return null;
                }
            }
        }
        /// <summary>
        /// obj转可空double类型，obj为空，返回空，obj转double成功，正常返回，否则返回null
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static double? objToDouble(this object obj)
        {
            if (obj == null)
            {
                return null;
            }
            else
            {
                if (double.TryParse(obj.ToString(), out double res))
                {
                    return res;
                }
                else
                {
                    return null;
                }
            }
        }
        public static decimal? objToDecimal(this object obj)
        {
            if (obj == null)
            {
                return null;
            }
            else
            {
                if (decimal.TryParse(obj.ToString(), out decimal res))
                {
                    return res;
                }
                else
                {
                    return null;
                }
            }
        }

        public static float? objToFloat(this object obj)
        {
            if (obj == null)
            {
                return null;
            }
            else
            {
                if (float.TryParse(obj.ToString(), out float res))
                {
                    return res;
                }
                else
                {
                    return null;
                }
            }
        }
    }
}
