using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HatchArea
{
  public  static class CollectionHelper
    {
        public static  bool IsContain<T>(this List<string> list,string str)
        {
            return list.Any(str.Contains);
        }
    }
}
