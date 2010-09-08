using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

using System.Globalization;

namespace Resx2Xls
{
    public class CultureInfoComparer : IComparer
    {
        // Methods
        public int Compare(object x, object y)
        {
            if (((x == null) && (y == null)) || x.Equals(y))
            {
                return 0;
            }
            if (x.Equals(CultureInfo.InvariantCulture) || (y == null))
            {
                return -1;
            }
            if (y.Equals(CultureInfo.InvariantCulture) || (x == null))
            {
                return 1;
            }
            if (!(x is CultureInfo))
            {
                throw new ArgumentException("Can only compare CultureInfo objects.", "x");
            }
            string displayName = ((CultureInfo)x).DisplayName;
            if (!(y is CultureInfo))
            {
                throw new ArgumentException("Can only compare CultureInfo objects.", "y");
            }
            string strB = ((CultureInfo)y).DisplayName;
            return displayName.CompareTo(strB);
        }
    }


}
