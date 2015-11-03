using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Sparta.Utils
{
    public static class ExcelHelper
    {
        public static Range GetIntersection(this Range left, Range right) => left.Application.Intersect(left, right);
    }
}
