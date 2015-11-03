using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Sparta.Engine.Utils;

namespace Sparta.Controls
{
    public class ValidationProperty : IRangeProperty
    {
        ReadOnlyCollection<string> _list = ReadOnlyCollectionEx<string>.Empty;

        public ReadOnlyCollection<string> List
        {
            get { return _list; }
            set
            {
                if (value == null)
                    throw new ArgumentNullException();

                _list = value;
            }
        }

        public void Paint(Range range)
        {
            var listSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;

            var valuesWithSeparator = _list.Where(v => v.Contains(listSeparator)).ToList();

            if (valuesWithSeparator.Any())
                throw new ArgumentException("Validation values should not contain '{0}', got: {1}".FormatWith(listSeparator, valuesWithSeparator.Select(v => "'" + v + "'").JoinStrings(", ")));

            var validation = range.Validation;

            validation.Delete();

            if (List.Any())
                validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertInformation, XlFormatConditionOperator.xlBetween, List.JoinStrings(listSeparator));
        }
    }
}