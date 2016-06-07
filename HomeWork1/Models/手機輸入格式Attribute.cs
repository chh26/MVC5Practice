using System;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace HomeWork1.Models
{
    internal class 手機輸入格式Attribute : RegularExpressionAttribute//DataTypeAttribute
    {
        public 手機輸入格式Attribute() : base(@"\d{4}-\d{6}$")//base(DataType.PhoneNumber)
        {

        }

        //public override bool IsValid(object value)
        //{
        //    if (value != null)
        //        return value.ToString().Length == 11 ? Regex.IsMatch(value.ToString(), @"\d{4}-\d{6}") : false;
        //    else
        //        return true;
        //}
    }
}