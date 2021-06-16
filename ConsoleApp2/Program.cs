using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Reflection;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            var transformation = new ShipmentTrasnformation(666, "Jeffrey");
            var transformer = new Transformer();

            var result = transformer.ReplaceColumnNames(transformation).AddColumn("new column here", 1000).AddColumn("another column","aaaaa").RemoveColumn("LastName");

            Console.WriteLine(JsonConvert.SerializeObject(result));
        }
    }

    public class Transformer
    {
        public Dictionary<string, dynamic> newColumnNameValueDic { get; } = new Dictionary<string, dynamic>();

        public Transformer()
        {

        }
        public Transformer ReplaceColumnNames(dynamic obj) 
        {
            var properties = typeof(ShipmentTrasnformation).GetProperties();
            string columnName;
            foreach (var property in properties)
            {
                Console.WriteLine($"original property name: {property.Name}");

                if (Attribute.IsDefined(property, typeof(ExcelColumnNameAttribute)))
                {
                    columnName = property.GetCustomAttribute<ExcelColumnNameAttribute>().ExcelColumn;
                }
                else
                {
                    columnName = property.Name;
                }

                var currentValue = obj.GetType().GetProperty(property.Name).GetValue(obj);

                Console.WriteLine($"excel column name: {columnName}");
                Console.WriteLine($"excel column value: {currentValue}");


                newColumnNameValueDic.Add(columnName, currentValue);
                Console.WriteLine("----------------------------------------------------");
            }

            return this;
        }

        public Transformer AddColumn(string columnName, dynamic value)
        {
            newColumnNameValueDic.Add(columnName, value);

            return this;
        }

        public Transformer RemoveColumn(string columnName)
        {
            newColumnNameValueDic.Remove(columnName);

            return this;
        }
    }


    public class ShipmentTrasnformation
    {
        [ExcelColumnName("MappedColumn")]
        public int DataValue { get; set; }

        public string Author { get; set; }

        [ExcelColumnName("FirstName")]
        public string First { get; set; }

        [ExcelColumnName("LastName")]
        public string Last { get; set; }

        public ShipmentTrasnformation(int val, string aut)
        {
            DataValue = val;
            Author = aut;
            First = transformFrist();
            Last = "BBBBB";
        }

        public string transformFrist()
        {
            // do whatever...
            return "AAAA";
        }
    }

    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnNameAttribute : Attribute 
    {
        public string ExcelColumn { get; set; }

        public ExcelColumnNameAttribute(string excelColumn)
        {
            ExcelColumn = excelColumn;
        }
    }
}
