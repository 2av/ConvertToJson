using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace CRM.Models.Global
{
    public class GlobalFunctions
    {
        public static List<T> ConverDataTableToList<T>(DataTable dt)
        {
            var columnNames = dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName.ToLower()).ToList();
            var properties = typeof(T).GetProperties();
            return dt.AsEnumerable().Select(row =>
            {
                var objT = Activator.CreateInstance<T>();
                foreach (var pro in properties)
                {
                    if (columnNames.Contains(pro.Name.ToLower()))
                    {
                        try
                        {
                            pro.SetValue(objT, row[pro.Name]);
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
                return objT;
            }).ToList();
        }
    }

    
    
}