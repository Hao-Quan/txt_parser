//using Aspose.Cells;
//using Aspose.Cells.Utility;
using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportImport_MazzerTraduzioni
{
    class Program
    {
        static string column1 = "Area";
        static string column2 = "Id";

        static void Main(string[] args)
        {
            //if (args.Length == 0)
            //{
            //    Console.WriteLine("Scrivere percorso file");
            //    return;
            //}

            //string file = args[0];
            string path_dir_resources = @"C:\Users\quan\Documents\project_2023\parser\parser_resources";
            string[] path_files = Directory.GetFiles(path_dir_resources);
            string jsonInput;
            string languague_suffix;
            string column_name = "";
            List<Item> data;
            DataTable dt;
            DataRow? query_res;
            
            dt = new DataTable();
            dt.Columns.Add(column1, typeof(string));
            dt.Columns.Add(column2, typeof(string));

            foreach (string path_file in path_files)
            {                
                if (Path.GetExtension(path_file) == ".json")
                {
                    // get langugae suffix from file name, e.g., "it", "en", "es", etc.
                    languague_suffix = path_file.Split('\\').Last();
                    languague_suffix = languague_suffix.Split('.')[0];
                    languague_suffix = languague_suffix.Split('_')[1];

                    switch (languague_suffix)
                    {
                        case "it":
                            column_name = "Italiano/IT(it)";
                            break;
                        case "en":
                            column_name = "English/EN(en)";
                            break;
                    }
                    // add new language column
                    dt.Columns.Add(column_name, typeof(string));

                    // start to parse JSON file
                    jsonInput = File.ReadAllText(path_file);
                    data = JsonParser(jsonInput);

                    if (data.Count == 0)
                        continue;

                    foreach (Item item in data)
                    {   
                        // check whether the specific pair <area, id> already exists in the data table
                        bool contains = dt.AsEnumerable().Any(row => item.area == row.Field<String>("Area") && item.key == row.Field<String>("Id"));
                        // if the pair <area, id> does not exist in the data table, insert it values of <area, key, value> as a new row
                        if (contains == false)
                        {
                            DataRow newRow = dt.NewRow();
                            newRow[column1] = item.area;
                            newRow[column2] = item.key;
                            newRow[column_name] = item.value;
                            dt.Rows.Add(newRow);
                        }
                        // if the pair <area, id> exists in the data table, search the row of the pair <area, id>, and add the new language translated record to that row. 
                        else
                        {                            
                            query_res = dt.AsEnumerable()
                                          .SingleOrDefault(r => r.Field<String>("Area") == item.area && r.Field<String>("Id") == item.key);   
                            // if find the result, then insert the value
                            if (query_res != null)
                            {
                                query_res[column_name] = item.value;
                            }
                        }
                    }
                }                
            }
                       
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "translate_mapping");
            string xlsPath = @"C:\Users\quan\Documents\project_2023\parser\parser_output\webapp.xlsx";
            wb.SaveAs(xlsPath);          

            //string fileName = "WorksheetName_" + DateTime.Now.ToLongTimeString() + ".xlsx";
            //string xlsPath = @"C:\Users\fonio.TXTGROUP\Desktop\Mazzer\file traduzioni\" + fileName;
            //string xlsPath = @"C:\Users\fonio.TXTGROUP\Desktop\Mazzer\file traduzioni\webapp_it.xlsx";
        }

        static public List<Item> JsonParser(string json)
        {
            try
            {
                List<Item> records = new List<Item>();
                JObject my_obj = JsonConvert.DeserializeObject<JObject>(json);
                foreach (KeyValuePair<string, JToken> sub_obj in my_obj)
                {
                    string key = sub_obj.Key;
                    var value = sub_obj.Value.ToString();
                    JToken token = sub_obj.Value;

                    // case 'value' includes sub-subject 
                    if (value.Contains("\r\n"))
                    {
                        RecursiveParse(records, token, string.Empty);
                    }
                    // simple <key: value> pairs
                    else
                    {
                        Item record = new Item();
                        record.area = string.Empty;
                        record.key = sub_obj.Key;
                        record.value = sub_obj.Value.ToString();
                        records.Add(record);
                    }
                }

                return records;

            }
            catch (Exception e)
            {
                string mesg = e.Message;
                return new List<Item>();
            }
        }

        public static void RecursiveParse(List<Item> records, JToken token, string area)
        {
            foreach (JToken innerItem in token.Values())
            {
                if (innerItem.Type == JTokenType.Object)
                {
                    RecursiveParse(records, innerItem, string.Empty);
                }
                else
                {
                    string str = innerItem.Parent.ToString();
                    string innerkey = str.Substring(0, str.LastIndexOf(':'));
                    innerkey = innerkey.Replace(" ", string.Empty);
                    innerkey = innerkey.Replace(@"""", string.Empty);
                    string innerValue = str.Substring(str.LastIndexOf(':') + 2);
                    innerValue = innerValue.Replace(@"""", string.Empty);

                    Item record = new Item();
                    if (string.IsNullOrEmpty(area))
                        record.area = token.Path.ToString();
                    else
                        record.area = area + "-" + token.Path.ToString();
                    record.key = innerkey;
                    record.value = innerValue;
                    records.Add(record);
                }
            }
        }

        public class Item
        {
            public string area;
            public string key;
            public string value;
        }

    }
}


