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

using System.Windows.Forms;
using System.Resources;
using System.Collections;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportImport_MazzerTraduzioni
{
    

    class Program
    {       
        static void Main(string[] args)
        {
            //if (args.Length == 0)
            //{
            //    Console.WriteLine("Scrivere percorso file");
            //    return;
            //}

            //string file = args[0];

            // example JSON
            //string path_dir_resources = @"C:\Users\quan\Documents\project_2023\parser\parser_resources_json";
            // example RESX
            string path_dir_resources = @"C:\Users\quan\Documents\project_2023\parser\parser_resources_resx";

            string column1 = "";
            string column2 = "";
            string[] path_files = Directory.GetFiles(path_dir_resources);
            string jsonInput;
            string? languague_suffix;
            string? column_name = "";
            bool contains = false;
            List<Item> data;
            DataTable dt;
            DataRow? query_res;
            DataColumnCollection dt_columns;


            dt = new DataTable();

            foreach (string path_file in path_files)
            {                
                if (Path.GetExtension(path_file) == ".json")
                {
                    column1 = "Area";
                    column2 = "Id";

                    dt_columns = dt.Columns;
                    if (dt_columns.Contains(column1) == false && dt_columns.Contains(column2) == false)
                    {
                        dt.Columns.Add(column1, typeof(string));
                        dt.Columns.Add(column2, typeof(string));

                    }

                    // get langugae suffix from file name, e.g., "it", "en", "es", etc.
                    languague_suffix = path_file.Split('\\').Last();
                    languague_suffix = languague_suffix.Split('.')[0];
                    languague_suffix = languague_suffix.Split('_')[1];                    
                    column_name = languague_suffix;
                    
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
                        contains = dt.AsEnumerable().Any(row => item.area == row.Field<String>("Area") && item.key == row.Field<String>("Id"));
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
                                          .SingleOrDefault(row => row.Field<String>("Area") == item.area && row.Field<String>("Id") == item.key);   
                            // if find the result, then insert the value
                            if (query_res != null)
                            {
                                query_res[column_name] = item.value;
                            }
                        }
                    }
                }

                if (Path.GetExtension(path_file) == ".resx")
                {
                    column1 = "Id";
                    dt_columns = dt.Columns;
                    if (!dt_columns.Contains(column1))
                    {
                        dt.Columns.Add(column1, typeof(string));
                    }

                    using (ResXResourceReader resxReader = new ResXResourceReader(path_file))
                    {
                        if (resxReader == null)
                        {
                            continue;
                        }

                        foreach (DictionaryEntry entry in resxReader)
                        {                                                          
                            if ((string?)entry.Key == "Language")
                            {
                                languague_suffix = (string?)entry.Value;
                                column_name = languague_suffix;
                                
                                // add new language column
                                dt.Columns.Add(column_name, typeof(string));
                            }

                            contains = dt.AsEnumerable().Any(row => (string)entry.Key == row.Field<String>("Id"));
                            if (contains == false)
                            {
                                if ((string)entry.Key != "Language")
                                {
                                    DataRow newRow = dt.NewRow();
                                    newRow[column1] = (string)entry.Key;
                                    newRow[column_name] = (string?)entry.Value;
                                    dt.Rows.Add(newRow);
                                }                                
                            }
                            else
                            {
                                query_res = dt.AsEnumerable()
                                          .SingleOrDefault(row => (string)entry.Key == row.Field<String>("Id"));
                                // if find the result, then insert the value
                                if (query_res != null)
                                {
                                    query_res[column_name] = (string?)entry.Value;
                                }
                            }
                        }
                    }                    
                }
            }
            //TO DO JSON > EXCEL
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "translate_mapping");
            string xlsPath = @"C:\Users\quan\Documents\project_2023\parser\parser_output\webapp.xlsx";
            wb.SaveAs(xlsPath);          

            //string fileName = "WorksheetName_" + DateTime.Now.ToLongTimeString() + ".xlsx";
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


