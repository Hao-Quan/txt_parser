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
using System.Data.OleDb;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace ExportImport_MazzerTraduzioni
{
    

    class Program
    {       
        static void Main(string[] args)
        {
            //ToExcel();
            ToOriginalFile();
        }

        public static void ToOriginalFile()
        {
            string path_dir_resources = @"C:\Users\quan\Documents\project_2023\parser\parser_resources_excel";
            string path_output_original_files = @"C:\Users\quan\Documents\project_2023\parser\parser_output_original_files";
            Directory.CreateDirectory(path_output_original_files);
            string[] path_files = Directory.GetFiles(path_dir_resources);
            string sSheetName = null;
            string sConnection = null;
            string first_cell = null;
            DataTable dtTablesList = default(DataTable);
            string id_string = null;
            string value_string = null;
            string[] columnNames = null;
            DataRow current_row = null;
            DataRow next_row = null;

            OleDbConnection oleExcelConnection = default(OleDbConnection);

            foreach (string path_file in path_files)
            {
                if (Path.GetExtension(path_file) != ".xlsx")
                {
                    continue;
                }
                sConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path_file + ";" + "Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\"";
                //sConnection = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" + path_file + ";" + "Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\"";
                oleExcelConnection = new OleDbConnection(sConnection);
                oleExcelConnection.Open();

                dtTablesList = oleExcelConnection.GetSchema("Tables");

                if (dtTablesList.Rows.Count > 0)
                {
                    sSheetName = dtTablesList.Rows[0]["TABLE_NAME"].ToString();
                }

                dtTablesList.Clear();
                dtTablesList.Dispose();


                if (string.IsNullOrEmpty(sSheetName) == true)
                {
                    continue;
                }

                /*Just for read excel */
                //OleDbCommand oleExcelCommand = default(OleDbCommand);
                //OleDbDataReader oleExcelReader = default(OleDbDataReader);
                //oleExcelCommand = oleExcelConnection.CreateCommand();
                //oleExcelCommand.CommandText = "Select * From [" + sSheetName + "]";
                //oleExcelCommand.CommandType = CommandType.Text;
                //oleExcelReader = oleExcelCommand.ExecuteReader();
                //int nOutputRow = 0;
                //while (oleExcelReader.Read())
                //{
                //    var row1Col0 = oleExcelReader[1];
                //    Console.WriteLine(row1Col0);
                //}
                //oleExcelReader.Close();



                //dynamic product = new JObject();
                //product.ProductName = "Elbow Grease";
                //product.Enabled = true;
                //product.Price = 4.90m;
                //product.StockCount = 9000;
                //product.StockValue = 44100;
                //product.Tags = new JArray("Real", "OnSale");

                //string id = "title";
                //product.id = 44100;

                //Console.WriteLine(product.ToString());


                // my pretend dataset
                //List<string> fields = new List<string>();
                //// my 'columns'
                //fields.Add("this_thing");
                //fields.Add("that_thing");
                //fields.Add("the_other");               

                /*Read excel and store into datatable */
                //oleExcelConnection.GetOleDbSchemaTable();

                DataSet ds = new DataSet();
                //ds.Tables[0].Columns[0].MaxLength = 
                string sqlquery = "Select * From [" + sSheetName + "]";                
                OleDbDataAdapter da = new OleDbDataAdapter(sqlquery, sConnection);
                da.Fill(ds);
                DataTable dt_excel = ds.Tables[0];

                //int s = ds.Tables[0].Columns[1].MaxLength;

                first_cell = (string)dt_excel.Rows[0][0];
                if (first_cell == "Area")
                {
                    /*Convert EXCEL to JSON*/

                    /* move first row as header columns' names*/
                    foreach (DataColumn column in dt_excel.Columns)
                    {
                        string cName = dt_excel.Rows[0][column.ColumnName].ToString();
                        if (!dt_excel.Columns.Contains(cName) && cName != "")
                        {
                            column.ColumnName = cName;
                        }

                    }
                    dt_excel.Rows[0].Delete();
                    dt_excel.Rows.RemoveAt(0);

                    columnNames = (from dc in dt_excel.Columns.Cast<DataColumn>()
                                            select dc.ColumnName).ToArray();
                    // ignore first 2 column names "area", "id", start from only language tags. 
                    columnNames = columnNames.Skip(2).ToArray();

                    dynamic exo = new System.Dynamic.ExpandoObject();
                    dynamic exo_1 = new System.Dynamic.ExpandoObject();

                    foreach (string languageColumn in columnNames)
                    {                        
                        //foreach (DataRow row in dt_excel.Rows)
                        for (int i = 0; i < dt_excel.Rows.Count - 1; i++)
                        {
                            current_row = dt_excel.Rows[i];
                            next_row = dt_excel.Rows[i + 1];

                            id_string = (string)current_row["Id"];
                            value_string = null;
                            if (current_row[languageColumn] == DBNull.Value)
                            {
                                value_string = "";
                            }
                            else
                            {
                                value_string = (string)current_row[languageColumn];
                            }
                            // when "Area" is empty
                            if (current_row["Area"] == DBNull.Value)
                            {
                                ((IDictionary<String, Object>)exo).Add(id_string, value_string);
                            }
                            // when "Area" is empty, it means it has subobjects
                            else
                            {
                                List<string> listArea = ((string)current_row["Area"]).Split('.').ToList<string>();
                                //List<string> listArea = ((string)row["Area"]).Split('.').Reverse().ToList<string>();
                                foreach (string areaElement in listArea)
                                {                                    
                                    // the next row record has the same area with the current row
                                    if ((string)next_row["Area"]==((string)current_row["Area"]))
                                    {
                                        ((IDictionary<String, Object>)exo_1).Add(id_string, value_string);
                                    }
                                    // the next row record has the different area with the current row, so we can add the subobject to the parent node
                                    else
                                    {
                                        ((IDictionary<String, Object>)exo_1).Add(id_string, value_string);
                                        ((IDictionary<String, Object>)exo).Add(areaElement, exo_1);
                                        exo_1 = new System.Dynamic.ExpandoObject();
                                    }                                    
                                }
                                

                            }
                            //((IDictionary<String, Object>)exo_1).Add("subkey", "subvalore");
                            //((IDictionary<String, Object>)exo).Add("sotto", exo_1);

                            string e_json = Newtonsoft.Json.JsonConvert.SerializeObject(exo, Formatting.Indented);
                            var sqq = 2;
                                  
                        }
                        var sss = 1;
                        Newtonsoft.Json.JsonConvert.SerializeObject(exo);
                        
                    }

                    //    dynamic exo = new System.Dynamic.ExpandoObject();
                    
                    //foreach (string field in fields)
                    //{
                    //    ((IDictionary<String, Object>)exo).Add(field, field + "_data");
                    //}
                    
                }
                else
                {
                    /*Convert EXCEL to RESX*/

                    /* move first row as header columns' names*/
                    foreach (DataColumn column in dt_excel.Columns)
                    {
                        string cName = dt_excel.Rows[0][column.ColumnName].ToString();
                        if (!dt_excel.Columns.Contains(cName) && cName != "")
                        {
                            column.ColumnName = cName;
                        }

                    }
                    dt_excel.Rows[0].Delete();
                    dt_excel.Rows.RemoveAt(0);

                    columnNames = (from dc in dt_excel.Columns.Cast<DataColumn>()
                                            select dc.ColumnName).ToArray();
                    // ignore first column name "id", start from only language tags. 
                    columnNames = columnNames.Skip(1).ToArray();

                    ResXResourceWriter resx = null;
                    
                    foreach (string languageColumn in columnNames)
                    {
                        // TODO: change output file with dynamic names: en, it, es. 
                        resx = new ResXResourceWriter(Path.Combine(path_output_original_files, "AppResources." + languageColumn + ".resx"));
                        resx.AddResource("Language", languageColumn);

                        foreach (DataRow row in dt_excel.Rows)
                        {                            
                            id_string = (string)row["Id"];
                            value_string = null;
                            if (row[languageColumn] == DBNull.Value)
                            {
                                value_string = "";
                            }
                            else
                            {
                                value_string = (string)row[languageColumn];
                            }

                            //if (id_string == "PolicyViewController_Text")
                            //{
                            //    int yy = 1;
                            //}

                            resx.AddResource(id_string, value_string);
                        }
                        // Important to close the WRITER, otherwise it will raise error
                        resx.Close();
                    }


                }                                               
                oleExcelConnection.Close();
            }
            
        }

        public static void RecursiveParseToJson(string area, string id, string lang)
        {

        }

        public static void ToExcel()
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

            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "translate_mapping");
            string xlsPath = @"C:\Users\quan\Documents\project_2023\parser\parser_output\webapp.xlsx";
            wb.SaveAs(xlsPath);

            string fileName = "WorksheetName_" + DateTime.Now.ToLongTimeString() + ".xlsx";
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

        public class Item_Resx
        {
            public string key;
            public string value;
        }

    }
}


