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
using ExcelDataReader;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;

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

            //FileStream stream = File.Open(strFileName, FileMode.Open, FileAccess.Read);
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //DataSet result = excelReader.AsDataSet();
            //excelReader.Close();
            //return result.Tables[0];

            foreach (string path_file in path_files)
            {
                if (Path.GetExtension(path_file) != ".xlsx")
                {
                    continue;
                }

                /*START ACE.OLEDB read EXCEL (have 255 characters limitation)*/
                /*
                //sConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path_file + ";" + "Extended Properties=\"Excel 12.0;HDR=No;IMEX=1\"";
                sConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path_file + ";" + "Extended Properties=\"Excel 12.0;HDR=No;IMEX=1;ImportMixedTypes=Text;TypeGuessRows=0;\"";

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

                DataSet ds = new DataSet();
                string sqlquery = "Select * From [" + sSheetName + "]";
                //string sqlquery = "Select * From [" + sSheetName + "A1:C340]";
                OleDbDataAdapter da = new OleDbDataAdapter(sqlquery, sConnection);
                da.Fill(ds);
                DataTable dt_excel = ds.Tables[0];
                */
                /*END ACE.OLEDB read EXCEL */

                /* Excel data reader FOR avoiding 255 caratteri limitation */
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                FileStream stream = File.Open(path_file, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                DataSet ds = excelReader.AsDataSet();                
                DataTable dt_excel = ds.Tables[0];                             

                first_cell = (string)dt_excel.Rows[0][0];
                if (first_cell == "Area")
                {
                    /*Convert EXCEL to JSON*/

                    /* move first row as header columns' names*/
                    foreach (DataColumn excel_column in dt_excel.Columns)
                    {
                        string cName = dt_excel.Rows[0][excel_column.ColumnName].ToString();
                        if (!dt_excel.Columns.Contains(cName) && cName != "")
                        {
                            excel_column.ColumnName = cName;
                        }

                    }
                    dt_excel.Rows[0].Delete();
                    dt_excel.Rows.RemoveAt(0);

                    /*create a copy of the ordered datatable */
                    DataTable dt_ordered = dt_excel;
                    dt_ordered.DefaultView.Sort = "Area";
                    dt_ordered = dt_ordered.DefaultView.ToTable();

                    columnNames = (from dc in dt_excel.Columns.Cast<DataColumn>()
                                            select dc.ColumnName).ToArray();
                    // ignore first 2 column names "area", "id", start from only language tags. 
                    columnNames = columnNames.Skip(2).ToArray();


                    // NEW VERSION OF PROCESSING EXCEL -> JSON
                    // get "area" has > 1 sub-subjects
                    List<string> areaWithMultipleSubObj = new List<string>();
                    List<string> firstAreaWithMultipleSubObj = new List<string>();
                    DataTable dt_support = new DataTable();
                    dt_support.Clear();
                    dt_support.Columns.Add("area");
                    dt_support.Columns.Add("start_idx", typeof(Int32));
                    dt_support.Columns.Add("end_idx", typeof(Int32));
                    DataRow dr_support = dt_support.NewRow();
                    bool flagSameFirstArea = false;
                    string currentFirstArea = null;

                    DataTable dt_lookuptab_label = new DataTable();                    
                    DataColumn column = new DataColumn("label_id");
                    column.DataType = System.Type.GetType("System.Int32");
                    column.AutoIncrement = true;
                    column.AutoIncrementSeed = 0;
                    column.AutoIncrementStep = 1;
                    dt_lookuptab_label.Columns.Add(column);
                    dt_lookuptab_label.Columns.Add("level", typeof(Int32));                    
                    dt_lookuptab_label.Columns.Add("label");
                    dt_lookuptab_label.Columns.Add("label_complete");
                    dt_lookuptab_label.Columns.Add("num_subpair", typeof(Int32));
                    dt_lookuptab_label.Columns.Add("num_subobject", typeof(Int32));


                    DataTable dt_lookuptab_pair = new DataTable();
                    column = new DataColumn("pair_id");
                    column.DataType = System.Type.GetType("System.Int32");
                    column.AutoIncrement = true;
                    column.AutoIncrementSeed = 0;
                    column.AutoIncrementStep = 1;
                    dt_lookuptab_pair.Columns.Add(column);
                    dt_lookuptab_pair.Columns.Add("label_id", typeof(Int32));
                    dt_lookuptab_pair.Columns.Add("key");
                    dt_lookuptab_pair.Columns.Add("value");

                    /* Fill */
                    for (int i = 0; i < dt_excel.Rows.Count - 1; i++)
                    {
                        if (dt_excel.Rows[i]["Area"] != DBNull.Value)
                        {
                            if (areaWithMultipleSubObj.Contains((string)dt_excel.Rows[i]["Area"]) == false
                                && ((string)dt_excel.Rows[i]["Area"]).Contains("."))
                            {                                
                                areaWithMultipleSubObj.Add((string)dt_excel.Rows[i]["Area"]);
                                    
                                if (firstAreaWithMultipleSubObj.Contains(((string)dt_excel.Rows[i]["Area"]).Split(".")[0]) == false)
                                {
                                    firstAreaWithMultipleSubObj.Add(((string)dt_excel.Rows[i]["Area"]).Split(".")[0]);

                                    
                                    dr_support["area"] = ((string)dt_excel.Rows[i]["Area"]).Split(".")[0];
                                    dr_support["start_idx"] = i;

                                    currentFirstArea = ((string)dt_excel.Rows[i]["Area"]).Split(".")[0];
                                    flagSameFirstArea = true;
                                }                                
                            }

                            if ((((string)dt_excel.Rows[i]["Area"]).Split(".")[0]) != currentFirstArea && flagSameFirstArea == true)
                            {
                                dr_support["end_idx"] = i - 1;
                                dt_support.Rows.Add(dr_support);
                                dr_support = dt_support.NewRow();
                                flagSameFirstArea = false;
                            }
                        }                                
                    }                    

                    dynamic exo = new System.Dynamic.ExpandoObject();
                    dynamic exo_1 = new System.Dynamic.ExpandoObject();

                    foreach (string languageColumn in columnNames)
                    {                        
                        //foreach (DataRow row in dt_excel.Rows)
                        for (int i = 0; i < dt_excel.Rows.Count - 1; i++)
                        {
                            current_row = dt_excel.Rows[i];
                            next_row = dt_excel.Rows[i + 1];

                            /* get current <id, value> pair*/
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

                            // when "Area" is empty, not subobjects
                            if (current_row["Area"] == DBNull.Value)
                            {
                                ((IDictionary<String, Object>)exo).Add(id_string, value_string);
                            }

                            /* when "Area" is NOT empty, it means it has subobjects*/

                            /* just manage 1 level of subobject case*/
                            else if (((string)current_row["Area"]).Split('.').Length == 1)
                            {
                                exo_1 = new System.Dynamic.ExpandoObject();
                                List<string> listArea = ((string)current_row["Area"]).Split('.').ToList<string>();
                                //List<string> listArea = ((string)row["Area"]).Split('.').Reverse().ToList<string>();
                                foreach (string areaElement in listArea)
                                {

                                    // the next row record has the same area with the current row, add pairs to the same subobject "exo_1"
                                    if ((string)next_row["Area"] == ((string)current_row["Area"]))
                                    {
                                        ((IDictionary<String, Object>)exo_1).Add(id_string, value_string);
                                    }
                                    // the next row record has the different area with the current row, so we can add the subobject "exo_1" to the parent node "exo"
                                    else
                                    {
                                        ((IDictionary<String, Object>)exo_1).Add(id_string, value_string);
                                        ((IDictionary<String, Object>)exo).Add(areaElement, exo_1);
                                        // clean the first subobject contents
                                        exo_1 = new System.Dynamic.ExpandoObject();
                                    }
                                }

                            }
                            /* try to manage multiple subobject levels */
                            else if (((string)current_row["Area"]).Split('.').Length > 1)
                            {

                                //DataRow newRow = dt.NewRow();
                                //newRow[column1] = item.area;
                                //newRow[column2] = item.key;
                                //newRow[column_name] = item.value;
                                //dt.Rows.Add(newRow);

                                dt_lookuptab_label.Clear();
                                dt_lookuptab_pair.Clear();

                                DataRow row_tb_lable = dt_lookuptab_label.NewRow();
                                row_tb_lable["level"] = 0;
                                row_tb_lable["label"] = ((string)current_row["Area"]).Split('.')[0];
                                /*only set for level 0 */
                                row_tb_lable["label_complete"] = ((string)current_row["Area"]).Split('.')[0];
                                dt_lookuptab_label.Rows.Add(row_tb_lable);

                                DataRow? row_tb_pair = null;
                                DataRow? corrispond_row = null;

                                DataRow dr_currentSubObj = dt_support.AsEnumerable().Where(r => r.Field<string>("area") == (((string)current_row["Area"]).Split('.')[0])).First();
                                int currentSubObj_startIdx = (int)dr_currentSubObj["start_idx"];
                                int currentSubObj_endIdx = (int)dr_currentSubObj["end_idx"];
                                List<string> splittedArea = new List<string>();

                                int tot_level = 0;
                                int current_labelId = -1;
                                List<DataRow> labels_same_level = new List<DataRow>();
                                List<DataRow> selected_paris_row = new List<DataRow>();
                                exo_1 = new System.Dynamic.ExpandoObject();

                                /* fill table "dt_lookuptab_label" */
                                for (int j = currentSubObj_startIdx; j <= currentSubObj_endIdx; j++)
                                {
                                    splittedArea = ((string)dt_excel.Rows[j]["Area"]).Split(".").ToList();
                                    for (int k = 0; k < splittedArea.Count; k++)
                                    {
                                        if (dt_lookuptab_label.AsEnumerable().Any(row => splittedArea[k] == row.Field<String>("label")) == false)
                                        {
                                            row_tb_lable = dt_lookuptab_label.NewRow();
                                            row_tb_lable["level"] = k;
                                            row_tb_lable["label"] = splittedArea[k];
                                            row_tb_lable["label_complete"] = (string)dt_excel.Rows[j]["Area"];                                                                                      
                                            dt_lookuptab_label.Rows.Add(row_tb_lable);
                                        }
                                    }
                                }

                                /* fill table "dt_lookuptab_pair" */
                                for (int j = currentSubObj_startIdx; j <= currentSubObj_endIdx; j++)
                                {
                                    row_tb_pair = dt_lookuptab_pair.NewRow();
                                    splittedArea = ((string)dt_excel.Rows[j]["Area"]).Split(".").ToList();
                                    corrispond_row = dt_lookuptab_label.AsEnumerable().SingleOrDefault(row => row.Field<string>("label") == splittedArea[splittedArea.Count - 1]);
                                    row_tb_pair["label_id"] = (int)corrispond_row["label_id"];
                                    row_tb_pair["key"] = dt_excel.Rows[j]["Id"];
                                    row_tb_pair["value"] = dt_excel.Rows[j][languageColumn];
                                    dt_lookuptab_pair.Rows.Add(row_tb_pair);
                                }

                                /* complete num_subpair and num_subobject columns in "dt_lookuptab_label" by the help of "dt_lookuptab_pair" */
                                int current_label_id = -1;
                                int current_num_subpairs = -1;
                                int current_level_id = -1;
                                int current_num_subobjects = -1;
                                string current_complete_label = null;
                                foreach (DataRow dr_lookuptab_label in dt_lookuptab_label.Rows)
                                {
                                    current_label_id = (int)dr_lookuptab_label["label_id"];                                                                      
                                    current_num_subpairs = dt_lookuptab_pair.AsEnumerable().Count(row => row.Field<int>("label_id") == current_label_id);
                                    /* 4 means the coloumn of "num_subpair" */
                                    dt_lookuptab_label.Rows[current_label_id][4] = current_num_subpairs;

                                    current_level_id = (int)dr_lookuptab_label["level"];
                                    current_complete_label = (string)dr_lookuptab_label["label_complete"];
                                    current_num_subobjects = dt_lookuptab_label.AsEnumerable().Count(row => row.Field<int>("level") == current_level_id + 1
                                                                                                            && row.Field<string>("label_complete").ToString().Contains(current_complete_label));
                                    /* 5 means the coloumn of "num_subobject" */
                                    dt_lookuptab_label.Rows[current_label_id][5] = current_num_subobjects;
                                }

                                tot_level = dt_lookuptab_label.AsEnumerable().Max(row => row.Field<int>("level"));

                                /*create a copy of the ordered datatable "dt_lookuptab_label" */
                                DataTable dt_lookuptab_label_ordered = dt_lookuptab_label;
                                dt_lookuptab_label_ordered = dt_lookuptab_label_ordered.AsEnumerable()
                                                               .OrderBy(r => r.Field<int>("level"))
                                                               .ThenBy(r => r.Field<string>("label_complete"))
                                                               .CopyToDataTable();

                                Dictionary<string, Object> mainDict = new Dictionary<string, Object>();
                                Dictionary<string, Object> subDict = new Dictionary<string, Object>();
                                DataRow dr_subpair_current, dr_subobject_current;

                                int increment_cnt = 0;
                                string json_prova_2 = null;
                                foreach (DataRow dr_lookuptab_label_ordered in dt_lookuptab_label_ordered.Rows)
                                {                                   

                                    if ((int)dr_lookuptab_label_ordered["level"] == 0)
                                    {
                                        /* add sub pairs*/
                                        IEnumerable<DataRow> current_subpairs = dt_excel.AsEnumerable().Where(dr => dr.Field<string>("area") == (string)dr_lookuptab_label_ordered["label_complete"]);
                                        for (int num_subpair = 0; num_subpair < (int)dr_lookuptab_label_ordered["num_subpair"]; num_subpair++)
                                        {
                                            dr_subpair_current = current_subpairs.ElementAt(num_subpair);
                                            subDict.Add((string)dr_subpair_current[1], (string)dr_subpair_current[2]);
                                            var skl = 2;
                                        }

                                        /* add sub objects*/
                                        string cur_label_complete = (string)dr_lookuptab_label_ordered["label_complete"];
                                        IEnumerable<DataRow> current_subobjects = dt_lookuptab_label_ordered.AsEnumerable().Where(dr => dr.Field<int>("level") == (int)dr_lookuptab_label_ordered["level"] + 1
                                                                                                                                        && dr.Field<string>("label_complete").ToString().Contains(cur_label_complete));
                                        for (int num_subobject = 0; num_subobject < (int)dr_lookuptab_label_ordered["num_subobject"]; num_subobject++)
                                        {
                                            dr_subobject_current = current_subobjects.ElementAt(num_subobject);
                                            subDict.Add((string)dr_subobject_current[2], "");
                                        }

                                        mainDict.Add((string)dr_lookuptab_label_ordered["label"], subDict);
                                        json_prova_2 = JsonConvert.SerializeObject(mainDict, Formatting.Indented);
                                        /* It is important to create a new object rather than Clear it! */
                                        subDict = new Dictionary<string, object>();
                                    }
                                    /* not the root level 0, then traverse all the children*/
                                    else
                                    {
                                        /* add sub pairs*/
                                        IEnumerable<DataRow> current_subpairs = dt_excel.AsEnumerable().Where(dr => dr.Field<string>("area") == (string)dr_lookuptab_label_ordered["label_complete"]);
                                        for (int num_subpair = 0; num_subpair < (int)dr_lookuptab_label_ordered["num_subpair"]; num_subpair++)
                                        {
                                            dr_subpair_current = current_subpairs.ElementAt(num_subpair);
                                            subDict.Add((string)dr_subpair_current[1], (string)dr_subpair_current[2]);
                                            var skl = 2;
                                        }

                                        /* add sub objects*/
                                        string cur_label_complete = (string)dr_lookuptab_label_ordered["label_complete"];
                                        IEnumerable<DataRow> current_subobjects = dt_lookuptab_label_ordered.AsEnumerable().Where(dr => dr.Field<int>("level") == (int)dr_lookuptab_label_ordered["level"] + 1
                                                                                                                                        && dr.Field<string>("label_complete").ToString().Contains(cur_label_complete));
                                        for (int num_subobject = 0; num_subobject < (int)dr_lookuptab_label_ordered["num_subobject"]; num_subobject++)
                                        {
                                            dr_subobject_current = current_subobjects.ElementAt(num_subobject);
                                            subDict.Add((string)dr_subobject_current[2], "");
                                        }
                                        NestedDictIteration(mainDict, (string)dr_lookuptab_label_ordered["label"], subDict);
                                        json_prova_2 = JsonConvert.SerializeObject(mainDict, Formatting.Indented);
                                        subDict = new Dictionary<string, object>();
                                    }

                                    increment_cnt++;                                                                       
                                }                    

                                // loop each label level
                                for (int j = 0; j <= tot_level; j++)
                                    {                                    
                                    labels_same_level = dt_lookuptab_label.Select("level = " + j).ToList();
                                    
                                    // loop the same level has different labels
                                    for (int k = 0; k < labels_same_level.Count(); k++)
                                    {
                                        current_labelId = (int)labels_same_level[k]["label_id"];
                                        selected_paris_row = dt_lookuptab_pair.Select("label_id = " + k).ToList();

                                        // for each label of the specific level, get those (key, value) pairs
                                        for (int m = 0; m < selected_paris_row.Count; m++)
                                        {
                                            ((IDictionary<String, Object>)exo_1).Add((string)selected_paris_row[m]["key"], (string)selected_paris_row[m]["value"]);                                            
                                        }

                                        // add the set of pairs to its parent object label 
                                        ((IDictionary<String, Object>)exo).Add((string)labels_same_level[k]["label"], exo_1);
                                        exo_1 = new System.Dynamic.ExpandoObject();

                                        // please refer to: https://stackoverflow.com/questions/10252675/create-json-dynamically-in-c-sharp

                                        //var sk = exo[28];
                                        //IDictionary<String, Object> hao = exo;
                                        //hao["supportUser"] = new Dictionary<String, Object>();
                                        //hao["supportUser"].add
                                        //hao["supportUser"]["yang"] = "haiyan";
                                        //.Add("quan", "hao");
                                        //var q = hao["supportUser"];

                                        //object r = (IDictionary<String, Object>)exo.FirstOrDefault(row => row[0] == "requestsTable").Value;

                                        // for debug
                                        string tryjson = Newtonsoft.Json.JsonConvert.SerializeObject(exo, Formatting.Indented);
                                        var s = 1;
                                    }
                                    
                                    
                                }

                                // move index to the final of this first subobject
                                i = currentSubObj_endIdx;

                                //foreach (DataRow firstAreaObj in dt_support.Rows) {                                    
                                //    for (int fobj_idx = (int)firstAreaObj["start_idx"]; i <= (int)firstAreaObj["end_idx"]; i++)
                                //    {
                                //        var s = ;
                                //    }
                                //}
                   

                                //List<string> listArea = ((string)current_row["Area"]).Split('.').ToList<string>();
                                ////List<string> listArea = ((string)row["Area"]).Split('.').Reverse().ToList<string>();
                                //int cnt_subobj = listArea.Count;
                                ////RecursiveParseExcelToJson(listArea, id_string, value_string);

                                //foreach (string areaElement in listArea)
                                //{
                                //    // the next row record has the same area with the current row, add pairs to the same subobject "exo_1"
                                //    if ((string)next_row["Area"] == ((string)current_row["Area"]) && (listArea.Count == 1))
                                //    {
                                //        ((IDictionary<String, Object>)exo_1).Add(id_string, value_string);
                                //    }
                                //    // the next row record has the different area with the current row, so we can add the subobject "exo_1" to the parent node "exo"
                                //    else
                                //    {
                                //        ((IDictionary<String, Object>)exo_1).Add(id_string, value_string);
                                //        ((IDictionary<String, Object>)exo).Add(areaElement, exo_1);
                                //        exo_1 = new System.Dynamic.ExpandoObject();
                                //    }
                                //}
                            }                                

                            }
                            //((IDictionary<String, Object>)exo_1).Add("subkey", "subvalore");
                            //((IDictionary<String, Object>)exo).Add("sotto", exo_1);

                            string e_json = Newtonsoft.Json.JsonConvert.SerializeObject(exo, Formatting.Indented);
                            var sqq = 2;
                                  
                        }
                        var sss = 1;
                        Newtonsoft.Json.JsonConvert.SerializeObject(exo);
                        
                    

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

                /* close ACE.OLEDB connection*/
                //oleExcelConnection.Close();

                /*close excel reader stream*/
                excelReader.Close();
            }
            
        }

    public static void NestedDictIteration(Dictionary<string, object> nestedDict, string target_key, Dictionary<string, object> subDict)
    {
        foreach (string key in nestedDict.Keys)
        {
            Console.WriteLine(key);
            object nextLevel_value = nestedDict[key];
            //if (nextLevel == null || nextLevel.GetType() == typeof(string))
            if (nextLevel_value.GetType() == typeof(string) && nextLevel_value.ToString().Length != 0)
            {
                continue;
            }
            else if (nextLevel_value.GetType() == typeof(string) && nextLevel_value.ToString().Length == 0 && key != target_key)
            {
                continue;
            }
            else if (nextLevel_value.GetType() == typeof(string) && nextLevel_value.ToString().Length == 0 && key == target_key)
            {
                    nestedDict[key] = subDict;                   
                    return;
            }           
           
            NestedDictIteration((Dictionary<string, object>)nextLevel_value, target_key, subDict);
            
            
        }
    }


        public static void RecursiveParseExcelToJson(List<string> lstA, string idstr, string valuestr)
        {
            dynamic exo_recursive = new System.Dynamic.ExpandoObject();
            if (lstA.Count == 1)
            {
                ((IDictionary<String, Object>)exo_recursive).Add(idstr, valuestr);
            }
            else
            {   
                lstA.RemoveAt(0);
                RecursiveParseExcelToJson(lstA, idstr, valuestr);
            }
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
            string path_dir_resources = @"C:\Users\quan\Documents\project_2023\parser\parser_resources_json";
            // example RESX
            //string path_dir_resources = @"C:\Users\quan\Documents\project_2023\parser\parser_resources_resx";
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

                    //string innerkey = str.Substring(0, str.LastIndexOf(':'));
                    string innerkey = str.Substring(0, str.IndexOf(':'));
                    innerkey = innerkey.Replace(" ", string.Empty);
                    innerkey = innerkey.Replace(@"""", string.Empty);
                    //string innerValue = str.Substring(str.LastIndexOf(':') + 2);
                    string innerValue = str.Substring(str.IndexOf(':') + 2);
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
        

        public class JsonObject
        {
            public string label { get; set; }
            public string complete_label { get; set; }
            public List<Item> pairItem_list { get; set; }
            public List<JsonObject> jSubObj_list { get ; set; }
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


        //public class NestedDictionary<K, V> : Dictionary<K, NestedDictionary<K, V>>
        //{
        //    public V value { set; get; }

        //    public new NestedDictionary<K, V> this[K key]
        //    {
        //        set { base[key] = value;
        //        }

        //        get
        //        {
        //            if (!base.Keys.Contains<K>(key))
        //            {
        //                base[key] = new NestedDictionary<K, V>();
        //            }                    

        //            return base[key];
        //        }
        //    }
        //}
    }
}


