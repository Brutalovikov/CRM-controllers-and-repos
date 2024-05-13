using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using AspNetCoreVueStarter.Models;
using AutoMapper;
using AspNetCoreVueStarter.Entities;
using System.Reflection;
using System.Data.SqlClient;
using System.Data;
using OfficeOpenXml;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Http;

namespace AspNetCoreVueStarter.Repositories
{
    public class CustomCustRepository : ICustRepository
    {
        private readonly string _SqlConnStr = string.Empty;
        private readonly CRMDBContext _context;

        public CustomCustRepository(CRMDBContext context)
        {
            _context = context;
            _SqlConnStr = _context.Database.GetDbConnection().ConnectionString;
        }

        private static dynamic GetDbValue(IDataRecord reader, int index)
        {
            return reader.IsDBNull(index) ? null : reader[index];
        }

        // это легендарный парсинг экселя, содержащего ИКГ
        public ResultEntity ToBase_fromExcel(List<DictAction> dictActions, string path, string filename)
        {
            var resEnt = new ResultEntity();
            try
            {
                //проверка - а не заполнена ли уже база
                var alreadyInserted = false;
                var query = _context.OilTree.ToArray(); //кусты брал....
                var ikg_query = _context.IKG.ToArray();
                string[] ikg_from_base = new string[ikg_query.Length];
                if (ikg_query.Length > 0)
                {
                    ikg_from_base = _context.IKG.Where(p => p.Name == filename).Select(p => p.Name).ToArray();
                }    
                if ((query.Length > 0) && (ikg_from_base.Length > 0))
                    alreadyInserted = true;
                if (alreadyInserted)
                {
                    resEnt.ResultMessage = "В базу уже был загружен ИГК. Необходимо загружать новые графики бурения";
                    resEnt.ResultObject = null;
                }
                else
                {
                    /*   using (var memoryStream = new MemoryStream())
                       {
                           // Get MemoryStream from Excel file
                           file.CopyTo(memoryStream);
                           // Create a ExcelPackage object from MemoryStream
                           using (ExcelPackage package = new ExcelPackage(memoryStream))
                           {
                               // Get the first Excel sheet from the Workbook
                               ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                           }
                       }*/

                /*    using (var stream = new MemoryStream())
                    {
                        file.CopyTo(stream);
                        using (var package = new ExcelPackage(stream))
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                            var rowCount = worksheet.Dimension.Rows;
                        }
                    }*/

                    var app = new Excel.Application();
                    Workbook wb = app.Workbooks.Open(path);

                    //----------------новый икг----------------//
                    IKG ikg = new IKG()
                    {
                        Name = filename,
                        Date = DateTime.Now,
                        Ikg_status = "Загружен",
                    };
                    _context.IKG.Add(ikg);
                    _context.SaveChanges();

                    //----------------справочник кустов----------------//
                    Worksheet sheet = (Worksheet)wb.Sheets["ККСГ"];

                    Microsoft.Office.Interop.Excel.Range custColumn = (Microsoft.Office.Interop.Excel.Range)sheet.Columns["K"];
                    System.Array myvalues = (System.Array)custColumn.Cells.Value;
                    var CUSTS = myvalues.OfType<object>().Where(x => x.ToString().StartsWith("Куст")).Select(x => x.ToString()).ToArray();
                    var custsInBase = _context.OilTree.ToArray();
                    bool check = true;

                    using (var dbContext = new SqlConnection(_SqlConnStr))
                    {
                        dbContext.Open();
                        string sql = "INSERT INTO OilTree(name) VALUES(@param1)";
                        foreach (var cus in CUSTS)
                        {
                            for (int i = 0; i < custsInBase.Length; i++)
                            {
                                if (custsInBase[i].name == cus)
                                    check = false;
                            }
                            if (check)
                            {
                                using (SqlCommand cmd = new SqlCommand(sql, dbContext))
                                {
                                    cmd.Parameters.Add("@param1", SqlDbType.NVarChar).Value = cus;
                                    cmd.CommandType = CommandType.Text;
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            check = true;
                        }
                        dbContext.Close();
                    };


                    //---------------подрядчики----------------------//
                    custColumn = (Microsoft.Office.Interop.Excel.Range)sheet.Columns["S"];
                    myvalues = (System.Array)custColumn.Cells.Value;
                    var CONTRACTORS = myvalues.OfType<object>().Where(x => !x.ToString().Trim().Equals("")).Select(x => x.ToString()).Distinct().ToList();
                    var contractorsInBase = _context.Contractors.ToArray();

                    using (var dbContext = new SqlConnection(_SqlConnStr))
                    {
                        dbContext.Open();
                        string sql = "INSERT INTO contractors(name) VALUES(@param1)";
                        foreach (var con in CONTRACTORS)
                        {
                            for (int i = 0; i < contractorsInBase.Length; i++)
                            {
                                if (contractorsInBase[i].name == con)
                                    check = false;
                            }
                            if (check)
                            {
                                using (SqlCommand cmd = new SqlCommand(sql, dbContext))
                                {
                                    cmd.Parameters.Add("@param1", SqlDbType.NVarChar).Value = con;
                                    cmd.CommandType = CommandType.Text;
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            check = true;
                        }
                        dbContext.Close();
                    }

                    //----------------работы-----------------------------//

                    custColumn = (Microsoft.Office.Interop.Excel.Range)sheet.Columns["L"];
                    myvalues = (System.Array)custColumn.Cells.Value;
                    //var customJobs = myvalues.OfType<object>().Where(x => !dictActions.Select(o => o.name).Contains(x.ToString()));
                    //delete(except) first 6 items and go INSERT!


                    List<Relation> relations = new List<Relation>();

                    Microsoft.Office.Interop.Excel.Range allFile = (Microsoft.Office.Interop.Excel.Range)sheet.Columns["K:W"];

                    System.Array actions = (System.Array)allFile.Cells.Value;

                    var headerIKG = allFile.Find(@"№ Куста", Missing.Value, XlFindLookIn.xlValues, Missing.Value, Missing.Value, XlSearchDirection.xlNext, false, false, Missing.Value);
                    var bottomIKG = allFile.Find(@"ПНР АГЗУ", Missing.Value, XlFindLookIn.xlValues, Missing.Value, Missing.Value, XlSearchDirection.xlPrevious, false, false, Missing.Value);

                    allFile = sheet.Range[headerIKG.Column + ":" + headerIKG.Row + 1, bottomIKG.Column + 18 + ":" + bottomIKG.Row];//берем рендж (+18) захватывающий в экселе 2 доп истории
                    Microsoft.Office.Interop.Excel.Range firstFind = null;
                    var curFind = allFile.Find(@"Куст", Missing.Value, XlFindLookIn.xlValues, Missing.Value, Missing.Value, XlSearchDirection.xlNext, false, false, Missing.Value);
                    var allRelations = new List<Relation>();
                    int g = 0;
                    Console.WriteLine(curFind.Count);
                    while (curFind != null)
                    {
                        Console.WriteLine("====================================" + g++ + "====================================");
                        if (firstFind == null)
                        {
                            firstFind = curFind;
                        }
                        else if (curFind.get_Address(XlReferenceStyle.xlA1)
                          == firstFind.get_Address(XlReferenceStyle.xlA1))
                        {
                            break;
                        }



                        int counter_row = curFind.Row;
                        string cur_val = "";
                        var lstJobs = new List<Act>();
                        var lstJobs_History = new List<Activity_History>();
                        var relations_of_cur_cust = new List<Relation>();
                        var cur_name_par_act = "";


                        //обработка первой строчки куста (Разработка РД)
                        relations_of_cur_cust.Add(new Relation
                        {
                            act = new Act
                            {
                                contractor = new Contractors
                                {
                                    name = (sheet.Cells[counter_row, curFind.Column + 8] as Excel.Range).Value == null ? "" :
                                    (sheet.Cells[counter_row, curFind.Column + 8] as Excel.Range).Value.ToString()
                                },
                                timing = new Timing
                                {
                                    fact_start = Convert.ToDateTime((sheet.Cells[counter_row, curFind.Column + 4] as Excel.Range).Value),
                                    fact_end = Convert.ToDateTime((sheet.Cells[counter_row, curFind.Column + 5] as Excel.Range).Value)
                                },
                                name = dictActions.Where(x => x.name.Equals((sheet.Cells[counter_row, curFind.Column + 1] as Excel.Range).Value)).Select(x => x.name).First(),
                                comment = (sheet.Cells[counter_row, curFind.Column + 9] as Excel.Range).Value == null ? "" :
                                    (sheet.Cells[counter_row, curFind.Column + 9] as Excel.Range).Value.ToString(),
                                additional_info = "",
                                child_acts = new List<Act>()
                            },
                            cust = new Cust
                            {
                                name = (string)curFind.Value
                            }
                        }); 

                        counter_row++;
                        cur_val = (sheet.Cells[counter_row, curFind.Column + 1] as Excel.Range).Value.ToString();

                        while (cur_val != "achtung" && !cur_val.Equals("Разработка РД"))
                        {
                            Console.WriteLine(counter_row);
                            if (dictActions.Where(x => cur_val.Contains(x.name)).Select(x => x.name).Any()) //проверка на джоб из справочника (типа лайк)
                            {
                                if (!cur_name_par_act.Equals("")) cur_name_par_act = ""; //перебирали вложенные предидущие джобы
                                var additional = cur_val.Replace(dictActions.Where(x => cur_val.StartsWith(x.name)).Select(x => x.name).Single(), "").Trim(' ');
                                if (additional == null) additional = "";
                                relations_of_cur_cust.Add(new Relation
                                {
                                    act = new Act
                                    {
                                        contractor = new Contractors
                                        {
                                            name = (sheet.Cells[counter_row, curFind.Column + 8] as Excel.Range).Value == null ? "" :
                                            (sheet.Cells[counter_row, curFind.Column + 8] as Excel.Range).Value.ToString()
                                        },
                                        timing = new Timing
                                        {
                                            fact_start = Convert.ToDateTime((sheet.Cells[counter_row, curFind.Column + 4] as Excel.Range).Value),
                                            fact_end = Convert.ToDateTime((sheet.Cells[counter_row, curFind.Column + 5] as Excel.Range).Value)
                                        },
                                        //name = dictActions.Where(x => x.name.Equals(cur_val == "" ? sheet.Cells[counter_row, curFind.Column + 1].Value : cur_val)).Select(x => x.name).First(),
                                        //name = sheet.Cells[counter_row, curFind.Column + 1].Value == ""?"": sheet.Cells[counter_row, curFind.Column + 1].Value.Replace(":", ""),
                                        name = additional.Equals("") ? cur_val : cur_val.Replace(additional, "").Trim(' '),
                                        comment = (sheet.Cells[counter_row, curFind.Column + 9] as Excel.Range).Value == null ? "" :
                                            (sheet.Cells[counter_row, curFind.Column + 9] as Excel.Range).Value.ToString(),
                                        additional_info = additional,
                                        child_acts = new List<Act>()
                                    },
                                    cust = new Cust
                                    {
                                        name = (string)curFind.Value
                                    }
                                });
                            }
                            else //джоба нет в основных - значит он вложенный.
                            {
                                if (cur_name_par_act.Equals(""))
                                {                                //только нашли вложенный джоб, родитель - предидущий/ убираем двоеточие
                                    cur_name_par_act = (sheet.Cells[counter_row - 1, curFind.Column + 1] as Excel.Range).Value.ToString();
                                    cur_name_par_act = cur_name_par_act.Replace(":", "");
                                }
                                var additional = "";
                                if (cur_val.StartsWith("Нефтесбор"))
                                {
                                    additional = cur_val.Replace("Нефтесбор", "").Trim(' ');
                                }
                                if (cur_val.StartsWith("ВЛ"))
                                {
                                    additional = cur_val.Replace("ВЛ", "").Trim(' ');
                                }
                                if (cur_val.StartsWith("ПС"))
                                {
                                    additional = cur_val.Replace("ПС", "").Trim(' ');
                                }
                                if (additional == null) additional = "";
                                relations_of_cur_cust.First(x => x.act.name.Equals(cur_name_par_act, StringComparison.OrdinalIgnoreCase)).act.child_acts.Add(
                                        new Act
                                        {
                                            contractor = new Contractors
                                            {
                                                name = (sheet.Cells[counter_row, curFind.Column + 8] as Excel.Range).Value == null ? "" :
                                                (sheet.Cells[counter_row, curFind.Column + 8] as Excel.Range).Value.ToString()
                                            },
                                            timing = new Timing
                                            {
                                                fact_start = Convert.ToDateTime((sheet.Cells[counter_row, curFind.Column + 4] as Excel.Range).Value),
                                                fact_end = Convert.ToDateTime((sheet.Cells[counter_row, curFind.Column + 5] as Excel.Range).Value)
                                            },
                                            name = additional.Equals("") ? cur_val : cur_val.Replace(additional, "").Trim(' '),
                                            additional_info = additional,
                                            comment = (sheet.Cells[counter_row, curFind.Column + 9] as Excel.Range).Value == null ? "" :
                                                (sheet.Cells[counter_row, curFind.Column + 9] as Excel.Range).Value.ToString(),
                                        });
                            }
                            counter_row++;
                            cur_val = (sheet.Cells[counter_row, curFind.Column + 1] as Excel.Range).Value == null ? "achtung" :
                                (sheet.Cells[counter_row, curFind.Column + 1] as Excel.Range).Value.ToString();
                            if (cur_val != "achtung") cur_val = cur_val.Replace(":", "");
                        }
                        Console.WriteLine(relations_of_cur_cust.Count);
                        int b = 0;
                        //Добавляем релейшны в список общий
                        foreach (var currel in relations_of_cur_cust)
                        {
                            allRelations.Add(currel);
                            Console.WriteLine(b++);
                        }
                        curFind = allFile.FindNext(curFind);
                    }

                    //читаем идшники из справочников и добавляем релейшены в базу
                    var DictCust = new List<Cust>();
                    var DictContractors = new List<Contractors>();
                    var DictJobs = new List<Jobs>(); //JOBS NE IMAT id_rule
                    //ikg_from_base = _context.IKG.Find(filename);
                    ikg_query = _context.IKG.ToArray();
                    int ikg_target_number = 0;
                    for (int i = 0; i < ikg_query.Length; i++)
                    {
                        if (ikg_query[i].Name == filename)
                            ikg_target_number = i;
                    }

                    using (var dbContext = new SqlConnection(_SqlConnStr))
                    {
                        dbContext.Open();
                        var sqlquery = new SqlCommand("select * from oiltree", dbContext); //кусты брал....
                        using (var reader = sqlquery.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    DictCust.Add(new Cust
                                    {
                                        id = Convert.ToInt32(GetDbValue(reader, 0)),
                                        name = GetDbValue(reader, 1)
                                    });
                                }
                            }
                            reader.Close();
                        }

                        sqlquery = new SqlCommand("select * from contractors", dbContext); //контракторов брал....
                        using (var reader = sqlquery.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    DictContractors.Add(new Contractors
                                    {
                                        id = Convert.ToInt32(GetDbValue(reader, 0)),
                                        name = GetDbValue(reader, 1)
                                    });
                                }
                            }
                            reader.Close();
                        }

                        sqlquery = new SqlCommand("select * from Jobs", dbContext); //джобы брал....
                        using (var reader = sqlquery.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    DictJobs.Add(new Jobs
                                    {
                                        id = Convert.ToInt32(GetDbValue(reader, 0)),
                                        name = GetDbValue(reader, 1)
                                    });
                                }
                            }
                            reader.Close();
                        }

                        dbContext.Close();
                    };

                    //всё готово для заполнения релейшенов!!!
                    using (var dbContext = new SqlConnection(_SqlConnStr))
                    {
                        foreach (var rel in allRelations)
                        {
                            string sql = "insert into relations(id_job,id_oiltree,id_contractors,additional,id_ikg) values (@param1,@param2,@param3,@param4,@param5); select scope_identity();"; //insert главный джоб
                            var relationId = 0;
                            using (SqlCommand cmd = new SqlCommand(sql, dbContext))
                            {
                                dbContext.Open();
                                cmd.Parameters.Add("@param1", SqlDbType.BigInt).Value = DictJobs.Where(x => x.name.Equals(rel.act.name, StringComparison.OrdinalIgnoreCase)).Select(x => x.id).Single();
                                cmd.Parameters.Add("@param2", SqlDbType.BigInt).Value = DictCust.Where(x => x.name.Equals(rel.cust.name, StringComparison.OrdinalIgnoreCase)).Select(x => x.id).Single();
                                if (!rel.act.contractor.name.Trim().Equals(""))
                                {
                                    cmd.Parameters.Add("@param3", SqlDbType.BigInt).Value = DictContractors.Where(x => x.name.Equals(rel.act.contractor.name, StringComparison.OrdinalIgnoreCase)).Select(x => x.id).Single();
                                }
                                else
                                {
                                    cmd.Parameters.AddWithValue("@param3", DBNull.Value);
                                }
                                cmd.Parameters.Add("@param4", SqlDbType.NVarChar).Value = rel.act.additional_info;
                                cmd.Parameters.Add("@param5", SqlDbType.Int).Value = ikg_query[ikg_target_number].Id;
                                cmd.CommandType = CommandType.Text;
                                relationId = Convert.ToInt32(cmd.ExecuteScalar()); //надо ид свеженький
                                dbContext.Close();
                            }
                            //---------тайминги же! тайминги инсерт!------------------------
                            using (SqlCommand cmd = new SqlCommand("insert into timing(id_relation,gen_start,gen_end,fact_start,fact_end) " +
                                                                    "values (@id_rel,@gen_start,@gen_end,@fact_start,@fact_end)", dbContext))
                            {
                                dbContext.Open();
                                cmd.Parameters.Add("@id_rel", SqlDbType.BigInt).Value = relationId;
                                cmd.Parameters.Add("@gen_start", SqlDbType.DateTime2).Value = DateTime.Now;
                                cmd.Parameters.Add("@gen_end", SqlDbType.DateTime2).Value = DateTime.Now;
                                cmd.Parameters.Add("@fact_start", SqlDbType.DateTime2).Value = rel.act.timing.fact_start;
                                cmd.Parameters.Add("@fact_end", SqlDbType.DateTime2).Value = rel.act.timing.fact_end;
                                //cmd.Parameters.Add("@id_job", SqlDbType.BigInt).Value = DictJobs.Where(x => x.name.Equals(rel.act.name, StringComparison.OrdinalIgnoreCase)).Select(x => x.id).Single();
                                cmd.CommandType = CommandType.Text;
                                cmd.ExecuteNonQuery();
                                dbContext.Close();
                            }

                            if (rel.act.child_acts.Count > 0) //есть чайлд джобы - надо проверять соответствие и заполнять таблицу CustomJobs или тупо инсертить по сто раз(у каждого свой ид)
                                                              //для будущего отображения каких нить вещей - но тогда в конструкторе выбор будет адовый.
                                                              // UPD решено инсертить все джобы в одну таблицу
                            {
                                foreach (var customjob in rel.act.child_acts)
                                {
                                    var result = new List<int>();
                                    using (var sqlquery = new SqlCommand("dbo.Get_JobId", dbContext)) //ищем, есть ли такой джоб в базе
                                    {
                                        dbContext.Open();
                                        sqlquery.CommandType = CommandType.StoredProcedure;
                                        sqlquery.Parameters.Add("@name", SqlDbType.NVarChar).Value = customjob.name;
                                        using (var reader = sqlquery.ExecuteReader())
                                        {
                                            if (reader.HasRows)
                                            {
                                                while (reader.Read())
                                                {
                                                    result.Add(Convert.ToInt32(reader.GetInt64(0)));
                                                }
                                            }
                                            reader.Close();
                                        }
                                        dbContext.Close();
                                    }


                                    if (result.Count == 1) //джоб уже есть в системе используем
                                    {
                                        sql = "insert into relations(id_job,id_oiltree,id_contractors,additional,id_ikg) values (@param1,@param2,@param3,@param4,@param5);select scope_identity();";
                                        using (SqlCommand cmd = new SqlCommand(sql, dbContext))
                                        {
                                            dbContext.Open();
                                            cmd.Parameters.Add("@param1", SqlDbType.BigInt).Value = result.Select(x => x).Single();
                                            cmd.Parameters.Add("@param2", SqlDbType.BigInt).Value = DictCust.Where(x => x.name.Equals(rel.cust.name, StringComparison.OrdinalIgnoreCase)).Select(x => x.id).Single();
                                            if (!rel.act.contractor.name.Trim().Equals(""))
                                            {
                                                cmd.Parameters.Add("@param3", SqlDbType.BigInt).Value = DictContractors.Where(x => x.name.Equals(rel.act.contractor.name, StringComparison.OrdinalIgnoreCase)).Select(x => x.id).Single();
                                            }
                                            else
                                            {
                                                cmd.Parameters.AddWithValue("@param3", DBNull.Value);
                                            }
                                            cmd.Parameters.Add("@param4", SqlDbType.NVarChar).Value = customjob.additional_info;
                                            cmd.Parameters.Add("@param5", SqlDbType.Int).Value = ikg_query[ikg_target_number].Id;
                                            cmd.CommandType = CommandType.Text;
                                            relationId = Convert.ToInt32(cmd.ExecuteScalar());
                                            dbContext.Close();
                                        }
                                        using (SqlCommand cmd = new SqlCommand("insert into timing(id_relation,gen_start,gen_end,fact_start,fact_end) " +
                                                                    "values (@id_rel,@gen_start,@gen_end,@fact_start,@fact_end)", dbContext))
                                        {
                                            dbContext.Open();
                                            cmd.Parameters.Add("@id_rel", SqlDbType.BigInt).Value = relationId;
                                            cmd.Parameters.Add("@gen_start", SqlDbType.DateTime2).Value = DateTime.Now;
                                            cmd.Parameters.Add("@gen_end", SqlDbType.DateTime2).Value = DateTime.Now;
                                            cmd.Parameters.Add("@fact_start", SqlDbType.DateTime2).Value = customjob.timing.fact_start;
                                            cmd.Parameters.Add("@fact_end", SqlDbType.DateTime2).Value = customjob.timing.fact_end;
                                            //cmd.Parameters.Add("@id_job", SqlDbType.BigInt).Value = result.Select(x => x).Single();
                                            cmd.CommandType = CommandType.Text;
                                            cmd.ExecuteNonQuery();
                                            dbContext.Close();
                                        }
                                    }
                                    if (result.Count > 1) //джобов больше одного - не может быть такого
                                    {
                                        //ALARM! ALARM!
                                    }
                                    if (result.Count == 0) //джобов нет - добавляем в таблицу с полем парентид, получаем идшник и добавляем релейшн// UPD этого не может быть, Барон!
                                    {
                                        //ALARM! ALARM!
                                    }
                                    result.Clear();
                                    result = null;
                                }
                            }
                        }
                        dbContext.Close();
                    }

                    wb.Close(0);
                    app.Quit();
                    resEnt.ResultMessage = "Данные успешно загружены в базу";
                    resEnt.ResultObject = true;
                }
            }
            catch (Exception e)
            {
                resEnt.ResultObject = null;
                resEnt.ResultMessage = e.Message;
            }
            return resEnt;
        }

        // не менее легендарный метод (для заполнения актов (для последующего заполнения диаграммы Ганта))
        public ResultEntity GetCustActs(int cusId, int ikgId)
        {
            var resEnt = new ResultEntity();
            try
            {
                var CusActs = new List<Relation>();
                using (var dbContext = new SqlConnection(_SqlConnStr))
                {
                    dbContext.Open();
                    var query = new SqlCommand("select r.id_oiltree, r.id_job, j.parentId, j.name, c.name, t.fact_start,t.fact_end, r.additional, oil.name " +
                                               "from Relations r " +
                                               "left join Jobs j on j.id = r.id_job " +
                                               "left join Contractors c on c.id = r.id_contractors " +
                                               "left join Timing t on t.id_relation = r.id " +
                                               "left join OilTree oil on oil.id = r.id_oiltree " +
                                               "where id_oiltree = @cusId AND id_ikg = @ikgId", dbContext);
                    query.Parameters.Add("@cusId", SqlDbType.BigInt).Value = cusId;
                    query.Parameters.Add("@ikgId", SqlDbType.Int).Value = ikgId;
                    using (var reader = query.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                if (GetDbValue(reader, 2) != 0) //это не вложенный джоб
                                {
                                    CusActs.Add(new Relation
                                    {
                                        act = new Act
                                        {
                                            id = Convert.ToInt32(GetDbValue(reader, 1)),
                                            name = GetDbValue(reader, 3),
                                            contractor = new Contractors
                                            {
                                                name = GetDbValue(reader, 4)
                                            },
                                            timing = new Timing
                                            {
                                                fact_start = GetDbValue(reader, 5),
                                                fact_end = GetDbValue(reader, 6)
                                            },
                                            additional_info = GetDbValue(reader, 7)
                                        },
                                        cust = new Cust
                                        {
                                            id = Convert.ToInt32(GetDbValue(reader, 0)),
                                            name = GetDbValue(reader, 8)
                                        }
                                    });
                                }
                                else // это вложенный джоб
                                {
                                    CusActs.Add(new Relation
                                    {
                                        act = new Act
                                        {
                                            id = Convert.ToInt32(GetDbValue(reader, 1)),
                                            name = GetDbValue(reader, 3),
                                            contractor = new Contractors
                                            {
                                                name = GetDbValue(reader, 4)
                                            },
                                            parent_act = Convert.ToInt32(GetDbValue(reader, 2)),
                                            timing = new Timing
                                            {
                                                fact_start = GetDbValue(reader, 5),
                                                fact_end = GetDbValue(reader, 6)
                                            },
                                            additional_info = GetDbValue(reader, 7)
                                        },
                                        cust = new Cust
                                        {
                                            id = Convert.ToInt32(GetDbValue(reader, 0)),
                                            name = GetDbValue(reader, 8)
                                        }
                                    });
                                }
                            }
                        }
                        reader.Close();
                    }
                    dbContext.Close();
                }
                resEnt.ResultMessage = "";
                resEnt.ResultObject = CusActs;
            }
            catch (Exception e)
            {
                resEnt.ResultObject = null;
                resEnt.ResultMessage = e.Message;
            }

            return resEnt;
        }

        // здесь вычисляются джобы, не соответствующие регламентированным для них правилам
        public ResultEntity CheckRuleJobs(List<Relation> ikg)
        {
            var resEnt = new ResultEntity();
            try
            {
                var res = new Dictionary<string, string>();
                //начитываем справочник правил
                var Rules = new List<AnalyticRule>();
                using (var dbContext = new SqlConnection(_SqlConnStr))
                {
                    dbContext.Open();
                    var query = new SqlCommand("select * from Rules", dbContext);
                    using (var reader = query.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                Rules.Add(new AnalyticRule
                                {
                                    id = Convert.ToInt32(GetDbValue(reader, 0)),
                                    id_job_who = Convert.ToInt32(GetDbValue(reader, 1)),
                                    id_job_whom = Convert.ToInt32(GetDbValue(reader, 2)),
                                    condition = GetDbValue(reader, 3),
                                    days_diff = GetDbValue(reader, 4),
                                    key_date = GetDbValue(reader, 5)
                                });
                            }
                        }
                        reader.Close();
                    }
                    dbContext.Close();
                }
                //применяем правила к джобам
                foreach (var relation in ikg)
                {
                    if ((Rules.Where(x => x.id_job_who == relation.act.id).Any()) && (relation.act.parent_act != null))
                    {
                        var job_who = relation.act.id;
                        var job_whom = Rules.Where(x => x.id_job_who == job_who).Select(x => x.id_job_whom).First();
                        var checkres = false;
                        var message = ""; //потом заполнить мессагу после проверки правила
                                          //сборка и применение правила
                        switch (Rules.Where(x => x.id_job_who == job_who).Select(x => x.condition).First() + ' ' + Rules.Where(x => x.id_job_who == job_who).Select(x => x.key_date).First())
                        {
                            case "earlier start":

                                {
                                    if ((relation.act.timing.fact_end - ikg.Where(x => x.act.id == job_whom).Select(x => x.act.timing.fact_start).First()
                                        ).TotalDays * -1 >
                                        Rules.Where(x => x.id_job_who == job_who).Select(x => x.days_diff).First())
                                    {
                                        checkres = true;
                                    }
                                    else
                                    {
                                        checkres = false;
                                    };
                                    break;
                                }
                            case "earlier end":

                                {
                                    //var abc = relation.act.timing.fact_end;
                                    //var abd = ikg.Where(x => x.act.id == job_whom).Select(x => x.act.timing.fact_end).First();
                                    if ((relation.act.timing.fact_end - ikg.Where(x => x.act.id == job_whom).Select(x => x.act.timing.fact_end).First()
                                        ).TotalDays * -1 >
                                        Rules.Where(x => x.id_job_who == job_who).Select(x => x.days_diff).First())
                                    {
                                        checkres = true;
                                    }
                                    else
                                    {
                                        checkres = false;
                                    };
                                    break;
                                }
                            case "later start":
                                {
                                    if ((ikg.Where(x => x.act.id == job_whom).Select(x => x.act.timing.fact_start).First() -
                                        relation.act.timing.fact_start).TotalDays * -1 >
                                        Rules.Where(x => x.id_job_who == job_who).Select(x => x.days_diff).First())
                                    {
                                        checkres = true;
                                    }
                                    else
                                    {
                                        checkres = false;
                                    };
                                    break;
                                }
                            case "later end":
                                {
                                    if ((ikg.Where(x => x.act.id == job_whom).Select(x => x.act.timing.fact_start).First() -
                                        relation.act.timing.fact_end).TotalDays * -1 >
                                        Rules.Where(x => x.id_job_who == job_who).Select(x => x.days_diff).First())
                                    {
                                        checkres = true;
                                    }
                                    else
                                    {
                                        checkres = false;
                                    };
                                    break;
                                }
                        }
                        if (!checkres)
                        {
                            var relfullname = (relation.act.name + ' ' + relation.act.additional_info).TrimEnd();
                            if (res.ContainsKey(relfullname))    //эта работа уже выпала. просто собрать потом сообщение
                            {
                                res[relfullname] = res[relfullname] + ". Сообщение об ошибке ещё";
                            }
                            else
                            {
                                res.Add(relfullname, "Сообщение об ошибке");
                            }
                        }
                    }
                }
                resEnt.ResultObject = res;
                resEnt.ResultMessage = "";
            }
            catch (Exception e)
            {
                resEnt.ResultObject = null;
                resEnt.ResultMessage = e.Message;
            }

            return resEnt;
        }

        //сборка ответа для морды
        public GanttJson makeGanttJson(int cusId, List<Relation> res, Dictionary<string, string> resanalytic, List<Relation> oldIkgg)
        {
        /*    var oldIkg = new List<Relation>();
            using (var dbContext = new SqlConnection(_SqlConnStr))
            {
                dbContext.Open();
                var query = new SqlCommand("select r.id_oiltree, r.id_job, j.parentId, j.name, c.name, t.fact_start,t.fact_end, r.additional, oil.name " +
                                           "from Relations r " +
                                           "left join Jobs j on j.id = r.id_job " +
                                           "left join Contractors c on c.id = r.id_contractors " +
                                           "left join Timing t on t.id_relation = r.id " +
                                           "left join OilTree oil on oil.id = r.id_oiltree " +
                                           "where id_oiltree = @cusId AND id_ikg = @ikgId", dbContext);
                query.Parameters.Add("@cusId", SqlDbType.BigInt).Value = cusId;
                query.Parameters.Add("@ikgId", SqlDbType.Int).Value = 1;
                using (var reader = query.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            if (GetDbValue(reader, 2) != 0) //это не вложенный джоб
                            {
                                oldIkg.Add(new Relation
                                {
                                    act = new Act
                                    {
                                        id = Convert.ToInt32(GetDbValue(reader, 1)),
                                        name = GetDbValue(reader, 3),
                                        contractor = new Contractors
                                        {
                                            name = GetDbValue(reader, 4)
                                        },
                                        timing = new Timing
                                        {
                                            fact_start = GetDbValue(reader, 5),
                                            fact_end = GetDbValue(reader, 6)
                                        },
                                        additional_info = GetDbValue(reader, 7)
                                    },
                                    cust = new Cust
                                    {
                                        id = Convert.ToInt32(GetDbValue(reader, 0)),
                                        name = GetDbValue(reader, 8)
                                    }
                                });
                            }
                            else // это вложенный джоб
                            {
                                oldIkg.Add(new Relation
                                {
                                    act = new Act
                                    {
                                        id = Convert.ToInt32(GetDbValue(reader, 1)),
                                        name = GetDbValue(reader, 3),
                                        contractor = new Contractors
                                        {
                                            name = GetDbValue(reader, 4)
                                        },
                                        parent_act = Convert.ToInt32(GetDbValue(reader, 2)),
                                        timing = new Timing
                                        {
                                            fact_start = GetDbValue(reader, 5),
                                            fact_end = GetDbValue(reader, 6)
                                        },
                                        additional_info = GetDbValue(reader, 7)
                                    },
                                    cust = new Cust
                                    {
                                        id = Convert.ToInt32(GetDbValue(reader, 0)),
                                        name = GetDbValue(reader, 8)
                                    }
                                });
                            }
                        }
                    }
                    reader.Close();
                }
                dbContext.Close();
            }*/

            var acts = new List<Gantt>();
            var settings = new JsonSerializerSettings { DateFormatString = "dd-MM-yyyy" };
            foreach (var rel in res)
            {
                var color = "";
                if (rel.act.parent_act == 0) { color = resanalytic.ContainsKey((rel.act.name + ' ' + rel.act.additional_info).TrimEnd()) ? "red" : "green"; } //в словаре имя джоба который косяк
                else color = "green";
                //string start = JsonConvert.SerializeObject(rel.act.timing.fact_start, settings).Replace("\"", "");
                string start = rel.act.timing.fact_start.ToString("yyyy-MM-dd");
                string end = rel.act.timing.fact_end.ToString("yyyy-MM-dd");
                acts.Add(new Gantt
                {
                    id = rel.act.additional_info.Equals("") ? rel.act.name : rel.act.name + ' ' + rel.act.additional_info,//acts.Where(x => x.id == rel.act.id).Any() ? rel.act.id + 30 : rel.act.id,
                    text = rel.act.additional_info.Equals("") ? rel.act.name : rel.act.name + ' ' + rel.act.additional_info,
                    start_date = JsonConvert.SerializeObject(start, settings).Replace("\"", ""),
                    end_date = JsonConvert.SerializeObject(end, settings).Replace("\"", ""),
                    color = color,
                    contractor = rel.act.contractor.name == null ? "" : rel.act.contractor.name.ToString(),
                    open = "true"
                });
            }
            // добавляем предидущий икг (ид проставляем родительских джобов)
            if (oldIkgg != null)
            {
                foreach (var rel in oldIkgg)
                {
                    acts.Add(new Gantt
                    {
                        id = rel.act.additional_info.Equals("") ? "Предидущий " + rel.act.name : "Предидущий " + rel.act.name + ' ' + rel.act.additional_info,//acts.Where(x => x.text.Equals((rel.act.name + ' ' + rel.act.additional_info).TrimEnd())).Select(x => x.id + 100).Single(),
                        text = "",//rel.act.additional_info.Equals("") ? "Предидущий "+rel.act.name : "Предидущий " + rel.act.name + ' ' + rel.act.additional_info,
                        start_date = JsonConvert.SerializeObject(rel.act.timing.fact_start, settings).Replace("\"", ""),
                        end_date = JsonConvert.SerializeObject(rel.act.timing.fact_end, settings).Replace("\"", ""),
                        color = "rgb(255, 255, 102)",
                        contractor = "",//rel.act.contractor.name == null ? "" : rel.act.contractor.name.ToString(),
                        parent = rel.act.additional_info.Equals("") ? rel.act.name : rel.act.name + ' ' + rel.act.additional_info,//acts.Where(x => x.text.Equals((rel.act.name + ' ' + rel.act.additional_info).TrimEnd())).Select(x => x.id).Single()
                    });
                }
            }
            var cols = new List<GanttColumns>{
                        new GanttColumns{
                            name = "text",
                            //tree = "true",
                            label = res.Select(x=>x.cust.name).Distinct().Single(),
                            width = "*"
                        },
                        new GanttColumns{
                            name = "start_date",
                            label = "Дата начала",
                            width = "75"
                        },
                        new GanttColumns{
                            name = "end_date",
                            label = "Дата окончания",
                            width = "75"
                        },
                        new GanttColumns{
                            name = "contractor",
                            label = "Подрядчик",
                            width = "80"
                        }
                        };
            var ganttJson = new GanttJson
            {
                datalinks = new GanntDataLinks
                {
                    data = acts,
                    links = { }
                },
                cols = cols,
                curcusid = cusId
            };
            return ganttJson;
        }

        // децимация дат, не соответствующих датам из нового ГБ (пересчёт ИКГ на его основе)
        public ResultEntity CalculateNewIkg(List<Relation> igk, List<Relation> newDrillingSchedule)
        {
            var resEnt = new ResultEntity();
            try
            {
                var neededJobs = new List<string>()
            {
                {"Мобилизация БУ с учетом демонтажа"},
                {"Бурение"},
                {"Освоение"},
                {"Демонтаж"}
            };
                var res = new List<Relation>();
                var curcust = igk.Select(x => x.cust.name).Distinct().Single();
                foreach (var job in igk)
                {
                    var jobAdd = (Relation)job.Clone();
                    switch (job.act.name)
                    {
                        case "Мобилизация БУ с учетом демонтажа":
                            {
                                var cur_team_job = newDrillingSchedule.Where(x => x.cust.name == curcust).Select(x => x).OrderBy(x => x.act.timing.fact_start);
                                var cur_team_job_min = cur_team_job.Where(x => x.act.timing.fact_start == cur_team_job.Select(y => y.act.timing.fact_start).Min()).Select(x => x).Single();

                                var prev_team_job_demon = newDrillingSchedule.Where(x => x.act.name == "Демонтаж" &&
                                x.act.timing.fact_start < cur_team_job_min.act.timing.fact_start).Select(x => x);

                                if (prev_team_job_demon.Count() == 0) // это первый куст для данной бригады - берем первые даты
                                {
                                    jobAdd.act.timing.fact_start = cur_team_job.Where(x => x.act.timing.fact_start == cur_team_job.Where(y => y.act.name.Equals("Монтаж")).Select(y => y.act.timing.fact_start).Min()).Select(x => x.act.timing.fact_start).Single();
                                    jobAdd.act.timing.fact_end = cur_team_job.Where(x => x.act.timing.fact_start == cur_team_job.Where(y => y.act.name.Equals("Монтаж")).Select(y => y.act.timing.fact_start).Min()).Select(x => x.act.timing.fact_end).Single();
                                }
                                else
                                {
                                    var prev_team_job_demon_max = prev_team_job_demon.Where(x => x.act.timing.fact_start == prev_team_job_demon.Select(y => y.act.timing.fact_start).Max()).Single(); //искомое
                                    jobAdd.act.timing.fact_start = prev_team_job_demon_max.act.timing.fact_start;
                                    jobAdd.act.timing.fact_end = cur_team_job_min.act.timing.fact_end;
                                }
                                break;
                            }
                        case "Бурение":
                            {
                                var cur_team_job = newDrillingSchedule.Where(x => x.cust.name == curcust).Select(x => x).OrderBy(x => x.act.timing.fact_start);
                                var cur_team_job_min = cur_team_job.Where(x => x.act.timing.fact_start == cur_team_job.Where(y => y.act.name.Equals("Бурение")).Select(y => y.act.timing.fact_start).Min()).Select(x => x).Single();
                                if (cur_team_job_min.cust.well.type_well.Equals("Водозаборные"))
                                {
                                    //нужна следующая скважина, ибо ннада нефтяную
                                    var nextMinJob = cur_team_job.Where(y => !y.Equals(cur_team_job_min) && y.act.name.Equals("Бурение")).Select(y => y.act.timing.fact_start).Min();
                                    var nextJob = cur_team_job.Where(x => x.act.timing.fact_start == nextMinJob && x.act.name.Equals("Бурение")).Select(x => x).Single();
                                    if (!nextJob.cust.well.type_well.Equals("Водозаборные"))
                                    {
                                        jobAdd.act.timing = nextJob.act.timing;
                                        jobAdd.act.additional_info = "2 скв.";
                                    }
                                    else
                                    {
                                        //не может быть ALARM!!!!
                                    }

                                }
                                else
                                {
                                    //ага, вот оно
                                    jobAdd.act.timing = cur_team_job_min.act.timing;
                                    jobAdd.act.additional_info = "1 скв.";
                                }

                                break;
                            }
                        case "Освоение":
                            {
                                var cur_team_job = newDrillingSchedule.Where(x => x.cust.name == curcust).Select(x => x).OrderBy(x => x.act.timing.fact_start);
                                if (res.Where(x => x.act.name.Contains("Бурение")).Select(x => x.act.additional_info).Single() == "1 скв.")
                                { //первая - нефтяная
                                    var cur_team_job_min = cur_team_job.Where(x => x.act.timing.fact_start == cur_team_job.Where(y => y.act.name.Equals("Освоение")).Select(y => y.act.timing.fact_start).Min()
                                        && x.act.name.Equals("Освоение")).Select(x => x).Single();
                                    jobAdd.act.timing = cur_team_job_min.act.timing;
                                    jobAdd.act.additional_info = "1 скв.";
                                }
                                else
                                {//первая - водозаборная
                                    var cur_team_job_min = cur_team_job.Where(x => x.act.timing.fact_start == cur_team_job.Where(y => y.act.name.Equals("Освоение")).Select(y => y.act.timing.fact_start).Min()
                                    && x.act.name.Equals("Освоение")).Select(x => x).Single();
                                    var nextMinJob = cur_team_job.Where(y => !y.Equals(cur_team_job_min) && y.act.name.Equals("Освоение")).Select(y => y.act.timing.fact_start).Min();
                                    var nextJob = cur_team_job.Where(x => x.act.timing.fact_start == nextMinJob && x.act.name.Equals("Освоение")).Select(x => x).Single();
                                    jobAdd.act.timing = nextJob.act.timing;
                                    jobAdd.act.additional_info = "2 скв.";
                                }
                                break;
                            }
                        case "Демонтаж":
                            {
                                var cur_team_job = newDrillingSchedule.Where(x => x.cust.name == curcust).Select(x => x).OrderBy(x => x.act.timing.fact_start);
                                var cur_demon_job = cur_team_job.Where(x => x.act.name.Equals("Демонтаж")).Select(x => x).Single();
                                jobAdd.act.timing = cur_demon_job.act.timing;
                                break;
                            }
                    }
                    res.Add(jobAdd);
                }
                resEnt.ResultObject = res;
                resEnt.ResultMessage = "";
            }
            catch (Exception e)
            {
                resEnt.ResultMessage = e.Message;
                resEnt.ResultObject = null;
            }
            return resEnt;
        }
    }
}
