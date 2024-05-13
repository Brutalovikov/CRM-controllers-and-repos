using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using AspNetCoreVueStarter.Entities;
using System.Reflection;
using System.Data.SqlClient;
using System.Data;
using Microsoft.EntityFrameworkCore;
using AspNetCoreVueStarter.Models;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using System.IO;

namespace AspNetCoreVueStarter.Repositories
{
    public class CustomDrillScheduleRepository : IDrillScheduleRepository
    {
        private readonly string _SqlConnStr = string.Empty;
        private readonly CRMDBContext _context;

        public CustomDrillScheduleRepository(CRMDBContext context)
        {
            _context = context;
            _SqlConnStr = _context.Database.GetDbConnection().ConnectionString;
        }

        private static dynamic GetDbValue(IDataRecord reader, int index)
        {
            return reader.IsDBNull(index) ? null : reader[index];
        }

        // париснг файла excel с ГБ
        public ResultEntity ToBase_GBfromExcel(string path, string filename)
        {
            var resEnt = new ResultEntity();
            try
            {
                Application app = new Application();
                Workbook wb = app.Workbooks.Open(path);

                var ds_query = _context.DrillingSchedule.ToArray();
                string[] ds_from_base = new string[ds_query.Length];
                if (ds_query.Length > 0)
                {
                    ds_from_base = _context.DrillingSchedule.Where(p => p.Name == filename).Select(p => p.Name).ToArray();
                }

                //----------------новый гб----------------//
                DrillingSchedule DS = new DrillingSchedule()
                {
                    Name = filename,
                    Date = DateTime.Now,
                };
                _context.DrillingSchedule.Add(DS);
                _context.SaveChanges();

                //----------------čitam novi GB----------------//
                Worksheet sheet = (Worksheet)wb.Sheets["TASK"];
                Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
                var NewDrillingSchedule = new List<Relation>();
                var custsInBase = _context.OilTree.ToArray();
                bool check = true;

                for (int i = 3; i <= xlRange.Rows.Count; i++)
                {
                    check = true;
                    var rel = new Relation
                    {
                        act = new Act(),
                        cust = new Cust()
                    };
                    //============================????????===========================//
                    if ((sheet.Cells[i, 7] as Excel.Range).Value == null) //custs and wells
                    {
                        rel.cust.name = "Не назначено";
                        rel.cust.well = new Well { name = "Пусто" };
                    }
                    else
                    {
                        //1==============================================================//
                        rel.cust.name = Regex.Match((sheet.Cells[i, 7] as Excel.Range).Value.ToString(), @".*(?=[(])").Value.TrimEnd(); //берем регуляркой до скобок - Куст Х (Х скв)
                        var custs = _context.OilTree.ToArray();
                        for (int j = 0; j < custs.Length; j++)
                        {
                            if(custs[j].name == rel.cust.name)
                            {
                                check = false;
                                break;
                            }
                        }
                        if (check)
                        {
                            OilTree cus = new OilTree()
                            {
                                name = rel.cust.name,
                            };
                            _context.OilTree.Add(cus);
                            _context.SaveChanges();
                        }
                        check = true;
                        //2==============================================================//

                        rel.cust.well = new Well
                        {
                            //имя скважины
                            name = (sheet.Cells[i, 8] as Excel.Range).Value == null ? "Нет имени" : (sheet.Cells[i, 8] as Excel.Range).Value.ToString(),
                            //тип скважины
                            type_well = (sheet.Cells[i, 17] as Excel.Range).Value == null ? "Не указано" : (sheet.Cells[i, 17] as Excel.Range).Value.ToString(),
                            productions = new List<Production>()
                        };
                        //1==============================================================//
                        var oiltree = _context.OilTree.Where(p => p.name == rel.cust.name).Select(p => p.id).ToArray();
                        Wells well = new Wells()
                        {
                            name = rel.cust.well.name,
                            id_oiltree = oiltree[0],
                            type_well = rel.cust.well.type_well,
                        };
                        _context.Wells.Add(well);
                        _context.SaveChanges();
                        //2==============================================================//
                        rel.cust.count_wells = Convert.ToInt32(Regex.Match((sheet.Cells[i, 7] as Excel.Range).Value.ToString(), @"(?<=[(])[\d]*").Value);
                    }

                    //имя работы
                    rel.act.name = (sheet.Cells[i, 4] as Excel.Range).Value.ToString();
                    //1==============================================================//
                    var jobs = _context.Jobs.Select(p => p.name).ToArray();
                    for (int j = 0; j < jobs.Length; j++)
                    {
                        if (jobs[j] == rel.act.name)
                        {
                            check = false;
                            break;
                        }
                    }
                    if (check)
                    {
                        Jobs job = new Jobs()
                        {
                            name = rel.act.name,
                            parentId = null,
                        };
                        _context.Jobs.Add(job);
                        _context.SaveChanges();
                    }
                    //2==============================================================//

                    //тип скважины
                    rel.act.comment = (sheet.Cells[i, 17] as Excel.Range).Value.ToString();
                    rel.act.timing = new Timing
                    {
                        //дата начала работы
                        fact_start = Convert.ToDateTime((sheet.Cells[i, 21] as Excel.Range).Value),
                        //дата окончания
                        fact_end = Convert.ToDateTime((sheet.Cells[i, 22] as Excel.Range).Value)
                    };
                    rel.act.contractor = new Contractors
                    {
                        //подрядчик
                        name = (sheet.Cells[i, 24] as Excel.Range).Value == null ? "Не назначено" : (sheet.Cells[i, 24] as Excel.Range).Value.ToString(),
                    };
                    NewDrillingSchedule.Add(rel);
                };

                //заполнение таймингов и отношений
                var DictCust = new List<Cust>();
                var DictContractors = new List<Contractors>();
                var DictJobs = new List<Jobs>(); //JOBS NE IMAT id_rule
                ds_query = _context.DrillingSchedule.ToArray();
                int ds_target_number = 0;
                for (int i = 0; i < ds_query.Length; i++)
                {
                    if (ds_query[i].Name == filename)
                        ds_target_number = i;
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

                using (var dbContext = new SqlConnection(_SqlConnStr))
                {
                    foreach (var rel in NewDrillingSchedule)
                    {
                        string sql = "insert into relations(id_job,id_oiltree,id_contractors,additional,id_ds) values (@param1,@param2,@param3,@param4,@param5); select scope_identity();"; //insert главный джоб
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
                            cmd.Parameters.Add("@param4", SqlDbType.NVarChar).Value = "";
                            cmd.Parameters.Add("@param5", SqlDbType.Int).Value = ds_query[ds_target_number].Id;
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
                    }
                    dbContext.Close();
                }

                //а теперь читаем добычу из второго листа
                sheet = (Worksheet)wb.Sheets["добыча"];
                xlRange = sheet.UsedRange;
                for (int i = 5; i <= xlRange.Rows.Count; i += 3) //прыгаем только по добыче нефти
                {
                    //начата или не начата
                    var currentWellName = (sheet.Cells[i, 2] as Excel.Range).Value.ToString();
                    if (NewDrillingSchedule.Where(x => x.cust.well.name.Equals(currentWellName)).Select(x => x).Any())
                    {
                        var listProds = new List<Production>();
                        for (int z = 4; z < 40; z++)
                        {
                            listProds.Add(new Production
                            {
                                dateprod = Convert.ToDateTime((sheet.Cells[1, z] as Excel.Range).Value),
                                oil = (sheet.Cells[i, z] as Excel.Range).Value == null ? 0 : Convert.ToDouble((sheet.Cells[i, z] as Excel.Range).Value)
                            });
                        }
                        foreach (var wel in NewDrillingSchedule.Where(x => x.cust.well.name.Equals(currentWellName)).Select(x => x.cust.well))
                        {
                            wel.productions = listProds;
                        }
                    }
                };

                wb.Close(0);
                app.Quit();
                resEnt.ResultMessage = "";
                resEnt.ResultObject = NewDrillingSchedule;
            }
            catch (Exception e)
            {
                resEnt.ResultMessage = e.Message;
                resEnt.ResultObject = null;
            }
            return resEnt;
        }

        // реализация выборки данных для построения ГБ по дате(году)
        public ResultEntity GetDrillScheduleYear(string year, List<Relation> drillingSchedule)
        {
            var resEnt = new ResultEntity();
            try
            {
                var jsonRes = new SkiGraphJson();
                jsonRes.textlabels = new List<SkiLabelsDictionary>();
                jsonRes.datasets = new List<SkiGraphEntity>();

                long posY = 1; //позиция на графике
                var settings = new JsonSerializerSettings { DateFormatString = "yyyy/MM/dd" }; //настройки даты для джейсона

                //var SortedDrillScheduleByYear = _context.Timing.Where(p => p.fact_start.Year == Convert.ToInt32(year)).ToList();

                // делаем список с нужной датой. берем нужный год, а так же операции на границе годиков
                var SortedDrillScheduleYear = drillingSchedule.Where(x => x.act.timing.fact_start.Year == Convert.ToInt32(year) ||
                               (x.act.timing.fact_end.Year == Convert.ToInt32(year) && x.act.timing.fact_start.Year == Convert.ToInt32(year) - 1 && x.act.timing.fact_end > new DateTime(Convert.ToInt32(year), 1, 1)) ||
                               x.act.timing.fact_end.Year == Convert.ToInt32(year) + 1 && x.act.timing.fact_start.Year == Convert.ToInt32(year)).Select(x => x);

                //ищем бригады
                var SortedDrillSchedule = SortedDrillScheduleYear.Select(x => x.act.contractor.name).Distinct();
                //var sordrill = SortedDrillSchedule.Where(x => x.Equals("Бригада 13 ЭНГС")).Select(x => x); //для теста на 1 бригаду
                var neededworks = new string[] { "Бурение", "Монтаж", "Демонтаж" };
                foreach (var team in SortedDrillSchedule)
                {
                    var dataset = new SkiGraphEntity();
                    var colors = new List<string>();
                    var data = new List<SkiGraphDataset>();
                    var teamlabels = new List<SkiLabelsText>();
                    int counterForLabel = 0;
                    var demon = false;
                    colors.Add(""); //первый цвет не важен, т.к. генерация цвета нужна для линии, а не для первой точки
                    var SortedDrillDate = SortedDrillScheduleYear.Where(x => x.act.contractor.name.Equals(team) && neededworks.Contains(x.act.name)).Select(x => x).OrderBy(x => x.act.timing.fact_start);
                    foreach (var job in SortedDrillDate)
                    {
                        //лейбл заполнение
                        var labelText = "";
                        if (counterForLabel == 0)
                        {
                            //labelText += job.act.contractor.name + '\n' + job.cust.name + '(' + job.cust.well.name + ')';
                            labelText += job.act.contractor.name + '\n' + job.cust.name + '(' + job.cust.count_wells.ToString() + "скв.)";

                        }
                        else
                        {
                            if (!demon)
                            {
                                if (job.act.name.Equals("Демонтаж", StringComparison.Ordinal))
                                {
                                    demon = true;
                                }
                            }
                            else
                            {
                                //labelText += job.cust.name + '(' + job.cust.well.name + ')';
                                labelText += job.cust.name + '(' + job.cust.count_wells.ToString() + "скв.)";
                                demon = false;
                            }
                        }
                        //точки и цвет заполнение
                        if (job.act.timing.fact_start.Year < Convert.ToInt32(year)) //в прошлом году джоб начинается. рубим на 01.01.year
                        {
                            data.Add(new SkiGraphDataset
                            {
                                x = JsonConvert.SerializeObject(new DateTime(Convert.ToInt32(year), 1, 1), settings).Replace("\"", ""),
                                y = posY
                            });
                        }
                        else
                        {
                            data.Add(new SkiGraphDataset
                            {
                                x = JsonConvert.SerializeObject(job.act.timing.fact_start, settings).Replace("\"", ""),
                                y = posY
                            });
                        }

                        if (job.act.name.Equals("Демонтаж") || job.act.name.Equals("Монтаж")) colors.Add("rgb(0,0,0)"); // монтаж и демонтаж - черный цвет
                        else
                        {
                            if (job.act.comment.Contains("Горизонтальные ПК")) colors.Add("rgb(54, 162, 235)");//blue
                            if (job.act.comment.Contains("Горизонтальные МХ") || job.act.comment.Contains("Горизонтальные БУ")) colors.Add("rgb(75, 192, 192)"); //green
                            if (job.act.comment.Contains("Газовые")) colors.Add("rgb(153, 102, 255)");//purple
                            if (job.act.comment.Contains("Водозаборные")) colors.Add("rgb(255, 205, 86)"); //yellow
                        }

                        counterForLabel++;
                        teamlabels.Add(new SkiLabelsText
                        {
                            index = counterForLabel,
                            label = labelText
                        });
                    }
                    //обработка последнего джоба - так как брал только факт старт + если в следующем году джоб заканчивается, то рубим на 01.01.year++
                    if (SortedDrillDate.Max(x => x.act.timing.fact_end).Year > Convert.ToInt32(year))
                    {
                        data.Add(new SkiGraphDataset
                        {
                            x = JsonConvert.SerializeObject(new DateTime(Convert.ToInt32(year) + 1, 1, 1), settings).Replace("\"", ""),
                            y = posY
                        });
                    }
                    else
                    {
                        data.Add(new SkiGraphDataset
                        {
                            x = JsonConvert.SerializeObject(SortedDrillDate.Max(x => x.act.timing.fact_end), settings).Replace("\"", ""),
                            y = posY
                        });
                    }
                    dataset.data = data; //данные по бригаде добавляем в сущность
                    dataset.colors = colors; //данные по цветам
                                             //dataset.borderColor = "rgb(255, 99, 132)";
                    dataset.borderColor = "rgb(255, 255, 255)";
                    dataset.borderWidth = 15;
                    dataset.backgroundColor = "rgb(255, 255, 255)";
                    dataset.label = team;
                    dataset.radius = 0;
                    dataset.pointHitRadius = 10;
                    dataset.fill = false;
                    dataset.pointStyle = "line";
                    jsonRes.datasets.Add(dataset);
                    jsonRes.textlabels.Add(new SkiLabelsDictionary
                    {
                        team = team,
                        teamlabels = teamlabels
                    });
                    posY += 1;
                }
                resEnt.ResultMessage = "";
                resEnt.ResultObject = jsonRes;
            }
            catch (Exception e)
            {
                resEnt.ResultObject = null;
                resEnt.ResultMessage = e.Message;
            }
            return resEnt;
        }
    }
}
