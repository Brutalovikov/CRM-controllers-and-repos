using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using AspNetCoreVueStarter.Models;
using AutoMapper;
using AspNetCoreVueStarter.DTOs;
using AspNetCoreVueStarter.Repositories;
using Newtonsoft.Json;
using AspNetCoreVueStarter.Entities;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Text;
using Microsoft.AspNetCore.SignalR.Protocol;

namespace AspNetCoreVueStarter.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class HomeController : ControllerBase
    {
        private readonly CRMDBContext _context;
        private readonly IMapper _mapper;
        private readonly ICustRepository _custRepository;

        public HomeController(CRMDBContext context, IMapper mapper, ICustRepository custRepository)
        {
            _context = context;
            _mapper = mapper;
            _custRepository = custRepository;
        }

        // здесь берут все ИКГ для вывода архива
        [HttpGet("ikgGetAll")]
        public async Task<IActionResult> IKGGetAll()
        {
            var data = await _context.IKG.ToArrayAsync();
            for (int i = 0; i < data.Length; i++)
            {
                string date = data[i].Date.ToString("dd.MM.yyyy");
                data[i].Date = Convert.ToDateTime(date).Date;
                Console.WriteLine(date, data[i].Date);
            }
            return Ok(data);
        }

        // а здесь берут только один ИКГ непосредственно для отображения в таблицах
        [HttpPost("ikgGetOne")]
        public async Task<IActionResult> GetCusts([FromBody]IKGDto ikg)
        {
            long[] custsCounter = new long[4];
            int step = 0;
            var rels = await _context.Relations.Where(p => p.id_ikg == ikg.Id).Select(p => p.id_oiltree).ToArrayAsync();

            for (int i = 0; i < rels.Length; i++)
            {
                if (i == 0)
                {
                    custsCounter[step] = rels[i];
                    step++;
                }     
                if(rels[i] != custsCounter[step - 1])
                {
                    custsCounter[step] = rels[i];
                    step++;
                }
            }

            var targetCusts = await _context.OilTree.Where(p => custsCounter.Contains(p.id)).ToArrayAsync();
            return Ok(targetCusts);
        }

        // закрывает ИКГ для редактирования
        [HttpPut("closeSchedule")]
        public async Task<IActionResult> CloseSchedule([FromBody]IKGDto ikg)
        {
            var schedule = await _context.IKG.FindAsync(ikg.Id);
            schedule.Ikg_status = ikg.Status;
            _context.SaveChanges();

            return Ok(schedule);
        }

        // сохраняет новый тайминг, отредактированный через таблицу
        [HttpPut("saveNewTiming")]
        public async Task<IActionResult> SaveNewTiming([FromBody]TimingDto tim)
        {
            double days;
            var timing = await _context.Timing.FindAsync(tim.id_relation);

            if (tim.fact_end == null)
                days = (Convert.ToDateTime(timing.fact_end) - Convert.ToDateTime(tim.fact_start)).TotalDays;
            else
                days = (Convert.ToDateTime(tim.fact_end) - Convert.ToDateTime(timing.fact_start)).TotalDays;

            if (tim.fact_end == null && days >= 0)
                timing.fact_start = Convert.ToDateTime(tim.fact_start);
            else if (days >= 0)
                timing.fact_end = Convert.ToDateTime(tim.fact_end);
            else
                return BadRequest();
            _context.SaveChanges();

            return Ok(timing);
        }

        // не работает, должен был использоваться для пересчёта на основе ГБ
        [HttpPost("uploadDrillAndShow")]
        public async Task<IActionResult> UploadDrillAndShow([FromBody]OilTreeDto cust)
        {
            return Ok("kal");
        }    

        // здесь заполняется табличное представление ИКГ
        [HttpPost("fetchJobs")]
        public async Task<IActionResult> FetchJobs([FromBody]OilTreeDto cust)
        {
            var custId = await _context.OilTree.Where(p => p.name == cust.name).Select(p => p.id).ToArrayAsync();
            var rels = await _context.Relations.Where(p => p.id_oiltree == custId[0]).Where(p => p.id_ikg == cust.ikgnum).ToArrayAsync();

            var lastrels = new Relations[rels.Length];
            var deviations = new int[rels.Length];
            var lasttimingFactStarts = new string[rels.Length]; var lasttimingFactStart = new DateTime[1];
            var lasttimingFactEnds = new string[rels.Length]; var lasttimingFactEnd = new DateTime[1];
            if (cust.ikgnum > 1)
            {
                lastrels = await _context.Relations.Where(p => p.id_oiltree == custId[0]).Where(p => p.id_ikg == (cust.ikgnum - 1)).ToArrayAsync();
            }

            var jobsNames = new string[rels.Length]; var jobsName = new string[1];
            var contractorsNames = new string[rels.Length]; var contractorsName = new string[1];
            var timingFactStarts = new string[rels.Length]; var timingFactStart = new DateTime[1];
            var timingFactEnds = new string[rels.Length]; var timingFactEnd = new DateTime[1];
            string date = string.Empty;

            for (int i = 0; i < rels.Length; i++)
            {
                jobsName = await _context.Jobs.Where(p => p.id == rels[i].id_job).Select(p => p.name).ToArrayAsync();
                jobsNames[i] = jobsName[0];
                if (rels[i].id_contractors == null)
                    contractorsNames[i] = string.Empty;
                else
                {
                    contractorsName = await _context.Contractors.Where(p => p.id == rels[i].id_contractors).Select(p => p.name).ToArrayAsync();
                    contractorsNames[i] = contractorsName[0];
                }

                if (cust.ikgnum > 1)
                {

                    lasttimingFactStart = await _context.Timing.Where(p => p.id_relation == lastrels[i].id).Select(p => p.fact_start).ToArrayAsync();
                    date = lasttimingFactStart[0].ToString("dd.MM.yyyy");
                    lasttimingFactStarts[i] = date;

                    lasttimingFactEnd = await _context.Timing.Where(p => p.id_relation == lastrels[i].id).Select(p => p.fact_end).ToArrayAsync();
                    date = lasttimingFactEnd[0].ToString("dd.MM.yyyy");
                    lasttimingFactEnds[i] = date;

                    double lastdays = (lasttimingFactEnd[0] - lasttimingFactStart[0]).TotalDays;

                    timingFactStart = await _context.Timing.Where(p => p.id_relation == rels[i].id).Select(p => p.fact_start).ToArrayAsync();
                    date = timingFactStart[0].ToString("dd.MM.yyyy");
                    timingFactStarts[i] = date;

                    timingFactEnd = await _context.Timing.Where(p => p.id_relation == rels[i].id).Select(p => p.fact_end).ToArrayAsync();
                    date = timingFactEnd[0].ToString("dd.MM.yyyy");
                    timingFactEnds[i] = date;

                    double nowdays = (timingFactEnd[0] - timingFactStart[0]).TotalDays;
                    deviations[i] = Convert.ToInt32(lastdays) - Convert.ToInt32(nowdays);
                }
                else
                {
                    timingFactStart = await _context.Timing.Where(p => p.id_relation == rels[i].id).Select(p => p.fact_start).ToArrayAsync();
                    date = timingFactStart[0].ToString("dd.MM.yyyy");
                    timingFactStarts[i] = date;
                    timingFactEnd = await _context.Timing.Where(p => p.id_relation == rels[i].id).Select(p => p.fact_end).ToArrayAsync();
                    date = timingFactEnd[0].ToString("dd.MM.yyyy");
                    timingFactEnds[i] = date;
                }   
            }

            var resultIkg = new IKGForTable[rels.Length];
            for (int i = 0; i < resultIkg.Length; i++)
            {
                if(cust.ikgnum > 1)
                {
                    resultIkg[i] = new IKGForTable()
                    {
                        id = i,
                        job = jobsNames[i],
                        contractor = contractorsNames[i],
                        additional = rels[i].additional,
                        start = timingFactStarts[i],
                        end = timingFactEnds[i],
                        deviation = deviations[i],
                        id_relation = rels[i].id,
                    };
                }
                else
                {
                    resultIkg[i] = new IKGForTable()
                    {
                        id = i,
                        job = jobsNames[i],
                        contractor = contractorsNames[i],
                        additional = rels[i].additional,
                        start = timingFactStarts[i],
                        end = timingFactEnds[i],
                        deviation = 0,
                        id_relation = rels[i].id,
                    };
                }
            }
            return Ok(resultIkg);
        }

        // преобразует месяца для записи наименований ИКГ в бд
        public string ChangeMonth(int month_num)
        {
            if (month_num == 1)
                return "january";
            if (month_num == 2)
                return "february";
            if (month_num == 3)
                return "march";
            if (month_num == 4)
                return "april";
            if (month_num == 5)
                return "may";
            if (month_num == 6)
                return "june";
            if (month_num == 7)
                return "july";
            if (month_num == 8)
                return "august";
            if (month_num == 9)
                return "september";
            if (month_num == 10)
                return "october";
            if (month_num == 11)
                return "november";
            if (month_num == 12)
                return "december";
            else
                return string.Empty;
        }

        // здесь апдейтится график, если того пожелает мсье главный инженер
        [HttpPut("updateIkg")]
        public async Task<IActionResult> UpdateIkg([FromBody]IKGDto ikgDto)
        {
            var ikg = await _context.IKG.FindAsync(ikgDto.Id);

            ikg.Name = ikg.Name;
            ikg.Date = DateTime.Now;
            ikg.Ikg_status = ikgDto.Status;
            _context.SaveChanges();

            var last_ikg = await _context.IKG.ToArrayAsync();
            var last_relations = await _context.Relations.Where(p => p.id_ikg == last_ikg[last_ikg.Length - 2].Id).ToArrayAsync();
            var relations = await _context.Relations.Where(p => p.id_ikg == ikgDto.Id).ToArrayAsync();
            for (int i = 0; i < relations.Length; i++)
            {
                relations[i].id_job = last_relations[i].id_job;
                relations[i].id_oiltree = last_relations[i].id_oiltree;
                relations[i].id_contractors = last_relations[i].id_contractors;
                relations[i].additional = last_relations[i].additional;
            }
            _context.SaveChanges();

            var new_relations = await _context.Relations.Where(p => p.id_ikg == last_ikg[last_ikg.Length - 1].Id).Select(p => p.id).ToArrayAsync();
            for (int i = 0; i < relations.Length; i++)
            {
                var last_timings = await _context.Timing.Where(p => p.id == last_relations[i].id).ToArrayAsync();
                var timing = await _context.Timing.Where(p => p.id == relations[i].id).ToArrayAsync();
                timing[0].gen_start = DateTime.Now;
                timing[0].gen_end = DateTime.Now;
                timing[0].fact_start = last_timings[0].fact_start;
                timing[0].fact_end = last_timings[0].fact_end;
            }
            _context.SaveChanges();

            return Ok(ikg);
        }

        // здесь создаётся новый ИКГ по подобию существующего в бд
        [HttpPost("createNewIkg")]
        public async Task<IActionResult> CreateNewIkg([FromBody]IKGDto ikgDto)
        {
            var new_ikg = new IKG();
            int month_num = DateTime.Now.Month;
            string month = ChangeMonth(month_num);
            int counter = 0;
            var ikg = await _context.IKG.Where(p => p.Name == "IKG" + "_" + month).ToArrayAsync();
            if (ikg.Length != 0)
            {
                counter = ikg.Length;
                new_ikg = new IKG()
                {
                    Name = "IKG" + "_" + month + counter.ToString(),
                    Date = DateTime.Now,
                    Ikg_status = ikgDto.Status,
                };
            }
            else
            {
                new_ikg = new IKG()
                {
                    Name = "IKG" + "_" + month,
                    Date = DateTime.Now,
                    Ikg_status = ikgDto.Status,
                };
            }

            await _context.IKG.AddAsync(new_ikg);
            _context.SaveChanges();
            
            var last_ikg = await _context.IKG.ToArrayAsync();
            var relations = await _context.Relations.Where(p => p.id_ikg == last_ikg[last_ikg.Length - 2].Id).ToArrayAsync();
            var list_relations = new Relations[relations.Length];

            for (int i = 0; i < list_relations.Length; i++)
            {
                list_relations[i] = new Relations()
                {
                    id_job = relations[i].id_job,
                    id_oiltree = relations[i].id_oiltree,
                    id_contractors = relations[i].id_contractors,
                    additional = relations[i].additional,
                    id_ikg = last_ikg[last_ikg.Length - 1].Id,
                };
                await _context.Relations.AddAsync(list_relations[i]);
                _context.SaveChanges();
            }

            var new_relations = await _context.Relations.Where(p => p.id_ikg == last_ikg[last_ikg.Length - 1].Id).Select(p => p.id).ToArrayAsync();
            var list_timings = new Timing[relations.Length];
            for (int i = 0; i < relations.Length; i++)
            {
                var timing = await _context.Timing.Where(p => p.id == relations[i].id).ToArrayAsync();
                list_timings[i] = new Timing()
                {
                    id_relation = new_relations[i],
                    gen_start = DateTime.Now,
                    gen_end = DateTime.Now,
                    fact_start = timing[0].fact_start,
                    fact_end = timing[0].fact_end,
                };
                await _context.Timing.AddAsync(list_timings[i]);
                _context.SaveChanges();
            }

            return Ok(new_ikg);
        }

        // загрузка ИКГ из файла эксель
        [HttpPost("pushikg")]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> ToDB_fromExcel([FromForm]IFormFile file)
        {
            List<DictAction> action = new List<DictAction>()
                {
                  new DictAction{ id=1,name="Разработка РД"},
                  new DictAction{ id=2,name="Заявка МТР"},
                  new DictAction{ id=3,name="Подрядчик на ИП выбран"},
                  new DictAction{ id=4,name="Подрядчик на ОБУСТР, НС выбран"},
                  new DictAction{ id=5,name="Подрядчик на ВЛ выбран"},
                  new DictAction{ id=6,name="ИП"},
                  new DictAction{ id=7,name="Мобилизация БУ с учетом демонтажа"},
                  new DictAction{ id=8,name="Бурение"},
                  new DictAction{ id=9,name="Освоение"},
                  new DictAction{ id=10,name="Демонтаж"},
                  new DictAction{ id=11,name="СХЕМА ТРАНСПОРТА И СБОРА НЕФТИ"},
                  new DictAction{ id=14,name="Обустройство КП"},
                  //new DictAction{ id=15,name="ПС"}
              };

            var ms = new MemoryStream();
            // Get temp file name
            var temp = Path.GetTempPath(); // Get %TEMP% path
            var fileName = file.FileName; // Get random file name without extension
            var path = Path.Combine(temp, fileName); // Get random file path

                char[] file_name_string = file.FileName.ToCharArray();
                string file_name = file.FileName;
                string filename_for_base = string.Empty;
            /*  string[] allfiles = Directory.GetFiles(path);
              for (int i = 0; i < allfiles.Length; i++)
              {
                  if (allfiles[i] == (path + "\\" + file.FileName))
                  {
                      file_name = string.Empty;
                      for (int j = 0; j < file_name_string.Length; j++)
                      {
                          if (file_name_string[j] == '.')
                              file_name += "(" + allfiles.Length + ")";
                          file_name += file_name_string[j].ToString();
                      }
                  }
              }
*/
            for (int i = 0; i < file_name_string.Length; i++)
                {
                    if (file_name_string[i] == '.') 
                        break;
                    else
                        filename_for_base += file_name_string[i];
                }

            var stream = new FileStream(path, FileMode.CreateNew);
                await file.CopyToAsync(stream);

            var resEnt = _custRepository.ToBase_fromExcel(action, path, filename_for_base);
            if (resEnt.ResultObject != null)
            {
                resEnt.ResultMessage = "Данные успешно загружены в базу";
            }
            else
            {
                if (string.IsNullOrEmpty(resEnt.ResultMessage))
                {
                    resEnt.ResultMessage = "Ошибка загрузки файла в базу";
                }
                else
                {
                    resEnt.ResultMessage = "Ошибка загрузки файла в базу: " + resEnt.ResultMessage;
                }
            }

            return StatusCode(201, new { file = file.FileName, directory = path, mess = resEnt.ResultMessage, res = resEnt.ResultObject });
        }

        // берём данные для построения ленточной диаграммы здесь
        [HttpPost("getGantt")]
        public async Task<IActionResult> GetIkgCust([FromBody]OilTreeDto cust)
        {
            //Session["currentIgk"] = null;
            var custId = await _context.OilTree.Where(p => p.name == cust.name).Select(p => p.id).ToArrayAsync();
            var resEnt = _custRepository.GetCustActs(Convert.ToInt32(custId[0]), cust.ikgnum); //получили список игк
            if (resEnt.ResultObject == null)
            {
                if (string.IsNullOrEmpty(resEnt.ResultMessage))
                {
                    resEnt.ResultMessage = "Ошибка получения данных по кустам";
                }
                else
                {
                    resEnt.ResultMessage = "Ошибка получения данных по кустам: " + resEnt.ResultMessage;
                }
            }
            else
            {
                var resEntAnalytic = _custRepository.CheckRuleJobs((List<Relation>)resEnt.ResultObject);
                if (resEntAnalytic.ResultObject == null)
                {
                    if (string.IsNullOrEmpty(resEntAnalytic.ResultMessage))
                    {
                        resEnt.ResultMessage = "Ошибка проверки правил";
                    }
                    else
                    {
                        resEnt.ResultMessage = "Ошибка проверки правил: " + resEnt.ResultMessage;
                    }
                }
                else
                {
                    var ganttJson = _custRepository.makeGanttJson(Convert.ToInt32(custId[0]), (List<Relation>)resEnt.ResultObject, (Dictionary<string, string>)resEntAnalytic.ResultObject, null);
                    return StatusCode(201, new { ganttJson });
                }
            }
            return StatusCode(201, new { resEnt.ResultMessage });
        }

        /* public JsonResult UpdateTaskDate(int cusId, string start_date, string end_date, string job_name)
         {
             if (Session["currentIgk"] == null)
             {
                 return Json("Преобразование под новый ГБ не производилось", JsonRequestBehavior.AllowGet);
             }
             var curIgk = (List<Relation>)Session["currentIgk"];
             var upAct = curIgk.Where(x => (x.act.name + ' ' + x.act.additional_info).TrimEnd().Equals(job_name)).Select(x => x.act).Single();
             upAct.timing.fact_start = Convert.ToDateTime(start_date);
             upAct.timing.fact_end = Convert.ToDateTime(end_date);
             curIgk.Where(x => (x.act.name + ' ' + x.act.additional_info).TrimEnd().Equals(job_name)).Select(x => x).Single().act = upAct;
             var resEnt = _custRepository.CheckRuleJobs(curIgk);
             if (resEnt.ResultObject == null)
             {
                 if (string.IsNullOrEmpty(resEnt.ResultMessage))
                 {
                     resEnt.ResultMessage = "Ошибка расчета нового ИКГ";
                 }
                 else
                 {
                     resEnt.ResultMessage = "Ошибка расчета нового ИКГ: " + resEnt.ResultMessage;
                 }
             }
             else
             {
                 //Session["currentIgk"] = res; загружать или нет измененный в память???
                 var resEntOldIgk = _custRepository.GetCustActs(cusId);
                 if (resEntOldIgk.ResultObject == null)
                 {
                     if (string.IsNullOrEmpty(resEntOldIgk.ResultMessage))
                     {
                         resEnt.ResultMessage = "Ошибка получения данных по кусту";
                     }
                     else
                     {
                         resEnt.ResultMessage = "Ошибка получения данных по кусту: " + resEntOldIgk.ResultMessage;
                     }
                 }
                 else
                 {
                     var ganttJson = _custRepository.makeGanttJson(cusId, curIgk, (Dictionary<string, string>)resEnt.ResultObject, (List<Relation>)resEntOldIgk.ResultObject); //---!!!!!!!!!!!!!!!!!!!!!!!
                     return Json(ganttJson, JsonRequestBehavior.AllowGet);
                 }
             }
             return Json(resEnt.ResultMessage, JsonRequestBehavior.AllowGet);
         }*/

   /*     public JsonResult UploadDrillAndShow(int cusId)
        {
            if (Session["drilling_schedule"] == null)
            {
                return Json("Необходимо загрузить новый график бурения!", JsonRequestBehavior.AllowGet);
            }
            if (cusId == 0)
            {
                return Json("Куст не выбран!", JsonRequestBehavior.AllowGet);
            }
            var resEnt = _custRepository.GetCustActs(cusId);
            if (resEnt.ResultObject == null)
            {
                if (string.IsNullOrEmpty(resEnt.ResultMessage))
                {
                    resEnt.ResultMessage = "Ошибка получения данных по кусту";
                }
                else
                {
                    resEnt.ResultMessage = "Ошибка получения данных по кусту: " + resEnt.ResultMessage;
                }
            }
            else
            {
                var oldIkg = (List<Relation>)resEnt.ResultObject;

                var newDrillingSchedule = (List<Relation>)Session["drilling_schedule"];
                var resEntNewIkg = _custRepository.CalculateNewIkg(oldIkg, newDrillingSchedule);
                if (resEntNewIkg.ResultObject == null)
                {
                    if (string.IsNullOrEmpty(resEntNewIkg.ResultMessage))
                    {
                        resEnt.ResultMessage = "Ошибка расчета нового ИКГ";
                    }
                    else
                    {
                        resEnt.ResultMessage = "Ошибка расчета нового ИКГ: " + resEntNewIkg.ResultMessage;
                    }
                }
                else
                {
                    var res = (List<Relation>)resEntNewIkg.ResultObject;

                    var resEntAnalytic = _custRepository.CheckRuleJobs(res);
                    if (resEntAnalytic.ResultObject == null)
                    {
                        if (string.IsNullOrEmpty(resEntAnalytic.ResultMessage))
                        {
                            resEnt.ResultMessage = "Ошибка расчета нового ИКГ";
                        }
                        else
                        {
                            resEnt.ResultMessage = "Ошибка расчета нового ИКГ: " + resEntAnalytic.ResultMessage;
                        }
                    }
                    else
                    {
                        Session["currentIgk"] = res;
                        var ganttJson = _custRepository.makeGanttJson(cusId, res, (Dictionary<string, string>)resEntAnalytic.ResultObject, oldIkg);
                        return Json(ganttJson, JsonRequestBehavior.AllowGet);
                    }
                }
            }
            return Json(resEnt.ResultMessage, JsonRequestBehavior.AllowGet);
        }*/

    }
}
