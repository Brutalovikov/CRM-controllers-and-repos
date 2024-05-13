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
//using System.Web.Http;

namespace AspNetCoreVueStarter.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class DrillingScheduleController : ControllerBase
    {
        private readonly CRMDBContext _context;
        private readonly IMapper _mapper;
        private readonly IDrillScheduleRepository _drillScheduleRepository;
        //public const string DSsessio = "ds";

        public DrillingScheduleController(CRMDBContext context, IMapper mapper, IDrillScheduleRepository drillScheduleRepository)
        {
            _context = context;
            _mapper = mapper;
            _drillScheduleRepository = drillScheduleRepository;
        }

        // эта штука загружает график бурения в систему
        [HttpPost("pushds")]
        [DisableRequestSizeLimit]
        public async Task<IActionResult> ToDB_GBfromExcel([FromForm]IFormFile file)
        {
            var ms = new MemoryStream();
            // Get temp file name
            var temp = Path.GetTempPath(); // Get %TEMP% path
            var fileName = file.FileName; // Get random file name without extension
            var path = Path.Combine(temp, fileName); // Get random file path

          /*  var filesPath = Directory.GetCurrentDirectory() + "/files";
            var path = Path.GetFullPath(file.FileName);

            if (!System.IO.Directory.Exists(filesPath))
            {
                Directory.CreateDirectory(filesPath);
            }*/

            char[] file_name_string = file.FileName.ToCharArray();
            string file_name = file.FileName;
            string filename_for_base = string.Empty;
           /* string[] allfiles = Directory.GetFiles(filesPath);
            for (int i = 0; i < allfiles.Length; i++)
            {
                if (allfiles[i] == (filesPath + "\\" + file.FileName))
                {
                    file_name = string.Empty;
                    for (int j = 0; j < file_name_string.Length; j++)
                    {
                        if (file_name_string[j] == '.')
                            file_name += "(" + allfiles.Length + ")";
                        file_name += file_name_string[j].ToString();
                    }
                }
            }*/

            for (int i = 0; i < file_name_string.Length; i++)
            {
                if (file_name_string[i] == '.')
                    break;
                else
                    filename_for_base += file_name_string[i];
            }

            var stream = new FileStream(path, FileMode.CreateNew);
                await file.CopyToAsync(stream);

             var resEnt = _drillScheduleRepository.ToBase_GBfromExcel(path, filename_for_base);
             if (resEnt.ResultObject == null)
             {
                 if (string.IsNullOrEmpty(resEnt.ResultMessage))
                 {
                     resEnt.ResultMessage = "Ошибка загрузки графика бурения";
                 }
                 else
                 {
                     resEnt.ResultMessage = "Ошибка загрузки графика бурения: " + resEnt.ResultMessage + Environment.NewLine
                         + "График бурения не загружен: " + resEnt.ResultMessage;
                 }
             }
             else
             {
                //HttpContext context = HttpContext.Current;
                //HttpContext.Session.Set<object>(DSsessio, resEnt.ResultObject);
                 resEnt.ResultMessage = "Новый график бурения успешно загружен";
             }

            //return Ok("df");
        return StatusCode(201, new { resultDS = (List<Relation>)resEnt.ResultObject, file = file.FileName, directory = path, mess = resEnt.ResultMessage });
        }

        // ovde берём года по последнему графику бурения
        [HttpGet("getYears")]
        public async Task<IActionResult> GetYears()
        {
            var all_relations = await _context.Relations.Where(p => p.id_DS != null).ToArrayAsync();
            int? count = 0; int years_id = 0;
            for (int i = 0; i < all_relations.Length; i++)
            {
                if (i == 0 || all_relations[i].id_DS > count)
                    count = all_relations[i].id_DS;
            }

            var relations = await _context.Relations.Where(p => p.id_DS == count).ToArrayAsync();
            var all_timings = await _context.Timing.ToArrayAsync();
            var timings = new Timing[relations.Length];
            int?[] years = new int?[20];

            for (int i = 0; i < relations.Length; i++)
            {
                var timing = await _context.Timing.Where(p => p.id_relation == relations[i].id).ToArrayAsync();
                timings[i] = timing[0];
            }

            bool start_includes = true;
            bool end_includes = true;

            for (int i = 0; i < timings.Length; i++)
            {
                for (int j = 0; j < years.Length; j++)
                {
                    if (timings[i].fact_start.Year == years[j])
                        start_includes = false;
                    if (timings[i].fact_end.Year == years[j])
                    {
                        end_includes = false;
                        break;
                    }
                }
                if(start_includes)
                {
                    years[years_id] = timings[i].fact_start.Year;
                    years_id++;
                }
                if (end_includes && timings[i].fact_end.Year != timings[i].fact_start.Year)
                {
                    years[years_id] = timings[i].fact_end.Year;
                    years_id++;
                }
                start_includes = true;
                end_includes = true;
            }

            years = years.Distinct().ToArray();
            years = years.Where(p => p != null).ToArray();
            int[] result = new int[years.Length];
            int? temp;
            for (int i = 0; i < years.Length - 1; i++)
            {
                for (int j = i + 1; j < years.Length; j++)
                {
                    if (years[i] > years[j])
                    {
                        temp = years[i];
                        years[i] = years[j];
                        years[j] = temp;
                    }
                }
            }
            for (int i = 0; i < result.Length; i++)
                result[i] = Convert.ToInt32(years[i]);

            return Ok(result);
        }

        // здесь берутся данные для построения ГБ с выборкой по году
        [HttpPost("getSkiYear")]
        public async Task<IActionResult> GetSkiYear([FromBody]DrillingScheduleDto ds)
        {
            /*   if (Session["drilling_schedule"] == null)
               {
                   return Json("Необходимо загрузить график бурения", JsonRequestBehavior.AllowGet);
               }*/
            var resEnt = _drillScheduleRepository.GetDrillScheduleYear(ds.Year, ds.resultDS);
            if (resEnt.ResultObject == null)
            {
                if (string.IsNullOrEmpty(resEnt.ResultMessage))
                {
                    resEnt.ResultMessage = "Ошибка получения данных";
                }
                else
                {
                    resEnt.ResultMessage = "Ошибка получения данных: " + resEnt.ResultMessage;
                }
            }
            else
            {
                var jsonRes = (SkiGraphJson)resEnt.ResultObject;
                jsonRes.year = ds.Year;
                return StatusCode(201, new { jsonRes });
            }
            return StatusCode(201, new { resEnt.ResultMessage });
        }
    }
}
