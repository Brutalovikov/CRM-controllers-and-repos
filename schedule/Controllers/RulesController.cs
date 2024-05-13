using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using AspNetCoreVueStarter.Models;
using AspNetCoreVueStarter.Entities;
using AutoMapper;
using AspNetCoreVueStarter.DTOs;
using AspNetCoreVueStarter.Repositories;


namespace AspNetCoreVueStarter.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class RulesController : ControllerBase
    {
        private readonly CRMDBContext _context;
        private readonly IMapper _mapper;
        private readonly IAuthRepository _repo;

        public RulesController(CRMDBContext context, IMapper mapper, IAuthRepository repo)
        {
            _context = context;
            _mapper = mapper;
            _repo = repo;
        }

        // гет для заполнения таблицы правил
        [HttpGet("getRules")]
        public async Task<IActionResult> GetRules()
        {
            var data = await _context.Rules.ToArrayAsync();
            var jobs = await _context.Jobs.ToArrayAsync();
            var jobs_who = new string[data.Length];
            var jobs_whom = new string[data.Length];
            var key_dates = new string[data.Length];
            var conditions = new string[data.Length];

            for (int i = 0; i < data.Length; i++)
            {
                for (int j = 0; j < jobs.Length; j++)
                {
                    if (data[i].id_job_who == jobs[j].id)
                        jobs_who[i] = jobs[j].name;
                    if (data[i].id_job_whom == jobs[j].id)
                        jobs_whom[i] = jobs[j].name;
                }
                if (data[i].condition == "earlier")
                    conditions[i] = "Раньше";
                if (data[i].condition == "later")
                    conditions[i] = "Позже";
                if (data[i].key_date == "start")
                    key_dates[i] = "Начало";
                if (data[i].key_date == "end")
                    key_dates[i] = "Конец";
            }

            var rules_result = new RulesForUser[data.Length];
            for (int i = 0; i < rules_result.Length; i++)
            {
                rules_result[i] = new RulesForUser()
                {
                    id = data[i].id,
                    id_job_who = jobs_who[i],
                    id_job_whom = jobs_whom[i],
                    condition = conditions[i],
                    days_diff = data[i].days_diff,
                    key_date = key_dates[i],
                    interpretation = data[i].interpretation,
                };
            }

            return Ok(rules_result);
        }

        // здесь берутся работы для более симпатишного отображения и понятного пользователю (хотя, возможно это делается в предыдущем методе)
        [HttpGet("getJobs")]
        public async Task<IActionResult> GetJobs()
        {
            var jobs = await _context.Jobs.ToArrayAsync();
            return Ok(jobs);
        }

        // для обновления правил
        [HttpPut("updateRule")]
        public async Task<IActionResult> UpdateRule([FromBody]RulesDto rulesDto)
        {
            var rule = await _context.Rules.FindAsync(rulesDto.id);
            var jobs = await _context.Jobs.ToArrayAsync();
            long job_who = 0; long job_whom = 0;
            string condition = string.Empty; string key_date = string.Empty;

            for (int i = 0; i < jobs.Length; i++)
            {
                if (jobs[i].name == rulesDto.id_job_who)
                    job_who = jobs[i].id;
                if (jobs[i].name == rulesDto.id_job_whom)
                    job_whom = jobs[i].id;
            }
            if (rulesDto.condition == "Раньше")
                condition = "earlier";
            if (rulesDto.condition == "Позже")
                condition = "later";
            if (rulesDto.key_date == "Начало")
                key_date = "start";
            if (rulesDto.key_date == "Конец")
                key_date = "end";

            rule.id_job_who = job_who;
            rule.id_job_whom = job_whom;
            rule.condition = condition;
            rule.days_diff = Convert.ToInt32(rulesDto.days_diff);
            rule.key_date = key_date;
            rule.interpretation = rulesDto.interpretation;

            _context.SaveChanges();
            return Ok(rulesDto);
        }

        // правила можно и удалять при желании. Это делается в данном методе
        [HttpDelete("deleteRule")]
        public async Task<IActionResult> DeleteRule([FromBody]RulesDto rulesDto)
        {
            var rule = await _context.Rules.FindAsync(rulesDto.id);
            _context.Rules.Remove(rule);
            _context.SaveChanges();
            return Ok(rulesDto);
        }

        // какой crud без инсерта?
        [HttpPost("insertRule")]
        public async Task<IActionResult> InsertRule([FromBody]RulesDto rulesDto)
        {
            var jobs = await _context.Jobs.ToArrayAsync();
            long job_who = 0; long job_whom = 0;
            string condition = string.Empty; string key_date = string.Empty;

            for (int i = 0; i < jobs.Length; i++)
            {
                if (jobs[i].name == rulesDto.id_job_who)
                    job_who = jobs[i].id;
                if (jobs[i].name == rulesDto.id_job_whom)
                    job_whom = jobs[i].id;
            }
            if (rulesDto.condition == "Раньше")
                condition = "earlier";
            if (rulesDto.condition == "Позже")
                condition = "later";
            if (rulesDto.key_date == "Начало")
                key_date = "start";
            if (rulesDto.key_date == "Конец")
                key_date = "end";

            var new_rule = new Rules()
            {
                id = rulesDto.id,
                id_job_who = job_who,
                id_job_whom = job_whom,
                condition = condition,
                days_diff = Convert.ToInt32(rulesDto.days_diff),
                key_date = key_date,
                interpretation = rulesDto.interpretation,
            };

            await _context.Rules.AddAsync(new_rule);
            _context.SaveChanges();
            return Ok(new_rule);
        }
    }
}
