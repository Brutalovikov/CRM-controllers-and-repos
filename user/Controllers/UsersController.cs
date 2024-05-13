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


namespace AspNetCoreVueStarter.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class UsersController : ControllerBase
    {
        private readonly CRMDBContext _context;
        private readonly IMapper _mapper;
        private readonly IAuthRepository _repo;

        public UsersController(CRMDBContext context, IMapper mapper, IAuthRepository repo)
        {
            _context = context;
            _mapper = mapper;
            _repo = repo;
        }

        // берём узеров для заполнения таблицы
        [HttpGet("getUsers")]
        public async Task<IActionResult> GetUsers()
        {
            var data = await _context.TblUser.ToArrayAsync();
            return Ok(data);
        }

        // обновляем узеров здесь
        [HttpPut("updateUser")]
        public async Task<IActionResult> UpdateUser([FromBody]RegisterDto registerDto)
        {
            var user = await _context.TblUser.FindAsync(registerDto.Id);
            var engaged = await _context.TblUser.FirstOrDefaultAsync(p => p.Email == registerDto.Email);

            registerDto.Email = registerDto.Email.ToLower();
            if ((await _repo.UserExists(registerDto.Email)) && (user.Id != engaged.Id))
                return BadRequest("Пользователь с электронной почтой " + registerDto.Email + " уже зарегистрирован.");
                
            user.Name = registerDto.Name;
            user.Surname = registerDto.Surname;
            user.Login = registerDto.Login;
            user.Email = registerDto.Email;
            user.Phone = registerDto.Phone;
            user.Role = registerDto.Role;
            _context.SaveChanges();
            return Ok(registerDto);
        }

        // а здесь удаляем
        [HttpDelete("deleteUser")]
        public async Task<IActionResult> DeleteUser([FromBody]RegisterDto registerDto)
        {
            var user = await _context.TblUser.FindAsync(registerDto.Id);
            _context.TblUser.Remove(user);
            _context.SaveChanges();
            return Ok(registerDto);
        }
    }
}
