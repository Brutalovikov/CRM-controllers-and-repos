using System;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using AspNetCoreVueStarter.DTOs;
using AspNetCoreVueStarter.Models;
using AspNetCoreVueStarter.Repositories;
using AutoMapper;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;

namespace AspNetCoreVueStarter.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AuthController : ControllerBase
    {
        private readonly IAuthRepository _repo;
        private readonly IConfiguration _config;
        private readonly IMapper _mapper;


        public AuthController(IAuthRepository repo, IConfiguration config, IMapper mapper)
        {
            _mapper = mapper;
            _config = config;
            _repo = repo;
        }

        // метод для регистрации новых лиц
        [Authorize]
        [HttpPost("register")]
        public async Task<IActionResult> Register([FromBody]RegisterDto registerDto)
        {
            registerDto.Email = registerDto.Email.ToLower();
            if (await _repo.UserExists(registerDto.Email))
                return BadRequest("Пользователь с электронной почтой " + registerDto.Email + " уже зарегистрирован.");

            var userToCreate = _mapper.Map<TblUser>(registerDto);
            var createdUser = await _repo.Register(userToCreate, registerDto.Password);

            return StatusCode(201, new { email = createdUser.Email, login = createdUser.Login });
        }

        // метод для авторизации
        [HttpPost("login")]
        public async Task<IActionResult> Login([FromBody]LoginDto loginDto)
        {
            var userFromRepo = await _repo.Login(loginDto.Login.ToLower(), loginDto.Password);
            if (userFromRepo == null)
                return Unauthorized();


            var claims = new[]
            {
                new Claim(ClaimTypes.NameIdentifier, userFromRepo.Id.ToString()),
                new Claim(ClaimTypes.Name, userFromRepo.Email)
            };

            DateTime expires = DateTime.UtcNow.AddMinutes(30);
            var key = new SymmetricSecurityKey(Encoding.UTF8
                .GetBytes(_config.GetSection("AppSettings:Token").Value));
            var creds = new SigningCredentials(key, SecurityAlgorithms.HmacSha512Signature);
            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = new ClaimsIdentity(claims),
                Expires = expires,
                //NotBefore = expires,
                SigningCredentials = creds
            };


            var tokenHandler = new JwtSecurityTokenHandler();
            var token = tokenHandler.CreateToken(tokenDescriptor);


            return Ok(new { token = tokenHandler.WriteToken(token), email = userFromRepo.Email, login = userFromRepo.Login, role = userFromRepo.Role, name = userFromRepo.Name, surname = userFromRepo.Surname });
        }
    }
}
