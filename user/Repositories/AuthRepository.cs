using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using AspNetCoreVueStarter.Models;
using Microsoft.EntityFrameworkCore;
using System.Text;

namespace AspNetCoreVueStarter.Repositories
{
    public class AuthRepository : IAuthRepository
    {
        private readonly CRMDBContext _context;


        public AuthRepository(CRMDBContext context)
        {
            _context = context;
        }

        // здесь отправляется ответ пользователю о успешной/не авторизации
        public async Task<TblUser> Login(string login, string password)
        {
            var user = await _context.TblUser.FirstOrDefaultAsync(x => x.Login == login);
            if (user == null)
                return null;

            if (!VerifyPasswordHash(password, user.Password, user.Salt))
                return null;

            return user;
        }

        // добавляем узера в бд + создаём хеши и соли паролей
        public async Task<TblUser> Register(TblUser user, string password)
        {
            byte[] passwordHash, salt;
            CreatePasswordHash(password, out passwordHash, out salt);

            user.Password = Convert.ToBase64String(passwordHash);
            user.Salt = Convert.ToBase64String(salt);

            //id = _context.TblUser.Max(e => e.Id) + 1;
            //user.Id = id;

            await _context.TblUser.AddAsync(user);
            await _context.SaveChangesAsync();


            return user;
        }

        // верификация паролей
        private bool VerifyPasswordHash(string password, string passwordHash, string salt)
        {
            using (var hmac = new System.Security.Cryptography.HMACSHA512(Convert.FromBase64String(salt)))
            {
                var computedHash = hmac.ComputeHash(Encoding.UTF8.GetBytes(password));
                string hash = Convert.ToBase64String(computedHash);

                char[] computed_h = hash.ToCharArray();
                char[] user_h = passwordHash.ToCharArray();

                for (int i = 0; i < computed_h.Length; i++)
                {
                    if (computed_h[i] != user_h[i]) return false;
                }
            }
            return true;
        }

        // создание непосредственно хеша пароля
        private void CreatePasswordHash(string password, out byte[] passwordHash, out byte[] salt)
        {
            using (var hmac = new System.Security.Cryptography.HMACSHA512())
            {
                salt = hmac.Key;
                passwordHash = hmac.ComputeHash(Encoding.UTF8.GetBytes(password));
            }
        }

        // проверка на мыло (один пользователь = одна электронная почта!)
        public async Task<bool> UserExists(string Username)
        {
            if (await _context.TblUser.AnyAsync(x => x.Email == Username))
                return true;
            return false;
        }
    }
}
