using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using QR_AUTH.Data;
using System.Security.Cryptography;
using QR_AUTH.DTO;

namespace QR_AUTH.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AuthorizationController : ControllerBase
    {
        private readonly DatabaseContext _dbContext;
        private readonly IConfiguration _configuration;
        public AuthorizationController(DatabaseContext context, IConfiguration configuration)
        {
            _dbContext = context;
            _configuration = configuration;
        }

        // private string GenerateApiKey()
        // {
        //     using var provider = new RNGCryptoServiceProvider();
        //     var bytes = new byte[32];
        //     provider.GetBytes(bytes);
        //     return Convert.ToBase64String(bytes).Replace("/", "").Replace("+", "").Replace("=", "");
        // }

        [HttpPost("login")]
        public async Task<IActionResult> Auth(LoginDTO model)
        {
            // Console.WriteLine(model.login);
            // Console.WriteLine(model.password);
            try
            {
                // if (model.apiKey != null)
                // {
                //     var userApi = await _dbContext.Auths.FirstOrDefaultAsync(x => x.Key == model.apiKey);
                //     if (userApi != null)
                //     {
                //         return Ok(userApi.Login);
                //     }
                //
                //     return Ok(new
                //     {
                //         msg = "Требуется авторизация",
                //         code = 500,
                //     });
                // }

                if (model is { login: not null, password: not null })
                {
                    var loginEnv = _configuration["Login"];
                    var passEnv = _configuration["Password"];
                    if(model.login == loginEnv && model.password == passEnv)
                    {
                        return Ok(new
                        {
                            msg = "Auth OK",
                            code = 200,
                            data = new { model.login }
                        });
                    }
                    //var user = await _dbContext.Auths.FirstOrDefaultAsync(x =>
                    //    x.Login == model.login && x.Password == model.password);
                    //if (user != null)
                    //{
                    //    // var key = GenerateApiKey();
                    //    // user.Key = key;
                    //    // await _dbContext.SaveChangesAsync();

                    //    return Ok(new
                    //    {
                    //        msg = "Auth OK",
                    //        code = 200,
                    //        data = new { model.login }
                    //    });
                    //}

                    return Ok(new
                    {
                        msg = "Неверный логин или пароль",
                        code = 401
                    });
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, new
                {
                    msg = "Internal Server Error",
                    code = 500,
                    error = ex.Message
                });
            }

            return Ok();
        }
    }
}