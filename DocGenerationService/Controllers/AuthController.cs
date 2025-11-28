using connector_with_auth.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using System.Data;
using Microsoft.Data.SqlClient;

namespace connector_with_auth.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AuthController : ControllerBase
    {
        private readonly IConfiguration _config;

        public AuthController(IConfiguration config)
        {
            _config = config;
        }

        [AllowAnonymous]
        [HttpPost("login")]
        public IActionResult Login([FromBody] LoginRequest request)
        {
            // Simple example – replace with real user validation:
            string connectionString = $"Server={_config["database_connection_properties:servername"]};Database={_config["database_connection_properties:database"]};User Id={_config["database_connection_properties:login"]};Password={_config["database_connection_properties:password"]};TrustServerCertificate=True;";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                string query = "SELECT * From apiUsers";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            //int id = reader.GetInt32(0);
                            string username = reader.GetString("username");
                            string password = reader.GetString("password");
                            string role = reader.GetString("roles");

                            if (request.username == username && request.password == password)
                            {
                                var token = GenerateToken(request.username, role);
                                conn.Close();
                                return Ok(new { token });
                            }
                            
                            
                            //Console.WriteLine($"{username} - {password} - {role}");

                        }
                    }
                    conn.Close();
                    return Unauthorized();
                }
                
            }

            /*
            if (request.username != "admin" || request.password != "123")
                return Unauthorized();

            var token = GenerateToken(request.username,"admin");
            return Ok(new { token });
            */
        }

        private string GenerateToken(string username,string role)
        {
            var claims = new[]
            {
                new Claim(JwtRegisteredClaimNames.Sub, username),
                new Claim("role", role)
            };

            var key = new SymmetricSecurityKey(
                Encoding.UTF8.GetBytes(_config["Jwt:Key"]));

            var creds = new SigningCredentials(key, SecurityAlgorithms.HmacSha256);

            var token = new JwtSecurityToken(
                issuer: _config["Jwt:Issuer"],
                audience: _config["Jwt:Audience"],
                claims: claims,
                expires: DateTime.Now.AddHours(1),
                signingCredentials: creds);

            return new JwtSecurityTokenHandler().WriteToken(token);
        }
    }
}
