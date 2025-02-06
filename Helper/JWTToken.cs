using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;

namespace sistem_informasi_produksi_backend.Helper
{
    public class JWTToken
    {
        public string IssueToken(IConfiguration configuration, string username, string role, string nama)
        {
            var securityKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(configuration["Key:jwtKey"]!));
            var credentials = new SigningCredentials(securityKey, SecurityAlgorithms.HmacSha256);
            var claims = new List<Claim>
            {
                new ("username", username),
                new ("role", role),
                new ("nama", nama),
            };

            var token = new JwtSecurityToken(
                issuer: configuration["Key:jwtIssuer"],
                audience: configuration["Key:jwtAudience"],
                claims: claims,
                expires: DateTime.Now.AddDays(1),
                signingCredentials: credentials
            );

            return new JwtSecurityTokenHandler().WriteToken(token);
        }
    }
}
