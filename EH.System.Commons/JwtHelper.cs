using EH.System.Models.Entities;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Tokens;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Commons
{
    public class JwtHelper
    {
        private readonly IConfiguration configuration;

        public JwtHelper(IConfiguration configuration)
        {
            this.configuration = configuration;
        }

        public string GenerateToken(Sys_Users user)
        {
            var secretKey = configuration.GetSection("Jwt:SecurityKey").Value;
            var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(secretKey));
            var credentials = new SigningCredentials(key, SecurityAlgorithms.HmacSha256);

            var claims = new[] {
                new Claim("UserID",user.ID.ToString()),
                new Claim("UserName",user.UserName),
            };

            var token = new JwtSecurityToken
            (issuer: configuration.GetSection("Jwt:Issuer").Value,
            audience: configuration.GetSection("Jwt:Audience").Value,
            claims: claims,
            notBefore: DateTime.Now,
            expires: DateTime.Now.AddMinutes(Convert.ToInt32(configuration.GetSection("Jwt:ExpirationMinutes").Value)),
            signingCredentials: credentials);

            return new JwtSecurityTokenHandler().WriteToken(token);
        }

        public bool ValidateToken(string token)
        {
            var secretKey = configuration.GetSection("Jwt:SecurityKey").Value;
            var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(secretKey));

            var validationParameters = new TokenValidationParameters
            {
                ValidateIssuer = true,
                ValidIssuer = configuration.GetSection("Jwt:Issuer").Value,
                ValidateAudience = true,
                ValidAudience = configuration.GetSection("Jwt:Audience").Value,
                ValidateIssuerSigningKey = true,
                IssuerSigningKey = key,
                ValidateLifetime = true,
            };

            try
            {
                new JwtSecurityTokenHandler().ValidateToken(token, validationParameters, out _);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public T GetClaimValue<T>(string token, string type)
        {
            var secretKey = configuration.GetSection("Jwt:SecurityKey").Value;
            var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(secretKey));

            var validationParameters = new TokenValidationParameters
            {
                ValidateIssuer = true,
                ValidIssuer = configuration.GetSection("Jwt:Issuer").Value,
                ValidateAudience = true,
                ValidAudience = configuration.GetSection("Jwt:Audience").Value,
                ValidateIssuerSigningKey = true,
                IssuerSigningKey = key,
                ValidateLifetime = true,
            };
            new JwtSecurityTokenHandler().ValidateToken(token, validationParameters, out SecurityToken validatedToken);
            var payload = ((JwtSecurityToken)validatedToken).Payload;
          
            if (payload.TryGetValue(type, out var value) && value is T claimValue)
            {
                return claimValue;
            }
            return default(T);
        }
    }
}
