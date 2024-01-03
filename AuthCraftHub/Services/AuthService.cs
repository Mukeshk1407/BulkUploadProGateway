﻿using AuthCraftHub.Models.DTO;
using System.Security.Claims;
using System.Security.Cryptography;
using DBUtilityHub.Data;
using DBUtilityHub.Models;
using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using System.IdentityModel.Tokens.Jwt;
using Microsoft.IdentityModel.Tokens;

namespace AuthCraftHub.Services
{
    public class AuthService
    {
        private readonly ApplicationDbContext _context;

        public AuthService(ApplicationDbContext context)
        {
            _context = context;
        }

        // Auth JWT
        internal async Task<UserTokenDTO> AuthenticateAsync(LoginDTO user)
        {
            var userModel = await _context.UserEntity.FirstOrDefaultAsync(u => u.Email == user.Email);

            if (userModel == null || !VerifyPassword(userModel.Password, user.Password))
            {
                return null;
            }

            var userDetailsDTO = new UserTokenDTO
            {
                UserName = userModel.Name,
                Token = GenerateJwtToken((UserDTO)userModel)
            };

            return userDetailsDTO;
        }

        private bool VerifyPassword(string storedHashedPassword, string enteredPassword)
        {
            var passwordHasher = new PasswordHasher<UserDTO>();
            var passwordVerificationResult = passwordHasher.VerifyHashedPassword(null, storedHashedPassword, enteredPassword);
            return passwordVerificationResult == PasswordVerificationResult.Success;
        }

        private string HashPassword(string password)
        {
            var passwordHasher = new PasswordHasher<UserDTO>();
            return passwordHasher.HashPassword(null, password);
        }
        private string GenerateJwtToken(UserDTO user)
        {

            var tokenHandler = new JwtSecurityTokenHandler();
            var key = GenerateRandomKey();

            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = new ClaimsIdentity(new[]
                {
                     new Claim(ClaimTypes.Name, user.Id.ToString()),
                     new Claim(ClaimTypes.NameIdentifier, user.Name.ToString()), // Use ClaimTypes.NameIdentifier for userId
                     new Claim(ClaimTypes.Email, user.Email),
                     new Claim(ClaimTypes.Role, GetRoleName(user.RoleId).ToString())
                }),

                Expires = DateTime.UtcNow.AddHours(1),
                SigningCredentials = new SigningCredentials(new SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha256Signature)
            };

            var token = tokenHandler.CreateToken(tokenDescriptor);
            return tokenHandler.WriteToken(token);
        }

        private byte[] GenerateRandomKey()
        {
            var key = new byte[32];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(key);
            }
            return key;
        }

        private string GetRoleName(int roleId)
        {
            string? roleName = _context.RoleEntity.Where(r => r.Id == roleId).Select(r => r.RoleName).FirstOrDefault();
            if (string.IsNullOrEmpty(roleName))
            {
                return string.Empty;
            }
            return roleName;
        }

        
        // User CRUD
        internal async Task<UserDTO> CreateUserAsync(UserDTO userModel)
        {
            var role = await _context.RoleEntity.FirstOrDefaultAsync(r => r.Id == userModel.RoleId);
            bool isEmailExist = await _context.UserEntity.AnyAsync(u => u.Email == userModel.Email);

            if (role != null && isEmailExist == false)
            {
                var newUser = new UserEntity
                {
                    Name = userModel.Name,
                    RoleId = role.Id,
                    Email = userModel.Email,
                    Password = HashPassword(userModel.Password),
                    Phonenumber = userModel.Phonenumber,
                    Gender = userModel.Gender,
                    DOB = userModel.DOB,
                    Status = userModel.Status,
                    CreatedDate = DateTime.UtcNow,
                    UpdatedDate = DateTime.UtcNow
                };

                _context.UserEntity.Add(newUser);
                await _context.SaveChangesAsync();
                return userModel;
            }
            return null;
        }

        internal async Task<IEnumerable<UserDTO>> GetUsersAsync()
        {
            var users = _context.UserEntity.ToList();

            var userDTO = new List<UserDTO>();

            userDTO.AddRange(users.Select(c => (UserDTO)c));

            return userDTO;
        }

        internal async Task<UserDTO?> GetUserAsync(int id)
        {

            var user = await _context.UserEntity.FirstOrDefaultAsync(c => c.Id == id);

            if (user == null)
            {
                return null;
            }

            UserDTO userDTO = (UserDTO)user;

            return userDTO;
        }

        internal async Task<UserDTO?> GetUserAsync(string email)
        {
            var user = await _context.UserEntity.FirstOrDefaultAsync(c => c.Email.Trim().ToLower() == email.Trim().ToLower());


            if (user == null)
            {
                return null;
            }

            UserDTO userDTO = (UserDTO)user;

            return userDTO;
        }

        internal async Task<UserDTO> UpdateUserAsync(UserDTO userTableModelDTO)
        {
            UserEntity user = await _context.UserEntity.FirstOrDefaultAsync(c => c.Id == userTableModelDTO.Id);
            var role = await _context.RoleEntity.FirstOrDefaultAsync(r => r.Id == userTableModelDTO.RoleId);

            if (user != null && role != null)
            {

                userTableModelDTO.Password = HashPassword(userTableModelDTO.Password);
                _context.Entry<UserEntity>(user).CurrentValues.SetValues(userTableModelDTO);
                _context.SaveChanges();
                return userTableModelDTO;
            }

            return null;

        }


        internal List<RoleDTO> GetRoleIdsAndNames()
        {
            try
            {
                var roleData = _context.RoleEntity
                    .Select(role => new RoleDTO
                    {
                        Id = role.Id,
                        RoleName = role.RoleName
                    })
                    .ToList();

                return roleData;
            }
            catch (Exception ex)
            {
                // Log the exception or handle it as needed
                Console.WriteLine($"An error occurred: {ex.Message}");
                // You might want to throw an exception or return an error response here
                return null;
            }
        }


    }
}

