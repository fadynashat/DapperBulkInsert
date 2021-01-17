using Dapper;
using ExcelDataReader;
using ImportEcelSheetApi.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ImportEcelSheetApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelSheetController : ControllerBase
    {
        private ApplicationContext context { get; set; }
        private  IConfiguration configuration { get; }
        public ExcelSheetController(ApplicationContext _context, IConfiguration _Configuration)
        {
            context = _context;
            configuration = _Configuration;
        }
        // GET: api/<ExcelSheetController>
      
            [Route("UploadExcel")]
            [HttpPost]
            public IActionResult ExcelUpload()
            {
             
             var httpRequest = HttpContext.Request;
            if (httpRequest.Form.Files.Count > 0)
            {
                var file = httpRequest.Form.Files[0];
                var folderName = Path.Combine("Files");
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", folderName);
                var fileName = ContentDispositionHeaderValue.Parse(file.ContentDisposition).FileName.Trim('"');
                var fullPath = Path.Combine(pathToSave, fileName);
                using (var stream = new FileStream(fullPath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                using (var stream = new FileStream(fullPath, FileMode.Open,FileAccess.Read))
                {
                    

                    using (IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream)) 
                    {

                        DataSet excelRecords = reader.AsDataSet();
                        reader.Close();

                        var finalRecords = excelRecords.Tables[0];
                        for (int i = 0; i < finalRecords.Rows.Count; i++)
                        {
                            User objUser = new User();
                            objUser.Name = finalRecords.Rows[i][0].ToString();
                            objUser.Email = finalRecords.Rows[i][1].ToString();


                            context.Users.Add(objUser);

                        }

                        int output = context.SaveChanges();
                        if (output > 0)
                        {
                            return Ok("Imported Successfully");
                        }
                        else
                        {
                            return Ok("Imported Successfully");
                        }
                    }
                }
            }

            else
            {
                return BadRequest();
            }
            
                
            }

        [HttpPost("InsertBulk")]
        public IActionResult Post([FromBody] User[] Users)
        {
       
            var sql = @"insert into Users (Name,Email) values 
                     (@Name,@Email)";
            try
            {
                using (var connection = new SqlConnection(configuration.GetConnectionString("MyConnectionString")))
                {

                    var affectedRows = connection.Execute(sql, Users);
                    return Ok(affectedRows);
                }
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }

        }
    }
    }

