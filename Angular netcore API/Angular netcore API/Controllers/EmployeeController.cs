using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;


using BLL.Model;
using OfficeOpenXml;
using Newtonsoft.Json;
using Angular_netcore_API.Process;
namespace Angular_netcore_API.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class EmployeeController : ControllerBase
    {

        private ExcelProcess excel;

        public EmployeeController()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            excel = new ExcelProcess();
        }

        [HttpGet("get/{id}")]
        public object get(string id)
        {
            try
            {
                //read from excel..find & return
                var ee = excel.readExceltoJson();

                if(ee==null)
                {
                    throw new Exception("fail");
                }
                else
                {
                    return ee.Where(x => x.employeeNumber == id);
                }

                //return new EmployeeModel { EmployeeNumber = "001", EmployeeStatus = "testing", FirstName = "tester", LastName = "double01" };
            }
            catch (Exception ex)
            {
                return new { msg= ex.Message};
            }
        }

        [HttpGet("getall")]

        public object getall()
        {
            try
            {
                //read from excel..find & return
                var ee = excel.readExceltoJson();

                if (ee == null)
                {
                    throw new Exception("fail");
                }
                else
                {
                    return ee;
                }

                //return new EmployeeModel { EmployeeNumber = "001", EmployeeStatus = "testing", FirstName = "tester", LastName = "double01" };
            }
            catch (Exception ex)
            {
                return new { msg = ex.Message };
            }
        }

        [HttpPost("save")]
        public bool save([FromBody] EmployeeModel ee)
        {
            try
            {
                return excel.writeToExcel(ee,false);
            }
            catch
            {
                return false;
            }
        }

        [HttpPut]
        public bool update([FromBody] EmployeeModel ee)
        {
            try
            {
                return excel.writeToExcel(ee,true);
            }
            catch
            {
                return false;
            }
        }
     

       
    }
}
