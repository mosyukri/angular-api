using BLL.Model;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Angular_netcore_API.Process
{
    public class ExcelProcess
    {
        private string mockdb = "..\\Angular netcore API\\excelfile\\mockdb.xlsx";
        private FileInfo mckdb;
        private ExcelPackage package;
        private ExcelWorksheet worksheet;
        private int colCount;
        private int rowCount;
        public ExcelProcess()
        {
            mckdb = new FileInfo(mockdb);
            package = new ExcelPackage(mckdb);
        }
        public List<EmployeeModel> readExceltoJson()
        {
            try
            {
                string replaceVar = "replaceVar";
                string fulljs = "";

                using (package)
                {
                    initWorksheet();

                    //get first row column name
                    string json = "";
                    for (int col = 1; col <= colCount; col++)
                    {
                        if (!string.IsNullOrWhiteSpace(worksheet.Cells[1, col].Value?.ToString()))
                            json = json + "\"" + worksheet.Cells[1, col].Value?.ToString().Trim().Replace(" ", "") + "\":\"replaceVar" + col + "\",";
                        else
                        {
                            json = "{" + json.Substring(0, json.Length - 1) + "}";
                            colCount = col; //limit to last column
                            break;
                        }
                    }
                    for (int row = 2; row <= rowCount; row++)
                    {
                        if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Value?.ToString().Trim()))
                        {
                            fulljs = fulljs.Substring(0, fulljs.Length - 1);
                            break;
                        }
                        else
                        {
                            string line = json;
                            for (int col = 1; col <= colCount; col++)
                            {
                                line = line.Replace(replaceVar + col, worksheet.Cells[row, col].Value?.ToString().Trim());
                            }
                            fulljs = fulljs + line + ",";
                        }
                    }
                }

                var e = JsonConvert.DeserializeObject<List<EmployeeModel>>("[" + fulljs + "]");

                return e;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        private void initWorksheet()
        {
            worksheet = package.Workbook.Worksheets[0];
            colCount = worksheet.Dimension.End.Column;  //get Column Count
            rowCount = worksheet.Dimension.End.Row;
        }

        public bool writeToExcel(EmployeeModel em,bool update)
        {
            try
            {
                int insertat = 0;
                using (package)
                {
                    initWorksheet();

                    //find last row or data
                
                    for(int row = 1;row<rowCount;row++)
                    {
                        if (!update) //insert
                        {
                            if (string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Value?.ToString()))
                            {
                                insertat = row;
                                break;
                            }
                        }
                        else
                        {//update
                            if(worksheet.Cells[row, 1].Value?.ToString() == em.employeeNumber)
                            {
                                insertat = row;
                                break;
                            }
                        }
                    }
                    

                    worksheet.Cells[insertat, 1].Value = em.employeeNumber;
                    worksheet.Cells[insertat, 2].Value = em.firstName;
                    worksheet.Cells[insertat, 3].Value = em.lastName;
                    worksheet.Cells[insertat, 4].Value = em.employeeStatus;
                    package.Save();
                }

                    return true;
            }catch(Exception ex)
            { return false;
            }
        }
    }
}
