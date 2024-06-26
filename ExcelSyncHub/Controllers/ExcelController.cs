﻿using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Net;
using Microsoft.AspNetCore.Cors;
using System.Data;
using System.Drawing;
using System.Text;
using DBUtilityHub.Data;
using Spire.Xls;
using ExcelSyncHub.Model.DTO;
using DBUtilityHub.Models;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using ExcelSyncHub.Models.DTO;
using Newtonsoft.Json;
using ExcelSyncHub.Service.IService;

namespace ExcelSyncHub.Controllers
{
    public class ExcelController : Controller
    {
        private readonly IExcelService _excelService;
        protected DBUtilityHub.Models.APIResponse _response;
        private readonly ApplicationDbContext _context;
        public ExcelController(IExcelService excelService, ApplicationDbContext context)
        {
            _excelService = excelService;
            _response = new();
            _context = context;
        }

        [HttpPost("generate")]
        public IActionResult GenerateExcelFile([FromBody] List<ColumnMetaDataDTO> columns,int? logId)
        {
            try
            {
                int? parentId = logId;

                byte[] excelBytes = _excelService.GenerateExcelFile(columns, parentId);
                var fileContentResult = new FileContentResult(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    FileDownloadName = "columns.xlsx"
                };

                return fileContentResult;
            }
            catch (Exception ex)
            {
                var apiResponse = new DBUtilityHub.Models.APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessage = new List<string> { ex.Message },
                    Result = null
                };
                return StatusCode((int)HttpStatusCode.InternalServerError, apiResponse);
            }
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadFile(IFormFile file, [FromQuery] DBConnectionDTO connectionDto,string tableName)
        {
            List<string> errorMessages = new List<string>();
            string successMessage = null;
            try
            {
                var columnsDTO = _excelService.GetColumnsForEntity(tableName).ToList();
                int? uploadedEntityId = null;
                uploadedEntityId = _excelService.GetEntityIdByEntityNamefromui(tableName);
                var idfromtemplatesheet1 = _excelService.GetEntityIdFromTemplate(file, 0); // Sheet 1

                var idfromtemplatesheet2 = _excelService.GetEntityIdFromTemplate(file, 1); // Sheet 2

                if (idfromtemplatesheet1 != uploadedEntityId && idfromtemplatesheet2 != uploadedEntityId)
                {
                    _response.StatusCode = HttpStatusCode.BadRequest;
                    _response.IsSuccess = false;
                    _response.ErrorMessage.Add("Uploaded excel file is not valid, use template file to upload the data");
                    return BadRequest(_response);
                }

                var excelData = _excelService.ReadExcelFromFormFile(file);
                if (excelData == null)
                {
                    _response.StatusCode = HttpStatusCode.BadRequest;
                    _response.IsSuccess = false;
                    _response.ErrorMessage.Add("No valid data found in the uploaded file.");
                    return BadRequest(_response);
                }
                if (file == null || file.Length == 0)
                {
                    _response.StatusCode = HttpStatusCode.BadRequest;
                    _response.IsSuccess = false;
                    _response.ErrorMessage.Add("No file uploaded.");
                    return BadRequest(_response);
                }
                string fileName = file.FileName;
                if (string.IsNullOrEmpty(tableName))
                {
                    _response.StatusCode = HttpStatusCode.BadRequest;
                    _response.IsSuccess = false;
                    _response.ErrorMessage.Add("Table name is required.");
                    return BadRequest(_response);
                }

                bool checkingtableName = _excelService.TableExists(tableName);
                if (checkingtableName == false)
                {
                    _response.StatusCode = HttpStatusCode.BadRequest;
                    _response.IsSuccess = false;
                    _response.ErrorMessage.Add("Table does not exist");
                    return BadRequest(_response);
                }
                DataTable validRowsDataTable = excelData.Clone(); // Create a DataTable to store valid rows
                DataTable successdata = validRowsDataTable.Clone(); // Create a DataTable to store valid rows
                List<string> badRows = new List<string>();
                List<string> filedatas = new List<string>();

                List<string> ErrorRowNumber = new List<string>();
                List<string> errorcolumnnames = new List<string>();
                List<string> columns = new List<string>();
                int total_count = excelData.Rows.Count;
                string comma_separated_string = null;
                List<string> badRowsPrimaryKey = null;
                ValidationResultData validationResultData = null; // Declare it outside the using block
                string columnName = string.Empty;
                string primaryKey = null;
                foreach (var val in columnsDTO)
                {
                    if (val.IsPrimaryKey == true)
                    {
                        primaryKey = val.ColumnName;
                    }
                }

                using (var excelFileStream = file.OpenReadStream())
                {
                    var data = _excelService.ReadDataFromExcel(excelFileStream, excelData.Rows.Count);
                    if (data == null || data.Count == 0)
                    {
                        _response.StatusCode = HttpStatusCode.NoContent;
                        _response.ErrorMessage.Add($"No data found in the '{tableName}' template");
                        _response.IsSuccess = false;
                        return Ok(_response);
                    }

                    // Get the columns from the first row
                    var columnnames = data.First().Keys.ToList();
                    columns = columnnames.ToList();
                    comma_separated_string = string.Join(",", columns.ToArray());

                    // Validate NotNull
                    ValidationResultData validationResult = await _excelService.ValidateNotNull(excelData, columnsDTO);

                    if (validationResult.BadRows.Count > 0)
                    {
                        var resultparams = await _excelService.resultparams(validationResult, comma_separated_string);

                        if (resultparams != null)
                        {
                            filedatas.Add(resultparams.Filedatas);
                            errorMessages.Add(resultparams.errorMessages);
                            ErrorRowNumber.Add(resultparams.ErrorRowNumber);
                        }
                    }

                    //DataTypeValidationResult dataTypeValidationResult = _excelService.ValidateDataTypes(validationResult, columnsDTO);

                    //Primary Kye Validation

                    validationResultData = await _excelService.ValidatePrimaryKeyAsync(validationResult, columnsDTO, tableName,connectionDto);

                    if (validationResultData.BadRows.Count > 0)
                    {
                        var resultparams = await _excelService.resultparamsforprimary(validationResultData, comma_separated_string, tableName);

                        if (resultparams != null)
                        {
                            filedatas.Add(resultparams.Filedatas);
                            errorMessages.Add(resultparams.errorMessages);
                            ErrorRowNumber.Add(resultparams.ErrorRowNumber);
                        }
                    }

                    //Range Validation

                    validationResultData = await _excelService.ValidateRange(validationResultData, columnsDTO, tableName);

                    if (validationResultData.BadRows.Count > 0)
                    {
                        var resultparams = await _excelService.resultparamsforrange(validationResultData, comma_separated_string, tableName);

                        if (resultparams != null)
                        {
                            filedatas.Add(resultparams.Filedatas);
                            errorMessages.Add(resultparams.errorMessages);
                            ErrorRowNumber.Add(resultparams.ErrorRowNumber);
                        }
                    }

                    validationResultData = await _excelService.ValidateLength(validationResultData, columnsDTO, tableName);

                    if (validationResultData.BadRows.Count > 0)
                    {
                        var resultparams = await _excelService.resultparamsforlength(validationResultData, comma_separated_string, tableName);

                        if (resultparams != null)
                        {
                            filedatas.Add(resultparams.Filedatas);
                            errorMessages.Add(resultparams.errorMessages);
                            ErrorRowNumber.Add(resultparams.ErrorRowNumber);
                        }
                    }

                }

                var result = await _excelService.Createlog(tableName, filedatas, fileName, validationResultData.SuccessData.Rows.Count, errorMessages, total_count, ErrorRowNumber);

                // Build the values for the SQL INSERT statement
                _excelService.InsertDataFromDataTableToPostgreSQL(validationResultData.SuccessData, tableName, columns, file, connectionDto);
                if (validationResultData.SuccessData.Rows.Count == 0)
                {
                    _response.Result = result.LogParentDTOs.ID;
                    _response.StatusCode = HttpStatusCode.BadRequest;
                    _response.IsSuccess = false;
                    _response.ErrorMessage.Add("All Records are incorrect");
                    return Ok(_response);
                }
                else if (filedatas.Count == 0)
                {
                    _response.Result = result.LogParentDTOs.ID;
                    _response.StatusCode = HttpStatusCode.Created;
                    _response.IsSuccess = true;
                    _response.ErrorMessage.Add("All records are successfully stored");
                    return Ok(_response);
                }
                else
                {
                    _response.Result = result.LogParentDTOs.ID;
                    _response.StatusCode = HttpStatusCode.BadRequest;
                    _response.IsSuccess = true;
                    _response.ErrorMessage.Add("Passcount records are successfully stored failcount records are incorrect ");
                    return Ok(_response);
                }
            }
            catch (Exception ex)
            {
                var response = new DBUtilityHub.Models.APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false
                };
                response.ErrorMessage.Add(ex.Message);
                return StatusCode((int)HttpStatusCode.InternalServerError, response);
            }

        }
    }
}
