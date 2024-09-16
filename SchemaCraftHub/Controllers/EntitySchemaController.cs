using System.Net;
using System.Web;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using SchemaCraftHub.Model.DTO;
using SchemaCraftHub.Service;
using SchemaCraftHub.Service.IService;

namespace SchemaCraftHub.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class EntitySchemaController : ControllerBase
    {
        private readonly IEntitySchemaService _entitySchemaService;
        private readonly ILogService _logService; // Inject the LogService

        public EntitySchemaController(IEntitySchemaService entitySchemaService, ILogService logService)
        {
            _entitySchemaService = entitySchemaService;
            _logService = logService;
        }

        [HttpGet("tables")]
        public async Task<IActionResult> GetAllTables()

        {
            try
            {
                var tables = await _entitySchemaService.GetAllTablesAsync();

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = tables
                };
                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };
                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("tables/{id}")]
        public async Task<IActionResult> GetTableById(int id)
        {
            try
            {
                var table = await _entitySchemaService.GetTableByIdAsync(id);
                if (table == null)
                {
                    return NotFound();
                }

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = table
                };
                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };
                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("tables/GetTablesByHostProviderDatabase")] //display list in frontend
        public async Task<IActionResult> GetTablesByHostProviderDatabase(string? hostName, string provider, string? databaseName, string? accessKey, string? secretkey, string? region, string? keyspace, string? ec2Instance, string? ipAddress, string? influxDbToken, string? influxDbOrg, string? influxDbUrl, string? influxDbBucket)
        {
            try
            {
                var tables = await _entitySchemaService.GetTablesByHostProviderDatabaseAsync(hostName, provider, databaseName, accessKey, secretkey, region, keyspace, ec2Instance, ipAddress, influxDbToken,influxDbOrg, influxDbUrl, influxDbBucket);

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = tables
                };

                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };

                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("tables/GetTableByHostProviderDatabaseTableName")]
        public async Task<IActionResult> GetTableByHostProviderDatabaseTableName(string? hostName, string provider, string? databaseName, string? accessKey, string? secretkey, string? region, string? keyspace, string? ec2Instance, string? ipAddress, string? tableName, string? influxDbToken, string? influxDbOrg, string? influxDbUrl, string? influxDbBucket)
        {
            try
            {
                var table = await _entitySchemaService.GetTableByHostProviderDatabaseTableNameAsync(hostName, provider, databaseName, accessKey, secretkey, region, keyspace, ec2Instance, ipAddress, tableName, influxDbToken, influxDbOrg, influxDbUrl, influxDbBucket);
                if (table == null)
                {
                    return NotFound();
                }

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = table
                };

                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };

                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("columns")]
        public async Task<IActionResult> GetAllColumns()
        {
            try
            {
                var columns = await _entitySchemaService.GetAllColumnsAsync();

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = columns
                };
                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };
                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("columns/{id}")]
        public async Task<IActionResult> GetColumnById(int id)
        {
            try
            {
                var column = await _entitySchemaService.GetColumnByIdAsync(id);
                if (column == null)
                {
                    return NotFound();
                }

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = column
                };
                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };
                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("columns/{id}/{entityid}")]
        public async Task<IActionResult> GetColumnByIdAndEntityId(int id, int entityid)
        {
            try
            {
                var column = await _entitySchemaService.GetColumnByIdAndEntityIDAsync(id, entityid);
                if (column == null)
                {
                    return NotFound();
                }

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = column
                };
                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };
                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("columns/entity/{entityId}")]
        public async Task<IActionResult> GetColumnsByEntityId(int entityId)
        {
            try
            {
                var columns = await _entitySchemaService.GetColumnsByEntityIdAsync(entityId);

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = columns
                };
                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };
                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpPost("createtables")]
        public async Task<IActionResult> InsertTable([FromQuery] DBConnectionDTO connectionDTO, [FromBody] TableRequest tableRequest)
        {
            try
            {
                var craetetable = await _entitySchemaService.convertandcallcreatetablemodel(connectionDTO, tableRequest);

                // Insert table
                var tableId = await _entitySchemaService.CreateTableAsync(tableRequest.Table);

                // Set the inserted table ID in each column
                foreach (var column in tableRequest.Columns)
                {
                    column.EntityId = tableId;
                }

                // Insert columns
                await _entitySchemaService.InsertColumnsAsync(tableRequest.Columns);

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.Created,
                    IsSuccess = true,
                    Result = new { TableId = tableId }
                };
                return CreatedAtAction(nameof(GetTableById), new
                {
                    id = tableId
                }, responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };
                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }


        [HttpGet("cliententity")]
        public async Task<IActionResult> ClientEntity([FromQuery] DBConnectionDTO connectionDTO)
        {
            try
            {
                string otherApiBaseUrl = "https://localhost:7246/";

                using (var httpClient = new HttpClient())
                {
                    httpClient.BaseAddress = new Uri(otherApiBaseUrl);

                    string requestUri = BuildRequestUri(connectionDTO);

                    // Log the request URI for debugging
                    Console.WriteLine($"Request URI: {requestUri}");

                    var response = await httpClient.GetAsync(requestUri);

                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = await response.Content.ReadAsStringAsync();

                        // Add detailed error handling
                        APIResponse tableDetails;
                        try
                        {
                            tableDetails = JsonConvert.DeserializeObject<APIResponse>(responseContent);
                        }
                        catch (JsonException jsonEx)
                        {
                            var responseModel = new APIResponse
                            {
                                StatusCode = HttpStatusCode.InternalServerError,
                                IsSuccess = false,
                                ErrorMessages = new List<string> { "Deserialization error: " + jsonEx.Message },
                                Result = null
                            };
                            return StatusCode((int)responseModel.StatusCode, responseModel);
                        }

                        if (!tableDetails.IsSuccess)
                        {
                            var responseModel = new APIResponse
                            {
                                StatusCode = HttpStatusCode.OK,
                                IsSuccess = true,
                                Result = tableDetails.Result
                            };
                            return Ok(responseModel);
                        }
                        else
                        {
                            var responseDetails = await _entitySchemaService.GetClientSchema(tableDetails, connectionDTO);

                            var responseModel = new APIResponse
                            {
                                StatusCode = HttpStatusCode.Created,
                                IsSuccess = true,
                                Result = responseDetails
                            };
                            return Ok(responseModel);
                        }
                    }
                    else
                    {
                        var responseModel = new APIResponse
                        {
                            StatusCode = response.StatusCode,
                            IsSuccess = false,
                            Result = null
                        };

                        return StatusCode((int)responseModel.StatusCode, responseModel);
                    }
                }
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };
                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }



        private string BuildRequestUri(DBConnectionDTO connectionDTO)
        {
            if (connectionDTO.Provider.Contains("Dynamo", StringComparison.OrdinalIgnoreCase))
            {
                return $"EntityMigrate/GetTableDetails?Provider={Uri.EscapeDataString(connectionDTO.Provider)}" +
                       $"&AccessKey={Uri.EscapeDataString(connectionDTO.AccessKey)}" +
                       $"&SecretKey={Uri.EscapeDataString(connectionDTO.SecretKey)}" +
                       $"&Region={Uri.EscapeDataString(connectionDTO.Region)}";
            }

            if (connectionDTO.Provider.Contains("Scylla", StringComparison.OrdinalIgnoreCase))
            {
                return $"EntityMigrate/GetTableDetails?Provider={Uri.EscapeDataString(connectionDTO.Provider)}" +
                       $"&IPAddress={Uri.EscapeDataString(connectionDTO.IPAddress)}" +
                       $"&PortNumber=9042" +
                       $"&Keyspace={Uri.EscapeDataString(connectionDTO.Keyspace)}" +
                       $"&Ec2Instance={Uri.EscapeDataString(connectionDTO.Ec2Instance)}";
            }



            else if (connectionDTO.Provider.Contains("Influx", StringComparison.OrdinalIgnoreCase))
            {
                //string decodedUrl = Uri.UnescapeDataString(connectionDTO.InfluxDbUrl);

                var uriBuilder = new UriBuilder("https://localhost:7246/EntityMigrate/GetTableDetails");
                var query = HttpUtility.ParseQueryString(uriBuilder.Query);

                query["Provider"] = connectionDTO.Provider;
                query["InfluxDbUrl"] = connectionDTO.InfluxDbUrl;  // No need to encode this part
                query["InfluxDbToken"] = connectionDTO.InfluxDbToken;  // No need to encode this part
                query["InfluxDbOrg"] = connectionDTO.InfluxDbOrg;  // No need to encode this part
                query["InfluxDbBucket"] = connectionDTO.InfluxDbBucket;  // No need to encode this part

                uriBuilder.Query = query.ToString();
                return uriBuilder.ToString();
                
            }



            else
            {
                string encodedPassword = Uri.EscapeDataString(connectionDTO.Password);
                return $"EntityMigrate/GetTableDetails?Provider={Uri.EscapeDataString(connectionDTO.Provider)}" +
                       $"&HostName={Uri.EscapeDataString(connectionDTO.HostName)}" +
                       $"&DataBase={Uri.EscapeDataString(connectionDTO.DataBase)}" +
                       $"&UserName={Uri.EscapeDataString(connectionDTO.UserName)}" +
                       $"&Password={encodedPassword}";
            }
        }


        [HttpPost("updatetables")]
        public async Task<IActionResult> UpdateTable([FromBody] List<ColumnDTO> columns)
        {
            try
            {
                await _entitySchemaService.UpdateColumnsAsync(columns);

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = null  // You can assign a meaningful result here if needed
                };

                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };

                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("columns/GetColumnsByHostProviderDatabaseTableName")]
        public async Task<IActionResult> GetColumnsByHostProviderDatabaseTableName(string? hostName, string provider, string? databaseName, string? tableName, string? accessKey, string ?secretKey, string? region, string? keyspace, string? ec2Instance, string? ipAddress, string? influxDbToken, string? influxDbOrg, string? influxDbUrl, string? influxDbBucket)
        {
            try
            {

                var table = await _entitySchemaService.GetColumnsByHostProviderDatabaseTableNameAsync(hostName, provider, databaseName, tableName, accessKey, secretKey, region, keyspace, ec2Instance, ipAddress,influxDbToken, influxDbOrg,influxDbUrl, influxDbBucket);
                if (table == null)
                {
                    return NotFound();
                }

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = table
                };

                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };

                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("log/{logParentId}")]
        public async Task<IActionResult> GetLogByParentId(int logParentId)
        {
            try
            {
                var logDTO = await _logService.GetLogByParentIdAsync(logParentId);

                if (logDTO == null)
                {
                    return NotFound(); // or handle the case where log with the specified ID is not found
                }

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = logDTO
                };

                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };

                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("log/child/{logChildId}")]
        public async Task<IActionResult> GetLogByChildId(int logChildId)
        {
            try
            {
                var logDTO = await _logService.GetLogByChildIdAsync(logChildId);

                if (logDTO == null)
                {
                    return NotFound(); // or handle the case where log with the specified child ID is not found
                }

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = logDTO
                };

                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };

                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("logs")]
        public async Task<IActionResult> GetAllLogs()
        {
            try
            {
                var logDTOs = await _logService.GetAllLogsAsync();

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = logDTOs
                };

                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };

                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("logs/user/{userId}")]
        public async Task<IActionResult> GetLogsByUserId(int userId)
        {
            try
            {
                var logDTOs = await _logService.GetLogsByUserIdAsync(userId);

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = logDTOs
                };

                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };

                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }

        [HttpGet("logs/entity/{entityId}")]
        public async Task<IActionResult> GetLogsByEntityId(int entityId)
        {
            try
            {
                var logDTOs = await _logService.GetLogsByEntityIdAsync(entityId);

                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.OK,
                    IsSuccess = true,
                    Result = logDTOs
                };

                return Ok(responseModel);
            }
            catch (Exception ex)
            {
                var responseModel = new APIResponse
                {
                    StatusCode = HttpStatusCode.InternalServerError,
                    IsSuccess = false,
                    ErrorMessages = new List<string> { ex.Message },
                    Result = null
                };

                return StatusCode((int)responseModel.StatusCode, responseModel);
            }
        }
    }
}
