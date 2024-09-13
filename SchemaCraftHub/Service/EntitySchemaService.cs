using Amazon;
using Amazon.DynamoDBv2.Model;
using Amazon.DynamoDBv2;
using ClientSchemaHub.Models.DTO;
using DBUtilityHub.Data;
using DBUtilityHub.Models;
using InfluxDB.Client;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using SchemaCraftHub.Model.DTO;
using SchemaCraftHub.Service.IService;
using System.Text;
using APIResponse = SchemaCraftHub.Model.DTO.APIResponse;
using DBConnectionDTO = SchemaCraftHub.Model.DTO.DBConnectionDTO;
using Cassandra;



namespace SchemaCraftHub.Service
{
    public class EntitySchemaService : IEntitySchemaService
    {
        private readonly ApplicationDbContext _context;

        public EntitySchemaService(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<List<TableMetaDataDTO>> GetAllTablesAsync()
        {
            try
            {
                var tables = await _context.TableMetaDataEntity.ToListAsync();

                var tableDTOs = tables.Select(table => new TableMetaDataDTO
                {
                    Id = table.Id,
                    EntityName = table.EntityName,
                    HostName = table.HostName,
                    DatabaseName = table.DatabaseName,
                    Provider = table.Provider,
                    Ec2Instance = table.Ec2Instance
                    // Map other properties as needed
                }).ToList();

                return tableDTOs;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred while fetching all tables.", ex);
            }
        }

        public async Task<TableMetaDataDTO> GetTableByIdAsync(int id)
        {
            try
            {
                var table = await _context.TableMetaDataEntity
                    .Where(t => t.Id == id)
                    .Select(table => new TableMetaDataDTO
                    {
                        Id = table.Id,
                        EntityName = table.EntityName,
                        HostName = table.HostName,
                        DatabaseName = table.DatabaseName,
                        Provider = table.Provider
                        // Map other properties as needed
                    })
                    .FirstOrDefaultAsync();

                return table;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred while fetching all tables.", ex);
            }
        }



        public async Task<List<TableMetaDataDTO>> GetTablesByHostProviderDatabaseAsync(
    string? hostName, string provider, string? databaseName,
    string? accessKey, string? secretKey, string? region,
    string? keyspace, string? ec2Instance, string? ipAddress, string? influxDbToken, string? influxDbOrg, string? influxDbUrl, string? influxDbBucket)
        {
            try
            {
                List<TableMetaDataDTO> tableDTOs = new List<TableMetaDataDTO>();

                if (provider.Equals("Dynamo", StringComparison.OrdinalIgnoreCase))
                {
                    //Console.WriteLine("Fetching tables from DynamoDB.");

                    //// Set up DynamoDB client
                    //var client = new AmazonDynamoDBClient(accessKey, secretKey, RegionEndpoint.GetBySystemName(region));

                    //// Fetch tables from DynamoDB
                    //var request = new ListTablesRequest();
                    //var response = await client.ListTablesAsync(request);

                    //foreach (var tableName in response.TableNames)
                    //{
                    //    tableDTOs.Add(new TableMetaDataDTO
                    //    {
                    //        EntityName = tableName,
                    //        DatabaseName = databaseName, // Use the provided database name dynamically
                    //        Provider = provider
                    //    });
                    //}


                    var tables = await _context.TableMetaDataEntity
                        .Where(table => table.AccessKey.ToLower() == accessKey.ToLower() &&
                                        table.SecretKey.ToLower() == secretKey.ToLower() &&
                                        table.Region.ToLower() == region.ToLower())
                        .ToListAsync();

                    tableDTOs = tables.Select(table => new TableMetaDataDTO
                    {
                        Id = table.Id,
                        EntityName = table.EntityName,
                        AccessKey = table.AccessKey,
                        SecretKey = table.SecretKey,
                        Region = table.Region
                        // Map other properties as needed
                    }).ToList();



                }
                else if (provider.Equals("Scylla", StringComparison.OrdinalIgnoreCase))
                {
                    //Console.WriteLine("Fetching tables from ScyllaDB.");

                    //// Set up ScyllaDB client using Cassandra driver
                    //var cluster = Cluster.Builder()
                    //    .AddContactPoint(ipAddress) // Use the IP address
                    //    .WithPort(9042)              // Use port 9042
                    //    .Build();

                    //using (var session = cluster.Connect(keyspace)) // Use the keyspace
                    //{
                    //    // Query system schema to find all table names in the keyspace
                    //    var query = $"SELECT table_name FROM system_schema.tables WHERE keyspace_name = '{keyspace}'";
                    //    var resultSet = session.Execute(query);

                    //    foreach (var row in resultSet)
                    //    {
                    //        tableDTOs.Add(new TableMetaDataDTO
                    //        {
                    //            EntityName = row.GetValue<string>("table_name"),
                    //            DatabaseName = keyspace,
                    //            Provider = provider,
                    //            HostName = hostName
                    //        });
                    //    }
                    //}


                    var tables = await _context.TableMetaDataEntity
                        .Where(table => table.Ec2Instance.ToLower() == ec2Instance.ToLower() &&
                                        table.IPAddress.ToLower() == ipAddress.ToLower() &&
                                        table.Keyspace.ToLower() == keyspace.ToLower())
                        .ToListAsync();

                    tableDTOs = tables.Select(table => new TableMetaDataDTO
                    {
                        Id = table.Id,
                        EntityName = table.EntityName,
                        Ec2Instance = table.Ec2Instance,
                        IPAddress = table.IPAddress,
                        Keyspace = table.Keyspace
                        // Map other properties as needed
                    }).ToList();




                }
                else if (provider.Equals("Influx", StringComparison.OrdinalIgnoreCase))
                {
                    var tables = await _context.TableMetaDataEntity
                        .Where(table => table.InfluxDbBucket.ToLower() == influxDbBucket.ToLower() &&
                                        table.InfluxDbOrg.ToLower() == influxDbOrg.ToLower() &&
                                        table.InfluxDbToken.ToLower() == influxDbToken.ToLower())
                        .ToListAsync();

                    tableDTOs = tables.Select(table => new TableMetaDataDTO
                    {
                        Id = table.Id,
                        EntityName = table.EntityName,
                        InfluxDbBucket = table.InfluxDbBucket,
                        InfluxDbToken = table.InfluxDbToken,
                        InfluxDbOrg = table.InfluxDbOrg
                        // Map other properties as needed
                    }).ToList();
                }
                else
                {
                    var tables = await _context.TableMetaDataEntity
                        .Where(table => table.HostName.ToLower() == hostName.ToLower() &&
                                        table.Provider.ToLower() == provider.ToLower() &&
                                        table.DatabaseName.ToLower() == databaseName.ToLower())
                        .ToListAsync();

                    tableDTOs = tables.Select(table => new TableMetaDataDTO
                    {
                        Id = table.Id,
                        EntityName = table.EntityName,
                        HostName = table.HostName,
                        DatabaseName = table.DatabaseName,
                        Provider = table.Provider
                        // Map other properties as needed
                    }).ToList();
                }

                return tableDTOs;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");

                throw new ApplicationException("An error occurred while fetching tables.", ex);
            }
        }


        public async Task<TableMetaDataDTO> GetTableByHostProviderDatabaseTableNameAsync(
    string? hostName, string provider, string? databaseName,
    string? accessKey, string? secretKey, string? region,
    string? keyspace, string? ec2Instance, string? ipAddress,
    string? tableName, string? influxDbToken, string? influxDbOrg, string? influxDbUrl, string? influxDbBucket)
        {
            try
            {
                TableMetaDataDTO tableDTO = null;

                if (provider.Equals("Dynamo", StringComparison.OrdinalIgnoreCase))
                {
                    //Console.WriteLine("Fetching tables from DynamoDB.");

                    //// Set up DynamoDB client
                    //var client = new AmazonDynamoDBClient(accessKey, secretKey, RegionEndpoint.GetBySystemName(region));

                    //// Fetch tables from DynamoDB
                    //var request = new ListTablesRequest();
                    //var response = await client.ListTablesAsync(request);

                    var table = await _context.TableMetaDataEntity
                        .FirstOrDefaultAsync(t => t.AccessKey.ToLower() == accessKey.ToLower() &&
                                                  t.SecretKey.ToLower() == secretKey.ToLower() &&
                                                  t.Region.ToLower() == region.ToLower() &&
                                                  t.Provider.ToLower() == provider.ToLower() &&
                                                  t.DatabaseName.ToLower() == databaseName.ToLower() &&
                                                  t.EntityName.ToLower() == tableName.ToLower());

                    if (table == null)
                    {
                        return null;
                    }    
                    
                            tableDTO = new TableMetaDataDTO
                            {
                                Id = table.Id,
                                EntityName = tableName,
                                DatabaseName = databaseName,
                                Provider = provider,
                                HostName = hostName,
                                AccessKey = accessKey,
                                Region = region,
                                SecretKey = secretKey
                            };                      
                   
                    return tableDTO;
                }
                else if (provider.Equals("Scylla", StringComparison.OrdinalIgnoreCase))
                {
                    //Console.WriteLine("Fetching tables from ScyllaDB.");

                    //var cluster = Cluster.Builder()
                    //    .AddContactPoint(ipAddress) // This should be the IP address
                    //    .WithPort(9042)              // Hardcode the port as it is always 9042
                    //    .Build();

                    //using (var session = cluster.Connect(keyspace)) // Keyspace is correct
                    //{
                    //    var query = $"SELECT table_name FROM system_schema.tables WHERE keyspace_name = '{keyspace}' AND table_name = '{tableName}'";
                    //    var resultSet = session.Execute(query);
                    //    var row = resultSet.FirstOrDefault();

                    //    if (row != null)
                    //    {
                            var table = await _context.TableMetaDataEntity
                                .FirstOrDefaultAsync(t => t.IPAddress.ToLower() == ipAddress.ToLower() &&
                                                          t.Keyspace.ToLower() == keyspace.ToLower() &&
                                                          t.Provider.ToLower() == provider.ToLower() &&
                                                          t.Ec2Instance.ToLower() == ec2Instance.ToLower() &&
                                                          t.EntityName.ToLower() == tableName.ToLower());

                            if (table == null)
                            {
                                return null;
                            }

                            tableDTO = new TableMetaDataDTO
                            {
                                Id = table.Id,
                               // EntityName = row.GetValue<string>("table_name"),
                               EntityName = tableName,
                                Provider = provider,
                                IPAddress = ipAddress,
                                Keyspace = keyspace,
                                Ec2Instance = ec2Instance
                            };

                            return tableDTO;
                        //}
                    //}

                 //   return null;
                }

                if (provider.Equals("Influx", StringComparison.OrdinalIgnoreCase))
                {
                    //Console.WriteLine("Fetching measurements from InfluxDB.");

                    //// Set up InfluxDB client
                    //var options = new InfluxDBClientOptions.Builder()
                    //    .Url(influxDbUrl)
                    //    .AuthenticateToken(influxDbToken.ToCharArray())
                    //    .Org(influxDbOrg)
                    //    .Build();

                    //var influxDbClient = new InfluxDBClient(options);
                    //var tableDetails = new TableDetailsDTO { TableName = tableName };        
                    
                    //    var query = $"from(bucket: \"{influxDbBucket}\") |> range(start: -1h) |> limit(n:1)";
                    //    var fluxTables = await influxDbClient.GetQueryApi().QueryAsync(query,influxDbOrg);

                    //    if (fluxTables.Count > 0)
                    //    {
                    //        var fluxTable = fluxTables[0];
                    //        tableDetails.Columns = fluxTable.Columns.Select(col => new ColumnDetailsDTO
                    //        {
                    //            ColumnName = col.Label,
                    //            DataType = col.DataType
                    //        }).ToList();
                    //    }      

                        var table = await _context.TableMetaDataEntity
                            .FirstOrDefaultAsync(table => table.InfluxDbToken.ToLower() == influxDbToken.ToLower() &&
                                             table.InfluxDbOrg.ToLower() == influxDbOrg.ToLower() &&
                                             table.EntityName.ToLower() == tableName.ToLower() &&
                                             table.InfluxDbBucket.ToLower() == influxDbBucket.ToLower());

                        if (table == null)
                        {
                            return null;
                        }

                        tableDTO = new TableMetaDataDTO
                        {
                            Id = table.Id,
                            EntityName = tableName,
                            DatabaseName = databaseName,
                            Provider = provider,
                            InfluxDbOrg = influxDbOrg,
                            InfluxDbToken = influxDbToken,
                            InfluxDbUrl = influxDbUrl,
                            InfluxDbBucket = influxDbBucket
                        };

                        return tableDTO;
                    
                }

                else
                {
                    var table = await _context.TableMetaDataEntity
                        .FirstOrDefaultAsync(t => t.HostName.ToLower() == hostName.ToLower() &&
                                                  t.Provider.ToLower() == provider.ToLower() &&
                                                  t.DatabaseName.ToLower() == databaseName.ToLower() &&
                                                  t.EntityName.ToLower() == tableName.ToLower());

                    if (table == null)
                    {
                        return null;
                    }

                    tableDTO = new TableMetaDataDTO
                    {
                        Id = table.Id,
                        EntityName = table.EntityName,
                        HostName = table.HostName,
                        DatabaseName = table.DatabaseName,
                        Provider = table.Provider
                    };
                    return tableDTO;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");

                throw new ApplicationException("An error occurred while fetching the table.", ex);
            }
        }

        public async Task<List<ColumnDTO>> GetAllColumnsAsync()
        {
            try
            {
                var columns = await _context.ColumnMetaDataEntity.ToListAsync();

                var columnDTOs = columns.Select(column => new Model.DTO.ColumnDTO
                {
                    Id = column.Id,
                    ColumnName = column.ColumnName,
                    Datatype = column.Datatype,
                    IsPrimaryKey = column.IsPrimaryKey,
                    IsForeignKey = column.IsForeignKey,
                    EntityId = column.EntityId,
                    ReferenceEntityID = column.ReferenceEntityID,
                    ReferenceColumnID = column.ReferenceColumnID,
                    Length = column.Length,
                    MinLength = column.MinLength,
                    MaxLength = column.MaxLength,
                    MaxRange = column.MaxRange,
                    MinRange = column.MinRange,
                    DateMinValue = column.DateMinValue,
                    DateMaxValue = column.DateMaxValue,
                    Description = column.Description,
                    IsNullable = column.IsNullable,
                    DefaultValue = column.DefaultValue,
                    True = column.True,
                    False = column.False,
                    // Include other properties as needed
                }).ToList();


                return columnDTOs;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred while fetching all tables.", ex);
            }
        }

        public async Task<ColumnDTO> GetColumnByIdAsync(int id)
        {
            try
            {
                var column = await _context.ColumnMetaDataEntity.FirstOrDefaultAsync(x => x.Id == id);

                if (column == null)
                {
                    // Handle the case where the column with the given ID is not found
                    return null;
                }

                var columnDTO = new Model.DTO.ColumnDTO
                {
                    Id = column.Id,
                    ColumnName = column.ColumnName,
                    Datatype = column.Datatype,
                    IsPrimaryKey = column.IsPrimaryKey,
                    IsForeignKey = column.IsForeignKey,
                    EntityId = column.EntityId,
                    ReferenceEntityID = column.ReferenceEntityID,
                    ReferenceColumnID = column.ReferenceColumnID,
                    Length = column.Length,
                    MinLength = column.MinLength,
                    MaxLength = column.MaxLength,
                    MaxRange = column.MaxRange,
                    MinRange = column.MinRange,
                    DateMinValue = column.DateMinValue,
                    DateMaxValue = column.DateMaxValue,
                    Description = column.Description,
                    IsNullable = column.IsNullable,
                    DefaultValue = column.DefaultValue,
                    True = column.True,
                    False = column.False,
                };

                return columnDTO;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred while fetching all tables.", ex);
            }
        }

        public async Task<ColumnDTO> GetColumnByIdAndEntityIDAsync(int id, int entityId)
        {
            try
            {
                var column = await _context.ColumnMetaDataEntity.FirstOrDefaultAsync(x => x.Id == id && x.EntityId == entityId);

                if (column == null)
                {
                    // Handle the case where the column with the given ID is not found
                    return null;
                }

                var columnDTO = new Model.DTO.ColumnDTO
                {
                    Id = column.Id,
                    ColumnName = column.ColumnName,
                    Datatype = column.Datatype,
                    IsPrimaryKey = column.IsPrimaryKey,
                    IsForeignKey = column.IsForeignKey,
                    EntityId = column.EntityId,
                    ReferenceEntityID = column.ReferenceEntityID,
                    ReferenceColumnID = column.ReferenceColumnID,
                    Length = column.Length,
                    MinLength = column.MinLength,
                    MaxLength = column.MaxLength,
                    MaxRange = column.MaxRange,
                    MinRange = column.MinRange,
                    DateMinValue = column.DateMinValue,
                    DateMaxValue = column.DateMaxValue,
                    Description = column.Description,
                    IsNullable = column.IsNullable,
                    DefaultValue = column.DefaultValue,
                    True = column.True,
                    False = column.False,
                };

                return columnDTO;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred while fetching all tables.", ex);
            }
        }

        public async Task<List<ColumnDTO>> GetColumnsByEntityIdAsync(int entityId)
        {
            try
            {
                var columns = await _context.ColumnMetaDataEntity
                    .Where(column => column.EntityId == entityId)
                    .ToListAsync();

                var columnDTOs = columns.Select(column => new Model.DTO.ColumnDTO
                {
                    Id = column.Id,
                    ColumnName = column.ColumnName,
                    Datatype = column.Datatype,
                    IsPrimaryKey = column.IsPrimaryKey,
                    IsForeignKey = column.IsForeignKey,
                    EntityId = column.EntityId,
                    ReferenceEntityID = column.ReferenceEntityID,
                    ReferenceColumnID = column.ReferenceColumnID,
                    Length = column.Length,
                    MinLength = column.MinLength,
                    MaxLength = column.MaxLength,
                    MaxRange = column.MaxRange,
                    MinRange = column.MinRange,
                    DateMinValue = column.DateMinValue,
                    DateMaxValue = column.DateMaxValue,
                    Description = column.Description,
                    IsNullable = column.IsNullable,
                    DefaultValue = column.DefaultValue,
                    True = column.True,
                    False = column.False,
                    // Include other properties as needed
                }).ToList();

                return columnDTOs;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred while fetching all tables.", ex);
            }
        }

        public async Task<Dictionary<string, List<ClientSchemaHub.Models.DTO.TableDetailsDTO>>> GetClientSchema(APIResponse tabledetails1,DBConnectionDTO connectionDTO)
        {
            try
            {
                if (tabledetails1.Result != null)
                {
                    try
                    {
                        var jsonString = JsonConvert.SerializeObject(tabledetails1.Result);
                        var tableDetailsDict = JsonConvert.DeserializeObject<Dictionary<string, List<ClientSchemaHub.Models.DTO.TableDetailsDTO>>>(jsonString);

                        foreach (var tabledetailsDTO in tableDetailsDict)
                        {
                            foreach (var table in tabledetailsDTO.Value)
                            {
                                var table_exists = await GetTableByHostProviderDatabaseTableNameAsync(connectionDTO.HostName, connectionDTO.Provider, connectionDTO.DataBase, connectionDTO.AccessKey, connectionDTO.SecretKey,  connectionDTO.Region, connectionDTO.Keyspace, connectionDTO.Ec2Instance, connectionDTO.IPAddress, table.TableName, connectionDTO.InfluxDbToken, connectionDTO.InfluxDbOrg, connectionDTO.InfluxDbUrl, connectionDTO.InfluxDbBucket);
                                if (table_exists == null)
                                {
                                    var tablename = new TableMetaDataDTO
                                    {
                                        EntityName = table.TableName,
                                        HostName = connectionDTO.HostName,
                                        DatabaseName = connectionDTO.DataBase,
                                        Provider = connectionDTO.Provider,
                                        AccessKey = connectionDTO.AccessKey,
                                        Region = connectionDTO.Region,
                                        SecretKey = connectionDTO.SecretKey,
                                        Keyspace = connectionDTO.Keyspace,
                                        Ec2Instance = connectionDTO.Ec2Instance,
                                        IPAddress = connectionDTO.IPAddress,
                                        InfluxDbOrg = connectionDTO.InfluxDbOrg,
                                        InfluxDbToken = connectionDTO.InfluxDbToken,
                                        InfluxDbUrl = connectionDTO.InfluxDbUrl,
                                        InfluxDbBucket = connectionDTO.InfluxDbBucket
                                    };
                                    await CreateTableAsync(tablename);
                                }
                            }
                        }

                        foreach (var tabledetailsDTO in tableDetailsDict)
                        {
                            foreach (var table in tabledetailsDTO.Value)
                            {
                                var table_exists = await GetTableByHostProviderDatabaseTableNameAsync(connectionDTO.HostName, connectionDTO.Provider, connectionDTO.DataBase, connectionDTO.AccessKey, connectionDTO.SecretKey, connectionDTO.Region, connectionDTO.Keyspace,  connectionDTO.Ec2Instance, connectionDTO.IPAddress, table.TableName, connectionDTO.InfluxDbToken, connectionDTO.InfluxDbOrg, connectionDTO.InfluxDbUrl, connectionDTO.InfluxDbBucket);

                                if (table_exists != null)
                                {
                                    var insertcolumnEntities = new List<ColumnDTO>();

                                    var updatecolumnEntities = new List<ColumnDTO>();

                                    foreach (var columnDTO in table.Columns)
                                    {
                                        var columns = await GetColumnsByEntityIdAsync(table_exists.Id);

                                        var column_exists = columns.FirstOrDefault(x => x.ColumnName.ToLower() == columnDTO.ColumnName.ToLower());

                                        if (column_exists == null)
                                        {
                                            var EntityId = (await GetTableByHostProviderDatabaseTableNameAsync(connectionDTO.HostName, connectionDTO.Provider, connectionDTO.DataBase, connectionDTO.AccessKey, connectionDTO.SecretKey, connectionDTO.Region, connectionDTO.Keyspace , connectionDTO.Ec2Instance, connectionDTO.IPAddress, table.TableName, connectionDTO.InfluxDbToken, connectionDTO.InfluxDbOrg, connectionDTO.InfluxDbUrl,connectionDTO.InfluxDbBucket)).Id;
                                            Nullable<int> ReferenceEntityID = null;
                                            if (columnDTO.HasForeignKey)
                                            {
                                                ReferenceEntityID = (await GetTableByHostProviderDatabaseTableNameAsync(connectionDTO.HostName, connectionDTO.Provider, connectionDTO.DataBase, connectionDTO.AccessKey, connectionDTO.SecretKey, connectionDTO.Region, connectionDTO.Keyspace, connectionDTO.Ec2Instance, connectionDTO.IPAddress, columnDTO.ReferencedTable, connectionDTO.InfluxDbToken, connectionDTO.InfluxDbOrg, connectionDTO.InfluxDbUrl, connectionDTO.InfluxDbBucket)).Id;
                                            }
                                            var columnEntity = new ColumnDTO
                                            {
                                                ColumnName = columnDTO.ColumnName,
                                                Datatype = columnDTO.DataType,
                                                IsPrimaryKey = columnDTO.IsPrimaryKey,
                                                IsForeignKey = columnDTO.HasForeignKey,
                                                EntityId = EntityId,
                                                ReferenceEntityID = ReferenceEntityID,
                                                IsNullable = columnDTO.IsNullable,
                                                // Map other properties from ColumnDetailsDTO as needed
                                            };

                                            insertcolumnEntities.Add(columnEntity);
                                        }
                                        else
                                        {
                                            var EntityId = (await GetTableByHostProviderDatabaseTableNameAsync(connectionDTO.HostName, connectionDTO.Provider, connectionDTO.DataBase, connectionDTO.AccessKey, connectionDTO.SecretKey, connectionDTO.Region, connectionDTO.Keyspace, connectionDTO.Ec2Instance, connectionDTO.IPAddress, table.TableName, connectionDTO.InfluxDbToken, connectionDTO.InfluxDbOrg, connectionDTO.InfluxDbUrl, connectionDTO.InfluxDbBucket)).Id;
                                            Nullable<int> ReferenceEntityID = null;
                                            if (columnDTO.HasForeignKey)
                                            {
                                                ReferenceEntityID = (await GetTableByHostProviderDatabaseTableNameAsync(connectionDTO.HostName, connectionDTO.Provider, connectionDTO.DataBase, connectionDTO.AccessKey, connectionDTO.SecretKey, connectionDTO.Region, connectionDTO.Keyspace, connectionDTO.Ec2Instance, connectionDTO.IPAddress, columnDTO.ReferencedTable, connectionDTO.InfluxDbToken, connectionDTO.InfluxDbOrg, connectionDTO.InfluxDbUrl, connectionDTO.InfluxDbBucket)).Id;
                                            }
                                            var columnEntity = new ColumnDTO
                                            {
                                                Id = column_exists.Id,
                                                ColumnName = columnDTO.ColumnName,
                                                Datatype = columnDTO.DataType,
                                                IsPrimaryKey = columnDTO.IsPrimaryKey,
                                                IsForeignKey = columnDTO.HasForeignKey,
                                                EntityId = EntityId,
                                                ReferenceEntityID = ReferenceEntityID,
                                                IsNullable = columnDTO.IsNullable,
                                                // Map other properties from ColumnDetailsDTO as needed
                                            };

                                            updatecolumnEntities.Add(columnEntity);
                                        }
                                    }
                                    if (insertcolumnEntities.Count > 0)
                                    {
                                        await InsertColumnsAsync(insertcolumnEntities);
                                    }
                                    if (updatecolumnEntities.Count > 0)
                                    {
                                        await UpdateColumnsAsync(updatecolumnEntities);
                                    }
                                }
                            }
                        }

                        return tableDetailsDict;
                    }
                    catch (JsonException)
                    {
                        // Handle the case where JSON deserialization fails
                        throw new ApplicationException("Failed to deserialize Result into the expected format");
                    }
                }
                else
                {
                    // Handle the case where Result is null
                    throw new ApplicationException("Result is null");
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message, ex);
            }
        }

        public async Task<int> CreateTableAsync(TableMetaDataDTO tableDTO)
        {
            try
            {
                var tableexists = await GetTableByHostProviderDatabaseTableNameAsync(tableDTO.HostName, tableDTO.Provider, tableDTO.EntityName, tableDTO.AccessKey, tableDTO.SecretKey, tableDTO.Region, tableDTO.Keyspace, tableDTO.Ec2Instance, tableDTO.IPAddress, tableDTO.DatabaseName, tableDTO.InfluxDbToken, tableDTO.InfluxDbOrg, tableDTO.InfluxDbUrl, tableDTO.InfluxDbBucket);
                if (tableexists == null)
                {
                    var table = new TableMetaDataEntity
                    {
                        EntityName = tableDTO.EntityName,
                        HostName = tableDTO.HostName,
                        DatabaseName = tableDTO.DatabaseName,
                        Provider = tableDTO.Provider,
                        AccessKey = tableDTO.AccessKey,
                        Region = tableDTO.Region,
                        SecretKey = tableDTO.SecretKey,
                        Keyspace = tableDTO.Keyspace,
                        IPAddress  = tableDTO.IPAddress,
                        Ec2Instance = tableDTO.Ec2Instance,
                        InfluxDbUrl = tableDTO.InfluxDbUrl,
                        InfluxDbToken = tableDTO.InfluxDbToken,
                        InfluxDbOrg = tableDTO.InfluxDbOrg,
                        InfluxDbBucket = tableDTO.InfluxDbBucket
                        // Map other properties as needed
                    };

                    _context.TableMetaDataEntity.Add(table);
                    await _context.SaveChangesAsync();
                    return table.Id;
                }
                return tableexists.Id;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred while fetching all tables.", ex);
            }
        }

        public async Task InsertColumnsAsync(List<ColumnDTO> columns)
        {
            try
            {
                var columnEntities = columns.Select(columnDTO => new ColumnMetaDataEntity
                {
                    Id = columnDTO.Id,
                    ColumnName = columnDTO.ColumnName,
                    Datatype = columnDTO.Datatype,
                    IsPrimaryKey = columnDTO.IsPrimaryKey,
                    IsForeignKey = columnDTO.IsForeignKey,
                    EntityId = columnDTO.EntityId,
                    ReferenceEntityID = columnDTO.ReferenceEntityID,
                    ReferenceColumnID = columnDTO.ReferenceColumnID,
                    Length = columnDTO.Length,
                    MinLength = columnDTO.MinLength,
                    MaxLength = columnDTO.MaxLength,
                    MaxRange = columnDTO.MaxRange,
                    MinRange = columnDTO.MinRange,
                    DateMinValue = columnDTO.DateMinValue,
                    DateMaxValue = columnDTO.DateMaxValue,
                    Description = columnDTO.Description,
                    IsNullable = columnDTO.IsNullable,
                    DefaultValue = columnDTO.DefaultValue,
                    True = columnDTO.True,
                    False = columnDTO.False
                    // Map other properties as needed
                }).ToList();

                _context.ColumnMetaDataEntity.AddRange(columnEntities);
                await _context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred while fetching all tables.", ex);
            }
        }

        public async Task<APIResponse> convertandcallcreatetablemodel(DBConnectionDTO connectionDTO, TableRequest tableRequest)
        {
            ClientSchemaHub.Models.DTO.TableDetailsDTO maptable = new ClientSchemaHub.Models.DTO.TableDetailsDTO
            {
                TableName = tableRequest.Table.EntityName,
                Columns = await MapColumns(tableRequest.Columns)
            };

            var createquery = GenerateCreateTableSql(maptable);

            string otherApiBaseUrl = "https://localhost:7246";

            // Create HttpClient instance
            using (var httpClient = new HttpClient())
            {
                // Set the base address of the other API
                httpClient.BaseAddress = new Uri(otherApiBaseUrl);

                string encodedPassword = Uri.EscapeDataString(connectionDTO.Password);

                // Call the other API to get table details
                var response = await httpClient.GetAsync($"/EntityMigrate/CreateTable?Provider={connectionDTO.Provider}&HostName={connectionDTO.HostName}&DataBase={connectionDTO.DataBase}&UserName={connectionDTO.UserName}&Password={encodedPassword}&query={createquery}");

                // Check if the request was successful
                if (response.IsSuccessStatusCode)
                {
                    // Read the response content
                    var responseContent = await response.Content.ReadAsStringAsync();

                    // Parse the response content as needed (assuming it's JSON)
                    var tableDetails = JsonConvert.DeserializeObject<APIResponse>(responseContent);

                    // Continue with your logic...

                    return tableDetails;
                }
                else
                {
                    // Handle unsuccessful response
                    var responseModel = new APIResponse
                    {
                        StatusCode = response.StatusCode,
                        IsSuccess = false,
                        Result = null // You might want to include more details here
                    };

                    return responseModel;
                }
            }

        }

        private async Task<List<ColumnDetailsDTO>> MapColumns(List<ColumnDTO> columns)
        {
            var columnDetailsList = new List<ColumnDetailsDTO>();

            foreach (var column in columns)
            {
                var referencedTable = await GetTableByIdAsync(column.ReferenceEntityID ?? 0);
                var referencedColumn = await GetColumnByIdAsync(column.ReferenceColumnID ?? 0);

                var columnDetails = new ColumnDetailsDTO
                {
                    ColumnName = column.ColumnName,
                    DataType = column.Datatype,
                    IsPrimaryKey = column.IsPrimaryKey,
                    IsNullable = column.IsNullable,
                    HasForeignKey = column.IsForeignKey,
                    ReferencedTable = referencedTable?.EntityName,
                    ReferencedColumn = referencedColumn?.ColumnName
                    // Add other properties as needed
                };

                columnDetailsList.Add(columnDetails);
            }

            return columnDetailsList;
        }

        private string GenerateCreateTableSql(ClientSchemaHub.Models.DTO.TableDetailsDTO mapTable)
        {
            StringBuilder sqlBuilder = new StringBuilder();

            // Basic SQL statement to create a table
            sqlBuilder.Append($"CREATE TABLE {mapTable.TableName} (");

            // Generate columns
            foreach (var column in mapTable.Columns)
            {
                sqlBuilder.Append($"{column.ColumnName} {column.DataType}");

                if (column.IsPrimaryKey)
                {
                    sqlBuilder.Append(" PRIMARY KEY");
                }

                if (column.HasForeignKey)
                {
                    // Assuming that ReferencedTable and ReferencedColumn are foreign key references
                    sqlBuilder.Append($" REFERENCES {column.ReferencedTable}({column.ReferencedColumn})");
                }

                // Add other column properties as needed
                sqlBuilder.Append(",");
            }

            // Remove the trailing comma
            sqlBuilder.Length--;

            // Close the CREATE TABLE statement
            sqlBuilder.Append(");");

            return sqlBuilder.ToString();
        }

        public async Task UpdateColumnsAsync(List<ColumnDTO> columns)
        {
            try
            {
                foreach (var columnDTO in columns)
                {
                    var existingColumnEntity = await _context.ColumnMetaDataEntity.FirstOrDefaultAsync(x => x.Id == columnDTO.Id);

                    if (existingColumnEntity != null)
                    {
                        // Update the properties of the existing entity
                        existingColumnEntity.ColumnName = string.IsNullOrEmpty(columnDTO.ColumnName) ? existingColumnEntity.ColumnName : columnDTO.ColumnName;
                        existingColumnEntity.Datatype = string.IsNullOrEmpty(columnDTO.Datatype) ? existingColumnEntity.Datatype : columnDTO.Datatype;
                        existingColumnEntity.IsPrimaryKey = columnDTO.IsPrimaryKey;
                        existingColumnEntity.IsForeignKey = columnDTO.IsForeignKey;
                        existingColumnEntity.EntityId = columnDTO.EntityId;
                        existingColumnEntity.ReferenceEntityID = columnDTO.ReferenceEntityID;
                        existingColumnEntity.ReferenceColumnID = columnDTO.ReferenceColumnID;
                        existingColumnEntity.Length = columnDTO.Length ?? existingColumnEntity.Length;
                        existingColumnEntity.MinLength = columnDTO.MinLength ?? existingColumnEntity.MinLength;
                        existingColumnEntity.MaxLength = columnDTO.MaxLength ?? existingColumnEntity.MaxLength;
                        existingColumnEntity.MaxRange = columnDTO.MaxRange ?? existingColumnEntity.MaxRange;
                        existingColumnEntity.MinRange = columnDTO.MinRange ?? existingColumnEntity.MinRange;
                        existingColumnEntity.DateMinValue = string.IsNullOrEmpty(columnDTO.DateMinValue) ? existingColumnEntity.DateMinValue : columnDTO.DateMinValue;
                        existingColumnEntity.DateMaxValue = string.IsNullOrEmpty(columnDTO.DateMaxValue) ? existingColumnEntity.DateMaxValue : columnDTO.DateMaxValue;
                        existingColumnEntity.Description = string.IsNullOrEmpty(columnDTO.Description) ? existingColumnEntity.Description : columnDTO.Description;
                        existingColumnEntity.IsNullable = columnDTO.IsNullable;
                        existingColumnEntity.DefaultValue = string.IsNullOrEmpty(columnDTO.DefaultValue) ? existingColumnEntity.DefaultValue : columnDTO.DefaultValue;
                        existingColumnEntity.True = string.IsNullOrEmpty(columnDTO.True) ? existingColumnEntity.True : columnDTO.True;
                        existingColumnEntity.False = string.IsNullOrEmpty(columnDTO.False) ? existingColumnEntity.False : columnDTO.False;

                        // Update other properties as needed

                        // Mark the entity as modified
                        _context.ColumnMetaDataEntity.Update(existingColumnEntity);
                    }
                    else
                    {
                        throw new ApplicationException("Column not found");
                    }
                }

                await _context.SaveChangesAsync(); // Use await to ensure the asynchronous operation is completed
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred while updating columns.", ex);
            }
        }
        public async Task<List<ColumnDTO>> GetColumnsByHostProviderDatabaseTableNameAsync(string? hostName, string provider, string? databaseName, string? tableName, string? accessKey, string? secretKey, string? region, string? keyspace, string? ec2Instance, string? ipAddress, string? influxDbToken, string? influxDbOrg, string? influxDbUrl, string? influxDbBucket)
        {
            try
            {
                List<ColumnDTO> columns = new List<ColumnDTO>();

                if (provider.Equals("Dynamo", StringComparison.OrdinalIgnoreCase))
                {
                    //Console.WriteLine("Fetching columns from DynamoDB table.");
                    //// Set up DynamoDB client
                    //var client = new AmazonDynamoDBClient(accessKey, secretKey, RegionEndpoint.GetBySystemName(region));
                    //// Describe the DynamoDB table
                    //var describeTableRequest = new DescribeTableRequest
                    //{
                    //    TableName = tableName
                    //};

                    //var describeTableResponse = await client.DescribeTableAsync(describeTableRequest);
                    //// Extract attribute definitions (columns) from the table description
                    //var attributeDefinitions = describeTableResponse.Table.AttributeDefinitions;
                    //var keySchema = describeTableResponse.Table.KeySchema;
                    //foreach (var attribute in attributeDefinitions)
                    //{
                    //    bool isPrimaryKey = keySchema.Any(k => k.AttributeName == attribute.AttributeName);
                    //    var columnDTO = new ColumnDTO
                    //    {
                    //        Id = 0, // Default to 0, should be set appropriately based on your requirements
                    //        ColumnName = attribute.AttributeName,
                    //        Datatype = attribute.AttributeType.ToString(),
                    //        IsPrimaryKey = isPrimaryKey,
                    //        IsForeignKey = false,
                    //        EntityId = 0,
                    //        ReferenceEntityID = null,
                    //        ReferenceColumnID = null,
                    //        Length = null,
                    //        MinLength = null,
                    //        MaxLength = null,
                    //        MaxRange = null,
                    //        MinRange = null,
                    //        DateMinValue = null,
                    //        DateMaxValue = null,
                    //        Description = $"Attribute of {tableName}",
                    //        IsNullable = true,
                    //        DefaultValue = null,
                    //        True = null,
                    //        False = null
                    //    };

                    //    columns.Add(columnDTO);
                    //}
                    //return columns;


                    var table = await _context.TableMetaDataEntity
                      .FirstOrDefaultAsync(t => t.AccessKey.ToLower() == accessKey.ToLower() && t.SecretKey.ToLower() == secretKey.ToLower() && t.Region.ToLower() == region.ToLower() && t.EntityName.ToLower() == tableName.ToLower());

                    if (table == null)
                    {
                        return null;
                    }

                    var column = await _context.ColumnMetaDataEntity
                            .Where(column => column.EntityId == table.Id)
                            .ToListAsync();
                    var columnDTOs = column.Select(column => new Model.DTO.ColumnDTO
                    {
                        Id = column.Id,
                        ColumnName = column.ColumnName,
                        Datatype = column.Datatype,
                        IsPrimaryKey = column.IsPrimaryKey,
                        IsForeignKey = column.IsForeignKey,
                        EntityId = column.EntityId,
                        ReferenceEntityID = column.ReferenceEntityID,
                        ReferenceColumnID = column.ReferenceColumnID,
                        Length = column.Length,
                        MinLength = column.MinLength,
                        MaxLength = column.MaxLength,
                        MaxRange = column.MaxRange,
                        MinRange = column.MinRange,
                        DateMinValue = column.DateMinValue,
                        DateMaxValue = column.DateMaxValue,
                        Description = column.Description,
                        IsNullable = column.IsNullable,
                        DefaultValue = column.DefaultValue,
                        True = column.True,
                        False = column.False,
                        // Include other properties as needed
                    }).ToList();

                    return columnDTOs;

                }

                if (provider.Equals("Influx", StringComparison.OrdinalIgnoreCase))
                {
                    var table = await _context.TableMetaDataEntity
                       .FirstOrDefaultAsync(t => t.InfluxDbBucket.ToLower() == influxDbBucket.ToLower() && t.InfluxDbToken.ToLower() == influxDbToken.ToLower() && t.InfluxDbOrg.ToLower() == influxDbOrg.ToLower() && t.EntityName.ToLower() == tableName.ToLower() && t.InfluxDbUrl.ToLower() == influxDbUrl.ToLower());

                    if (table == null)
                    {
                        return null;
                    }

                    var column = await _context.ColumnMetaDataEntity
                            .Where(column => column.EntityId == table.Id)
                            .ToListAsync();
                    var columnDTOs = column.Select(column => new Model.DTO.ColumnDTO
                    {
                        Id = column.Id,
                        ColumnName = column.ColumnName,
                        Datatype = column.Datatype,
                        IsPrimaryKey = column.IsPrimaryKey,
                        IsForeignKey = column.IsForeignKey,
                        EntityId = column.EntityId,
                        ReferenceEntityID = column.ReferenceEntityID,
                        ReferenceColumnID = column.ReferenceColumnID,
                        Length = column.Length,
                        MinLength = column.MinLength,
                        MaxLength = column.MaxLength,
                        MaxRange = column.MaxRange,
                        MinRange = column.MinRange,
                        DateMinValue = column.DateMinValue,
                        DateMaxValue = column.DateMaxValue,
                        Description = column.Description,
                        IsNullable = column.IsNullable,
                        DefaultValue = column.DefaultValue,
                        True = column.True,
                        False = column.False,
                        // Include other properties as needed
                    }).ToList();

                    return columnDTOs;
                }

                if (provider.Equals("Scylla", StringComparison.OrdinalIgnoreCase))
                {
                    var table = await _context.TableMetaDataEntity
                       .FirstOrDefaultAsync(t => t.IPAddress.ToLower() == ipAddress.ToLower() && t.Keyspace.ToLower() == keyspace.ToLower() && t.Ec2Instance.ToLower() == ec2Instance.ToLower() && t.EntityName.ToLower() == tableName.ToLower());

                    if (table == null)
                    {
                        return null;
                    }

                    var column = await _context.ColumnMetaDataEntity
                            .Where(column => column.EntityId == table.Id)
                            .ToListAsync();
                    var columnDTOs = column.Select(column => new Model.DTO.ColumnDTO
                    {
                        Id = column.Id,
                        ColumnName = column.ColumnName,
                        Datatype = column.Datatype,
                        IsPrimaryKey = column.IsPrimaryKey,
                        IsForeignKey = column.IsForeignKey,
                        EntityId = column.EntityId,
                        ReferenceEntityID = column.ReferenceEntityID,
                        ReferenceColumnID = column.ReferenceColumnID,
                        Length = column.Length,
                        MinLength = column.MinLength,
                        MaxLength = column.MaxLength,
                        MaxRange = column.MaxRange,
                        MinRange = column.MinRange,
                        DateMinValue = column.DateMinValue,
                        DateMaxValue = column.DateMaxValue,
                        Description = column.Description,
                        IsNullable = column.IsNullable,
                        DefaultValue = column.DefaultValue,
                        True = column.True,
                        False = column.False,
                        // Include other properties as needed
                    }).ToList();

                    return columnDTOs;
                }

                else
                    {
                    var table = await _context.TableMetaDataEntity
                       .FirstOrDefaultAsync(t => t.HostName.ToLower() == hostName.ToLower() && t.Provider.ToLower() == provider.ToLower() && t.DatabaseName.ToLower() == databaseName.ToLower() && t.EntityName.ToLower() == tableName.ToLower());

                    if (table == null)
                    {
                        return null;
                    }

                    var column = await _context.ColumnMetaDataEntity
                            .Where(column => column.EntityId == table.Id)
                            .ToListAsync();
                    var columnDTOs = column.Select(column => new Model.DTO.ColumnDTO
                    {
                        Id = column.Id,
                        ColumnName = column.ColumnName,
                        Datatype = column.Datatype,
                        IsPrimaryKey = column.IsPrimaryKey,
                        IsForeignKey = column.IsForeignKey,
                        EntityId = column.EntityId,
                        ReferenceEntityID = column.ReferenceEntityID,
                        ReferenceColumnID = column.ReferenceColumnID,
                        Length = column.Length,
                        MinLength = column.MinLength,
                        MaxLength = column.MaxLength,
                        MaxRange = column.MaxRange,
                        MinRange = column.MinRange,
                        DateMinValue = column.DateMinValue,
                        DateMaxValue = column.DateMaxValue,
                        Description = column.Description,
                        IsNullable = column.IsNullable,
                        DefaultValue = column.DefaultValue,
                        True = column.True,
                        False = column.False,
                        // Include other properties as needed
                    }).ToList();

                    return columnDTOs;
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred while fetching all tables.", ex);
            }
        }
    }
}
