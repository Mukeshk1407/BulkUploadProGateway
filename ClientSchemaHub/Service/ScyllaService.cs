
using Cassandra;
using ClientSchemaHub.Models.DTO;
using ClientSchemaHub.Service.IService;
using System.Dynamic;

namespace ClientSchemaHub.Service
{
    public class ScyllaService : IScyllaService
    {
        public async Task<Dictionary<string, List<TableDetailsDTO>>> GetTableDetailsForAllTablesAsync(DBConnectionDTO connectionDTO)
        {
            var session = GetScyllaDbSession(connectionDTO);
            var tableNames = await GetTableNamesAsync(session, connectionDTO.Keyspace);
            Dictionary<string, List<TableDetailsDTO>> tableDetailsDictionary = new Dictionary<string, List<TableDetailsDTO>>();

            foreach (var tableName in tableNames)
            {
                var tableDetails = await GetTableDetailsAsync(session, connectionDTO.Keyspace, tableName);

                if (!tableDetailsDictionary.ContainsKey(tableName))
                {
                    tableDetailsDictionary[tableName] = new List<TableDetailsDTO>();
                }

                tableDetailsDictionary[tableName].Add(tableDetails);
            }

            return tableDetailsDictionary;
        }

        private Cassandra.ISession GetScyllaDbSession(DBConnectionDTO connectionDTO)
        {
            var cluster = Cluster.Builder()
                .AddContactPoint(connectionDTO.IPAddress)
                .WithPort(connectionDTO.PortNumber ?? 9042)
                .Build();
            return cluster.Connect(connectionDTO.Keyspace);
        }

        private async Task<List<string>> GetTableNamesAsync(Cassandra.ISession session, string keyspace)
        {
            var tableNames = new List<string>();

            try
            {
                var query = $"SELECT table_name FROM system_schema.tables WHERE keyspace_name = '{keyspace}';";
                var rowSet = await session.ExecuteAsync(new SimpleStatement(query));

                foreach (var row in rowSet)
                {
                    tableNames.Add(row.GetValue<string>("table_name"));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                throw;
            }

            return tableNames;
        }

        public async Task<List<string>> GetTableNamesAsync(DBConnectionDTO connectionDTO)
        {
            var session = GetScyllaDbSession(connectionDTO);
            return await GetTableNamesAsync(session, connectionDTO.Keyspace);
        }

        private async Task<TableDetailsDTO> GetTableDetailsAsync(Cassandra.ISession session, string keyspace, string tableName)
        {
            var tableDetails = new TableDetailsDTO { TableName = tableName };

            try
            {
                var query = $@"
                    SELECT column_name, type 
                    FROM system_schema.columns 
                    WHERE keyspace_name = '{keyspace}' AND table_name = '{tableName}';";

                var rowSet = await session.ExecuteAsync(new SimpleStatement(query));

                tableDetails.Columns = rowSet.Select(row => new ColumnDetailsDTO
                {
                    ColumnName = row.GetValue<string>("column_name"),
                    DataType = row.GetValue<string>("type")
                }).ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching details for table {tableName}: {ex.Message}");
                throw;
            }

            return tableDetails;
        }

        public async Task<TableDetailsDTO> GetTableDetailsAsync(DBConnectionDTO connectionDTO, string tableName)
        {
            var session = GetScyllaDbSession(connectionDTO);
            return await GetTableDetailsAsync(session, connectionDTO.Keyspace, tableName);
        }

        public async Task<List<dynamic>> GetTableData(DBConnectionDTO dBConnection, string tableName)
        {
            var session = GetScyllaDbSession(dBConnection);

            // Query the ScyllaDB table to get all rows
            var query = $"SELECT * FROM {dBConnection.Keyspace}.{tableName}";

            try
            {
                var rowSet = await session.ExecuteAsync(new SimpleStatement(query));

                // Assume column names are fetched beforehand or known
                var columnNames = rowSet.Columns.Select(col => col.Name).ToList(); // Get column names

                var data = rowSet.Select(row => ConvertToDynamic(row, columnNames)).ToList();

                return data;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Error retrieving data from table {tableName}: {ex.Message}");
            }
        }

        // Helper method to convert ScyllaDB Row to dynamic object
        private dynamic ConvertToDynamic(Row row, List<string> columnNames)
        {
            var expandoObj = new ExpandoObject() as IDictionary<string, object>;

            foreach (var columnName in columnNames)
            {
                expandoObj[columnName] = row.GetValue<object>(columnName);
            }

            return expandoObj;
        }


        public async Task<bool> IsTableExists(DBConnectionDTO connectionDTO, string tableName)
        {
            try
            {
                var session = GetScyllaDbSession(connectionDTO);

                // Query to check if the table exists in the specified keyspace
                var query = $@"
            SELECT COUNT(*) 
            FROM system_schema.tables 
            WHERE keyspace_name = '{connectionDTO.Keyspace}' 
            AND table_name = '{tableName}';";

                var rowSet = await session.ExecuteAsync(new SimpleStatement(query));

                // Fetch the first row from the result set and check if the count is greater than 0
                var row = rowSet.FirstOrDefault();
                return row != null && row.GetValue<long>("count") > 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking if table exists: {ex.Message}");
                throw;
            }
        }


        public async Task<List<object>> GetPrimaryColumnDataAsync(DBConnectionDTO dBConnection, string tableName)
        {
            var session = GetScyllaDbSession(dBConnection);

            // Get primary key columns
            List<string> primaryKeyColumns = new List<string>();
            try
            {
                var query = $@"
        SELECT column_name 
        FROM system_schema.columns 
        WHERE keyspace_name = '{dBConnection.Keyspace}' 
        AND table_name = '{tableName}' 
        AND kind = 'partition_key' ALLOW FILTERING;";  // Adding ALLOW FILTERING

                var rowSet = await session.ExecuteAsync(new SimpleStatement(query));

                primaryKeyColumns = rowSet.Select(row => row.GetValue<string>("column_name")).ToList();

                if (primaryKeyColumns.Count == 0)
                {
                    throw new InvalidOperationException($"Table '{tableName}' does not have a primary key.");
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Error fetching primary key columns for table {tableName}: {ex.Message}");
            }


            // Prepare the SELECT query to retrieve primary key data
            var primaryKeySelect = string.Join(", ", primaryKeyColumns);
            var selectQuery = $"SELECT {primaryKeySelect} FROM {dBConnection.Keyspace}.{tableName};";

            try
            {
                var result = await session.ExecuteAsync(new SimpleStatement(selectQuery));
                var primaryKeys = new List<object>();

                foreach (var row in result)
                {
                    var primaryKeyValues = new ExpandoObject() as IDictionary<string, object>;

                    foreach (var column in primaryKeyColumns)
                    {
                        primaryKeyValues[column] = row.GetValue<object>(column);
                    }

                    primaryKeys.Add(primaryKeyValues);
                }

                return primaryKeys.Cast<object>().ToList();
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Error retrieving primary key data from table {tableName}: {ex.Message}");
            }
        }



    }
}
