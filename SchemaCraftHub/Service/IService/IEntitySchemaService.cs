using SchemaCraftHub.Model.DTO;

namespace SchemaCraftHub.Service.IService
{
    public interface IEntitySchemaService
    {
        public Task<List<TableMetaDataDTO>> GetAllTablesAsync();
        public Task<TableMetaDataDTO> GetTableByIdAsync(int id);
        public Task<List<TableMetaDataDTO>> GetTablesByHostProviderDatabaseAsync(string? hostName, string provider, string? databaseName, string? accessKey, string? secretKey, string? region, string? keyspace, string? ec2Instance, string? ipAddress, string? influxDbToken, string? influxDbOrg, string? influxDbUrl, string? influxDbBucket);
        public Task<TableMetaDataDTO> GetTableByHostProviderDatabaseTableNameAsync(string? hostName, string provider, string? databaseName, string? accessKey, string? secretKey, string? region, string? keyspace, string? ipAddress, string? ec2Instance, string? tableName, string? influxDbToken, string? influxDbOrg, string? influxDbUrl, string? influxDbBucket);
        public Task<List<ColumnDTO>> GetAllColumnsAsync();
        public Task<ColumnDTO> GetColumnByIdAsync(int id);
        public Task<ColumnDTO> GetColumnByIdAndEntityIDAsync(int id, int entityId);
        public Task<List<ColumnDTO>> GetColumnsByEntityIdAsync(int entityId);
        public Task<int> CreateTableAsync(TableMetaDataDTO tableDTO);
        public Task InsertColumnsAsync(List<ColumnDTO> columns);
        public Task<Dictionary<string, List<ClientSchemaHub.Models.DTO.TableDetailsDTO>>> GetClientSchema(APIResponse tabledetails1, DBConnectionDTO connectionDTO);
        public Task<APIResponse> convertandcallcreatetablemodel(DBConnectionDTO connectionDTO, TableRequest tableRequest);
        public Task UpdateColumnsAsync(List<ColumnDTO> columns);
        public Task<List<ColumnDTO>> GetColumnsByHostProviderDatabaseTableNameAsync(string? hostName, string provider, string? databaseName, string? tableName, string? accessKey, string? secretKey, string? region, string keyspace, string ec2Instance, string ipAddress, string? influxDbToken, string? influxDbOrg, string? influxDbUrl, string? influxDbBucket);

    }
}
