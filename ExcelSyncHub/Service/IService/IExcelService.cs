using DBUtilityHub.Models;
using ExcelSyncHub.Model.DTO;
using Spire.Xls;
using System.Data;

namespace ExcelSyncHub.Service.IService
{
    public interface IExcelService
    {
        public byte[] GenerateExcelFile(List<ColumnMetaDataDTO> columns, int? parentId, string UserName, string Password, string DataBase, string HostName);
        public int GetEntityIdByEntityName(string entityName);
        public DataTable ReadExcelFromFormFile(IFormFile excelFile);

        public IEnumerable<ColumnMetaDataDTO> GetColumnsForEntity(string entityName);
        public List<Dictionary<string, string>> ReadDataFromExcel(Stream excelFileStream, int rowCount);

        public Task<ValidationResultData> ValidateNotNull(DataTable excelData, List<ColumnMetaDataDTO> columnsDTO);

        public Task<ValidationResultData> ValidatePrimaryKeyAsync(ValidationResultData validationResult, List<ColumnMetaDataDTO> columnsDTO, string tableName, DBConnectionDTO connectionDTO);

        public Task<ValidationResult> resultparams(ValidationResultData validationResult, string comma_separated_string);

        public Task<ValidationResult> resultparamsforprimary(ValidationResultData validationResult, string comma_separated_string, string tableName);

        public Task<ValidationResult> resultparamsforrange(ValidationResultData validationResult, string comma_separated_string, string tableName);
        public Task<ValidationResult> resultparamsforlength(ValidationResultData validationResult, string comma_separated_string, string tableName);

        public Task<LogDTO> Createlog(string tableName, List<string> filedata, string fileName, int successdata, List<string> errorMessage, int total_count, List<string> ErrorRowNumber);

        public Task InsertDataFromDataTableToPostgreSQL(DataTable data, string tableName, List<string> columns, IFormFile file, DBConnectionDTO connectionDTO);

        public int GetEntityIdByEntityNamefromui(string entityName);

        public int? GetEntityIdFromTemplate(IFormFile file, int sheetIndex);

        public string GetPrimaryKeyColumnForEntity(string entityName);

        public bool TableExists(string tableName);

        public List<dynamic> GetTableDataByChecklistEntityValue(DBConnectionDTO connectionDTO, int checklistEntityValue);

        public Task<ValidationResultData> ValidateRange(ValidationResultData validationResult, List<ColumnMetaDataDTO> columnsDTO, string tableName);
        public Task<ValidationResultData> ValidateLength(ValidationResultData validationResult, List<ColumnMetaDataDTO> columnsDTO, string tableName);

    }
}
