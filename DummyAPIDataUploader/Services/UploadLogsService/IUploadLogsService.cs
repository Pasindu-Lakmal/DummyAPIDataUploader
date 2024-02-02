namespace DummyAPIDataUploader.Services.UploadLogsService
{
    public interface IUploadLogsService
    {
        Task<List<UploadLogDetail>> GetAllLogs();
        Task<UploadLogDetail?> GetSingleLog(int id);
        Task<List<UploadLogDetail>> AddUploadLog(UploadLogDetail logDetail);

        Task<List<UploadLogDetail>?> UpdateUploadLog(int id, UploadLogDetail requests);

        Task<List<UploadLogDetail>?> DeleteUploadLog(int id);

    }
}
