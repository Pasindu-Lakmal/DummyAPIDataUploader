namespace DummyAPIDataUploader.Services.UploadLogsService
{
  
    public class UploadLogsService : IUploadLogsService
    {
        
        private readonly DataContext _context;

        public UploadLogsService(DataContext context)
        {
            _context = context;
        }
        public async Task<List<UploadLogDetail>> AddUploadLog(UploadLogDetail logDetail)
        {
            _context.UploadLogDetails.Add(logDetail);
            await _context.SaveChangesAsync();
            
            return await _context.UploadLogDetails.ToListAsync();
        }

        public async Task<List<UploadLogDetail>?> DeleteUploadLog(int id)
        {
            var uploadLog = await _context.UploadLogDetails.FindAsync(id);

            if (uploadLog is null)
            {
                return null;
            }
            else
            {
                _context.UploadLogDetails.Remove(uploadLog);
                await _context.SaveChangesAsync();
                return await _context.UploadLogDetails.ToListAsync();
            }


            
        }

        public async Task<List<UploadLogDetail>> GetAllLogs()
        {
            var uploadLogs = await _context.UploadLogDetails.ToListAsync();
            return uploadLogs;
        }

        public async Task<UploadLogDetail?> GetSingleLog(int id)
        {
            var uploadLog = await _context.UploadLogDetails.FindAsync(id);
            if (uploadLog is null)
            {
                return null;
            }else 
            { 
                return uploadLog; 
            }
            
        }

        public async Task<List<UploadLogDetail>?> UpdateUploadLog(int id, UploadLogDetail requests)
        {
            var uploadLog = await _context.UploadLogDetails.FindAsync(id);
            if (uploadLog is null)
            {
                return null;
            }
            else
            {
                uploadLog.StaffData = requests.StaffData;
                uploadLog.UploadedDate = requests.UploadedDate;
                uploadLog.FileName = requests.FileName;
                uploadLog.UploadedBy = requests.UploadedBy;
                uploadLog.FileType = requests.FileType;

                await _context.SaveChangesAsync();

                return await _context.UploadLogDetails.ToListAsync();
            }
 
        }
    }
}
