using DummyAPIDataUploader.Services.UploadLogsService;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DummyAPIDataUploader.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UploadLogController : ControllerBase
    {
        private readonly IUploadLogsService _uploadLogsService;

        public UploadLogController(IUploadLogsService uploadLogsService)
        {
            _uploadLogsService = uploadLogsService;
        }


        [HttpGet]
        public async Task<ActionResult<List<UploadLogDetail>>> GetAllLogs()
        {
            var result = await _uploadLogsService.GetAllLogs();
        
            return Ok(result);
        }

        [HttpGet("{id}")]
        public async Task<ActionResult<UploadLogDetail>> GetSingleLog(int id)
        {
            var result = await _uploadLogsService.GetSingleLog(id);
            if (result is null)
            {
                return NotFound("Sorry, selected Upload Not Found");
            }
            else
            {
                return Ok(result);
            }
        }

        [HttpPost]
        public async Task<ActionResult<List<UploadLogDetail>>> AddUploadLog(UploadLogDetail logDetail)
        {
            var result =await  _uploadLogsService.AddUploadLog(logDetail);
            
            return Ok(result);
            

        }

        [HttpPut("{id}")]
        public async Task<ActionResult<List<UploadLogDetail>>> UpdateUploadLog(int id , UploadLogDetail requests)
        {
            var result =await _uploadLogsService.UpdateUploadLog(id, requests);
            
            if (result is null)
            {
                return NotFound("Sorry, selected Upload List Not Found");
            }
            else
            {
                return Ok(result);
            }

        }

        [HttpDelete("{id}")]
        public async Task<ActionResult<List<UploadLogDetail>>> DeleteUploadLog(int id)
        {
            var result =await _uploadLogsService.DeleteUploadLog(id);
            
            if(result is null)
            {
                return NotFound("Sorry, selected Upload List Not Found");
            }
            else
            {
                return Ok(result);
            }

        }
    }
}
