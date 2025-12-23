using ExcelTool.Models;
using ExcelTool.Services;
using Microsoft.AspNetCore.Mvc;
namespace ExcelTool.Controllers
{
    [ApiController]
    [Route("api/excel")]
    public class ExcelUploadController : ControllerBase
    {
        private readonly IExcelUploadService _service;

        public ExcelUploadController(IExcelUploadService service)
        {
            _service = service;
        }

        [HttpPost("upload")]
        [Consumes("multipart/form-data")]
        public async Task<ActionResult<ExcelUploadResult>> Upload(
            [FromForm] ExcelUploadRequest request)
        {
            if (request.File == null || request.File.Length == 0)
                return BadRequest("Excel file is required");

            var result = await _service.UploadAsync(
                request.TableName,
                request.File);

            return Ok(result);
        }
    }
}
