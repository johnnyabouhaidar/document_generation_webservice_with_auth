using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DocGenerationService.Controllers
{
    [Authorize]
    [Route("api/v1/[controller]")]
    [ApiController]
    public class GenerateWordDocumentController : ControllerBase
    {
        [HttpGet]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
               
            })
            .ToArray();
        }
    }
}
