using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExportWithAppOwnsData.Services;
using ExportWithAppOwnsData.Models;
using static System.Net.WebRequestMethods;

namespace ExportWithAppOwnsData.Controllers {

  [Route("api/[controller]")]
  [ApiController]
  public class ExportFileController : ControllerBase {

    private PowerBiServiceApi powerBiServiceApi;

    public ExportFileController(PowerBiServiceApi powerBiServiceApi) {
      this.powerBiServiceApi = powerBiServiceApi;
    }

    [HttpPost]
    public async Task<FileStreamResult> PostExportFile([FromBody] ExportFileRequestParams request) {

      var exportedReport = await this.powerBiServiceApi.ExportFile(request);
      exportedReport.ReportStream.Flush();

      string contentType = getContentTypeFromExtension(exportedReport.ResourceFileExtension);

      var file = new FileStreamResult(exportedReport.ReportStream, contentType);
      file.FileDownloadName = exportedReport.ReportName + exportedReport.ResourceFileExtension;
      return file;

    }

    private string getContentTypeFromExtension(string extension) {
      switch (extension) {
        case ".pdf":
          return "application/pdf";
        case ".pptx":
          return "application/pptx";
        case ".zip":
          return "application/zip";
        case ".png":
          return "image/png";
        default:
          throw new ApplicationException("Cannot handle extension of " + extension);
      }
    }


  }
}
