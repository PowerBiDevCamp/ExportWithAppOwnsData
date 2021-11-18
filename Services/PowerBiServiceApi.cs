using System;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Web;
using Microsoft.Rest;
using Microsoft.PowerBI.Api;
using Microsoft.PowerBI.Api.Models;
using Newtonsoft.Json;
using System.Collections.Generic;
using ExportWithAppOwnsData.Models;
using ExportWithAppOwnsData.Controllers;

namespace ExportWithAppOwnsData.Services {

  public class EmbeddedReportViewModel {
    public string ReportId;
    public string ReportEmbedUrl;
    public string PaginatedReportId;
    public string PaginatedReportEmbedUrl;
    public string TurduckenId;
    public string TurduckenEmbedUrl;
    public string Token;
  }

  public class ExportedReport {
    public string ReportName { get; set; }
    public string ResourceFileExtension { get; set; }
    public Stream ReportStream { get; set; }
  }

  public class PowerBiServiceApi {

    private ITokenAcquisition tokenAcquisition { get; }
    private string urlPowerBiServiceApiRoot { get; }
    private Guid workspaceId { get; }
    private Guid reportId { get; }
    private Guid paginatedReportId { get; }
    private Guid turduckenReportId { get; }

    public PowerBiServiceApi(IConfiguration configuration, ITokenAcquisition tokenAcquisition) {
      this.urlPowerBiServiceApiRoot = configuration["PowerBi:ServiceRootUrl"];
      this.workspaceId = new Guid(configuration["PowerBi:WorkspaceId"]);
      this.reportId = new Guid(configuration["PowerBi:ReportId"]);
      this.paginatedReportId = new Guid(configuration["PowerBi:PaginatedReportId"]);
      this.turduckenReportId = new Guid(configuration["PowerBi:TurduckenReportId"]);
      this.tokenAcquisition = tokenAcquisition;
    }

    public const string powerbiApiDefaultScope = "https://analysis.windows.net/powerbi/api/.default";

    public string GetAccessToken() {
      return this.tokenAcquisition.GetAccessTokenForAppAsync(powerbiApiDefaultScope).Result;
    }

    public PowerBIClient GetPowerBiClient() {
      string accessToken = GetAccessToken();
      var tokenCredentials = new TokenCredentials(accessToken, "Bearer");
      return new PowerBIClient(new Uri(urlPowerBiServiceApiRoot), tokenCredentials);
    }

    public async Task<EmbeddedReportViewModel> GetReportViewModel() {

      PowerBIClient pbiClient = GetPowerBiClient();

      // call to Power BI Service API to get embedding data
      var reportPowerBi = await pbiClient.Reports.GetReportInGroupAsync(this.workspaceId, this.reportId);
      var reportPaginated = await pbiClient.Reports.GetReportInGroupAsync(this.workspaceId, this.paginatedReportId);
      var reportTurducken = await pbiClient.Reports.GetReportInGroupAsync(this.workspaceId, this.turduckenReportId);

      // generate read-only embed token for the report
      //var datasetId = report.DatasetId;

      var datasets = (await pbiClient.Datasets.GetDatasetsInGroupAsync(workspaceId)).Value;
      var reports = (await pbiClient.Reports.GetReportsInGroupAsync(workspaceId)).Value;

      IList<GenerateTokenRequestV2Dataset> datasetRequests = new List<GenerateTokenRequestV2Dataset>();
      IList<string> datasetIds = new List<string>();
      foreach (var dataset in datasets) {
        datasetRequests.Add(new GenerateTokenRequestV2Dataset(dataset.Id));
        datasetIds.Add(dataset.Id);
      };

      IList<GenerateTokenRequestV2Report> reportRequests = new List<GenerateTokenRequestV2Report>();
      foreach (var report in reports) {
        reportRequests.Add(new GenerateTokenRequestV2Report(report.Id));
      };

      GenerateTokenRequestV2 tokenRequest =
        new GenerateTokenRequestV2 {
          Datasets = datasetRequests,
          Reports = reportRequests
        };

      // call to Power BI Service API and pass GenerateTokenRequest object to generate embed token
      string embedToken = (pbiClient.EmbedToken.GenerateToken(tokenRequest)).Token;

      // return report embedding data to caller
      return new EmbeddedReportViewModel {
        ReportId = reportPowerBi.Id.ToString(),
        ReportEmbedUrl = reportPowerBi.EmbedUrl,
        PaginatedReportId = reportPaginated.Id.ToString(),
        PaginatedReportEmbedUrl = reportPaginated.EmbedUrl,
        TurduckenId = reportTurducken.Id.ToString(),
        TurduckenEmbedUrl = reportTurducken.EmbedUrl,
        Token = embedToken
      };
    }
   
    public async Task<ExportedReport> ExportFile(ExportFileRequestParams ExportRequestParams) {

      FileFormat fileFormat;
      switch (ExportRequestParams.ExportType) {
        case "pdf":
          fileFormat = FileFormat.PDF;
          break;
        case "pptx":
          fileFormat = FileFormat.PPTX;
          break;
        case "png":
          fileFormat = FileFormat.PNG;
          break;
        default:
          throw new ApplicationException("Power BI reports do not support exort to " + ExportRequestParams.ExportType);
      }

      PowerBIClient pbiClient = GetPowerBiClient();

      var exportRequest = new ExportReportRequest {
        Format = fileFormat,
        PowerBIReportConfiguration = new PowerBIReportExportConfiguration()
      };

      if (!string.IsNullOrEmpty(ExportRequestParams.Filter)) {
        string[] filters = ExportRequestParams.Filter.Split(";");
        exportRequest.PowerBIReportConfiguration.ReportLevelFilters = new List<ExportFilter>();        
        foreach(string filter in filters) {
          exportRequest.PowerBIReportConfiguration.ReportLevelFilters.Add(new ExportFilter(filter));
        }
      }

      if (!string.IsNullOrEmpty(ExportRequestParams.BookmarkState)) {
        exportRequest.PowerBIReportConfiguration.DefaultBookmark = new PageBookmark { State = ExportRequestParams.BookmarkState };
      }

      if (!string.IsNullOrEmpty(ExportRequestParams.PageName)) {
        exportRequest.PowerBIReportConfiguration.Pages = new List<ExportReportPage>(){
          new ExportReportPage{PageName = ExportRequestParams.PageName}
        };
        if (!string.IsNullOrEmpty(ExportRequestParams.VisualName)) {
          exportRequest.PowerBIReportConfiguration.Pages[0].VisualName = ExportRequestParams.VisualName;
        }
      }

      Export export = await pbiClient.Reports.ExportToFileInGroupAsync(this.workspaceId, this.reportId, exportRequest);

      string exportId = export.Id;

      do {
        System.Threading.Thread.Sleep(3000);
        export = pbiClient.Reports.GetExportToFileStatusInGroup(this.workspaceId, this.reportId, exportId);
      } while (export.Status != ExportState.Succeeded && export.Status != ExportState.Failed);

      if (export.Status == ExportState.Failed) {
        Console.WriteLine("Export failed!");
      }

      if (export.Status == ExportState.Succeeded) {
        return new ExportedReport {
          ReportName = export.ReportName,
          ResourceFileExtension = export.ResourceFileExtension,
          ReportStream = pbiClient.Reports.GetFileOfExportToFileInGroup(this.workspaceId, this.reportId, exportId)
        };
      }
      else {
        return null;
      }
    }

  }

}


