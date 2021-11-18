
import * as powerbi from "powerbi-client";
import * as models from "powerbi-models";

import { UiBuilder } from './uibuilder';

import { Exporter, ExportFileRequest } from './exporter';

// ensure Power BI JavaScript API has loaded
require('powerbi-models');
require('powerbi-client');

export class ViewModel {
  reportId: string;
  embedUrl: string;
  paginatedReportId: string;
  paginatedEmbedUrl: string;
  turduckenReportId: string;
  turduckenEmbedUrl: string;
  token: string
}

export class Icons {
  static readonly iconPPTX: string = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAAZwAAAGcB1SjUJgAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAEYSURBVDiNndKxKwVQFAbw39FbXmKQzSBlkQwMUkoZiMFkNPkX/AEGk0H5F56ZRSlKMslIMrIRKZQigzqW+/To0Xu+Ot2vc+/57jnfvZGZICI20Ot3TGI6M2++ZTNTEanVebPAPS7Q3Zjv+OPGZhjBdkRU6ol2BD7KOovN/wiMY6jEl1cViIhB9EXEXItivRExmJlX0Ik3ZJvxhs4KulDFCi4xg7vCp/CAfSziBfPYwg66vtzEc5nzEH2YwDGOMIp1XGMMB/WinyYOYLfwfuz92D/DME6bCcxioYzTyGGprLdF9L1RtVrab9fEZ1QjM0VED2pY1RrWsJyZjxXIzKeIeMrM81aqy9lH2vuJTdH4jK8R0VIHOKmTTxSQgRUANXbNAAAAAElFTkSuQmCC";
  static readonly iconPDF: string = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAAnNCSVQICFXsRgQAAAAJcEhZcwAAAGcAAABnAdUo1CYAAAAZdEVYdFNvZnR3YXJlAHd3dy5pbmtzY2FwZS5vcmeb7jwaAAAA6klEQVQoU23OrUuDARAH4CcYBMEkKGKybFbLbIJg0X9gwtBgWREWzDpEGKIwQUUwyNsGsrBiWFOjScSmwSC6pFnLGfahezd+cHB3D8cJwZGkL89mQjvtknTbzrDl0fgAsC/bA6FpJA32FDrgTQgnaTDv2kQIpmVl1dJgU8axTO+TJA12glF161bMDgBjblQdKthQsyvXB0y6tdo7PiWfumBOpbse9kNDfkgaf+BB1ZY7TUXnvlwpOvDwHyxoyCg7c2pbyYWPNFh0oSzxYklJTiUNWkLZmmV1JQ3xH1x68iO8u/ctvPoULkP4BVtKvpjKykMaAAAAAElFTkSuQmCC";
  static readonly iconPNG: string = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAAnNCSVQICFXsRgQAAAAJcEhZcwAAAGcAAABnAdUo1CYAAAAZdEVYdFNvZnR3YXJlAHd3dy5pbmtzY2FwZS5vcmeb7jwaAAAA5UlEQVQoU43QIUuDYRSA0YNYxkDQBT+GcWGC4NYtNv+J9QuiVZNFMcjMQxAtlhUZqzZBVGwLLhjGEKaTDQzzGnSosIk8cOG9HG54hWBP9VdNC+Gzz1EdPb+WbXdm/gahbnoyeBTC4WSQV1R0+gUUNKyNqaEQyBq4cK6j70zPg2vhTd2lgSyJkNgw69qULVWpOwWbdoVkBE7kXSlbsi+17silA33zI7DtXkvZsUWp1I5bee3vC89CS8nQslTTqhU5HQkZXR3vwtCT8KonhBdDXRnBnJrSmGpyEz4qfmz/CSpuxlQJ4QMr17O3yep71gAAAABJRU5ErkJggg==";
}

export class Playbook {

  public static EmbedPowerBiReport(reportContainer: HTMLElement, viewModel: ViewModel): void {

    UiBuilder.ClearEmbedToolbar();
    window.powerbi.reset(reportContainer);

      var config: powerbi.IEmbedConfiguration = {
      type: 'report',
      id: viewModel.reportId,
      embedUrl: viewModel.embedUrl,
      accessToken: viewModel.token,
      permissions: models.Permissions.All,
      tokenType: models.TokenType.Embed,
      viewMode: models.ViewMode.View,
      settings: {
        panes: {
          pageNavigation: {
            visible: true,
            position: models.PageNavigationPosition.Left
          }
        },
        commands: [{ // hide built-in visual commands
          spotlight: { displayOption: models.CommandDisplayOption.Hidden },
          exportData: { displayOption: models.CommandDisplayOption.Hidden },
          seeData: { displayOption: models.CommandDisplayOption.Hidden }
        }],
        extensions: [
          { command: { name: "png", title: "Export Visual to PNG", extend: {
              visualOptionsMenu: {
                title: "Export Visual to PNG", menuLocation: models.MenuLocation.Top, icon: Icons.iconPNG
              }
            }}
          },
          { command: { name: "pptx", title: "Export Visual to PPTX", extend: { 
              visualOptionsMenu: {
                title: "Export Visual to PPTX", menuLocation: models.MenuLocation.Top, icon: Icons.iconPPTX
              }
            }}
          },
          { command: { name: "pdf", title: "Export Visual to PDF", extend: {
                visualOptionsMenu: {
                  title: "Export Visual to PDF", menuLocation: models.MenuLocation.Top, icon: Icons.iconPDF
                }
              }
            }
          }
        ]        
      }
    };

    var report: powerbi.Report = <powerbi.Report>window.powerbi.embed(reportContainer, config);

    console.log(report);

    var exportReport = async (ExportType: string) => {

      var bookmarkResult: models.IReportBookmark =
        await report.bookmarksManager.capture({
          personalizeVisuals: false,
          allPages: true
        });

      var exportRequest: ExportFileRequest = {
        ExportType: ExportType,
        BookmarkState: bookmarkResult.state
      };

      console.log("Submitting Export Request", exportRequest)
      var exportedFileData = await Exporter.ExportFile(exportRequest);
      UiBuilder.DisplayNotification("Report being exported as " + ExportType)

    };

    var exportReportPage = async (ExportType: string) => {

      var bookmarkResult: models.IReportBookmark =
        await report.bookmarksManager.capture({
          personalizeVisuals: false,
          allPages: false
        });

      var currentPage: powerbi.Page = await report.getActivePage();

      var exportRequest: ExportFileRequest = {
        ExportType: ExportType,
        BookmarkState: bookmarkResult.state,
        PageName: currentPage.name
      };

      var exportedFileData = await Exporter.ExportFile(exportRequest);
      UiBuilder.DisplayNotification("Current Report Page being export as " + ExportType);

    }

    var exportReportPageVisual = async (ExportType: string, PageName: string, VisualName: string) => {

      let slicerFilters: string[] = [];
      let currentPage: powerbi.Page = await report.getActivePage();
      let visuals = await currentPage.getVisuals();

      for (let visual of visuals) {
        if (visual.type == "slicer") {
          console.log(visual.name)
          let slicerState: any = await visual.getSlicerState();
          for (let filter of slicerState.filters) {
            console.log("filter");
            if (filter.operator == "In") {
              console.log(filter);
              let target = filter.target;
              let values = filter.values;
              let tableName = target.table;
              let columnName = target.column;
              if (isNaN(Number(values[0]))) {
                slicerFilters.push(tableName + "/" + columnName + " in ('" + values.join("', '") + "')");
              }
              else {
                slicerFilters.push(tableName + "/" + columnName + " in (" + values.join(", ") + ")");
              }
            }
          }
        }
      }

      var exportRequest: ExportFileRequest = {
        ExportType: ExportType,
        PageName: PageName,
        VisualName: VisualName
      };

      if (slicerFilters.length > 0) {
        exportRequest.Filter = slicerFilters.join(" and ");
      }

      console.log("Export Visual to File");
      console.log(exportRequest);

      var exportedFileData = await Exporter.ExportFile(exportRequest);
      UiBuilder.DisplayNotification("Current Report Page being export as PDF.")


    }

    UiBuilder.AddCommandButton("Export Report as PDF", async () => {
      exportReport("pdf");
    });

    UiBuilder.AddCommandButton("Export Report as PPTX", async () => {
      exportReport("pptx");
    });

    UiBuilder.AddCommandButton("Export Report as PNG", async () => {
      exportReport("png");
    });

    UiBuilder.AddCommandButton("Export Page as PDF", async () => {
      exportReportPage("pdf");
    });

    UiBuilder.AddCommandButton("Export Page as PPTX", async () => {
      exportReportPage("pptx");
    });

    UiBuilder.AddCommandButton("Export Page as PNG", async () => {
      exportReportPage("png");
    });

    report.on("commandTriggered", async (event: any) => {

      // get event data for user invoking visual export menu command
      let ExportType: string = event.detail.command;
      let PageName: string = event.detail.page.name;
      let VisualName: string = event.detail.visual.name;

      exportReportPageVisual(ExportType, PageName, VisualName);
    
    });

  }

  public static EmbedPaginatedReport(reportContainer: HTMLElement, viewModel: ViewModel): void {

    UiBuilder.ClearEmbedToolbar();
    window.powerbi.reset(reportContainer);

    var config: powerbi.IEmbedConfiguration = {
      type: 'report',
      id: viewModel.paginatedReportId,
      embedUrl: viewModel.paginatedEmbedUrl,
      accessToken: viewModel.token,
      permissions: models.Permissions.Read,
      tokenType: models.TokenType.Embed,
      viewMode: models.ViewMode.View,
      settings: {}
    };
  
    var report: powerbi.Report = <powerbi.Report>window.powerbi.embed(reportContainer, config);

  }

  public static EmbedTurduckenReport(reportContainer: HTMLElement, viewModel: ViewModel): void {

    UiBuilder.ClearEmbedToolbar();
    window.powerbi.reset(reportContainer);

    var config: powerbi.IEmbedConfiguration = {
      type: 'report',
      id: viewModel.turduckenReportId,
      embedUrl: viewModel.turduckenEmbedUrl,
      accessToken: viewModel.token,
      permissions: models.Permissions.All,
      tokenType: models.TokenType.Embed,
      viewMode: models.ViewMode.View,
      settings: {}
    };

    var report: powerbi.Report = <powerbi.Report>window.powerbi.embed(reportContainer, config);  

  }

}
