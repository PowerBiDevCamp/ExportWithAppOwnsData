import * as $ from 'jquery';

export class ExportFileRequest {
  ExportType: string;
  Filter?: string;
  BookmarkState?: string;
  PageName?: string;
  VisualName?: string;
}

export class Exporter {

  static apiRoot: string = "https://localhost:44300/api/";
  
  static ExportFile = async (ExportRequest: ExportFileRequest): Promise<void> => {

    var restUrl: string = this.apiRoot + "ExportFile/";

    var oReq = new XMLHttpRequest();
    oReq.open("POST", restUrl, true);
    oReq.responseType = "blob";

    // create function to implememt onload callback behavior
    oReq.onload = function (oEvent) {

      var blobData = oReq.response;
      var blob = new Blob([blobData], { type: oReq.response.type });

      var filename = "";
      var disposition = oReq.getResponseHeader('Content-Disposition');
      if (disposition && disposition.indexOf('attachment') !== -1) {
        var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
        var matches = filenameRegex.exec(disposition);
        if (matches != null && matches[1]) {
          filename = matches[1].replace(/['"]/g, '');
        }
      }


      // Download the File
      console.log("Downloading file named " + filename);
      var url = window.URL || window.webkitURL;
      var link = url.createObjectURL(blob);
      var a: JQuery = $("<a />");
      a.attr("download", filename);
      a.attr("href", link);
      $("body").append(a);
      a[0].click();
      a.remove();

    };

    // execute request against custom ExportFile Web API
    oReq.setRequestHeader("Content-Type", "application/json");
    oReq.send(JSON.stringify(ExportRequest));

  }

}
