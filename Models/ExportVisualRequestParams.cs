using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExportWithAppOwnsData.Models {
  public class ExportVisualRequestParams {
    public string ExportType { get; set; }
    public string PageName { get; set; }
    public string VisualName { get; set; }
  }
}
