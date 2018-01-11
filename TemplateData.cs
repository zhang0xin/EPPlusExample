using System.Collections.Generic;
using System.Data;

namespace EPPlusExample
{
  class TemplateData
  {
    public Dictionary<string, object> Fields {get; set;}
    public Dictionary<string, DataTable> Tables {get; set;}
    public TemplateData()
    {
      Fields = new Dictionary<string, object>();
      Tables = new Dictionary<string, DataTable>();
    }
  }
}
