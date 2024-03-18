using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.Base
{
    public interface IDocBase
    {
        string libreOfficeAppPath { get; set; }
        string locationTempPath { get; set; }
    }
}
