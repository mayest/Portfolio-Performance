using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace PortfolioPerformance
{
#if !DEBUG
    class ExcelAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // Versions before v1.1.0 required only a call to Register() in the AutoOpen().
            // The name was changed (and made obsolete) to highlight the pair of function calls now required.
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
#endif
}
