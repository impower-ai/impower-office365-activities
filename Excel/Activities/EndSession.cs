using Impower.Office365.Sharepoint;
using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static Impower.Office365.Excel.ExcelExtensions;

namespace Impower.Office365.Excel
{
    [DisplayName("End Session")]
    public class EndSession : SharepointWorkbookActivity
    {
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            await client.EndWorkbookSession(DriveItemReference, SessionConfiguration, token);   
            return ctx => { };
        }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);

        }
    }
}
