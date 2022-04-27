using Impower.Office365.Sharepoint;
using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static Impower.Office365.Excel.ExcelExtensions;

namespace Impower.Office365.Excel.Activities
{
    public class RefreshSharepointWorkbook : SharepointWorkbookActivity
    {
        public InArgument<TimeSpan> RefreshInterval { get; set; }
        internal TimeSpan RefreshIntervalValue;
        public InArgument<TimeSpan> Timeout { get; set; }
        internal TimeSpan TimeoutValue;
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            await client.RecalculateSharepointWorkbook(
                CalculationType.FullRebuild,
                RefreshIntervalValue,
                TimeoutValue,
                token,
                SiteId,
                DriveId,
                DriveItemId
                );
            return ctx => { };
        }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            TimeoutValue = Timeout.Get(context);
            RefreshIntervalValue = RefreshInterval.Get(context);
            base.ReadContext(context);

        }
    }
}
