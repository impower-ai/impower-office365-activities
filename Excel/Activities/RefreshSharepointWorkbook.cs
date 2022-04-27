using Microsoft.Graph;
using System;
using System.Activities;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;
using static Impower.Office365.Excel.ExcelExtensions;

namespace Impower.Office365.Excel
{
    [DisplayName("Refresh Sharepoint Workbook")]
    public class RefreshSharepointWorkbook : SharepointWorkbookActivity
    {
        [DisplayName("Refresh Interval")]
        [Category("Input")]
        [DefaultValue("0:00:15")]
        public InArgument<TimeSpan> RefreshInterval { get; set; }
        internal TimeSpan RefreshIntervalValue;

        [DisplayName("Timeout")]
        [Category("Config")]
        [DefaultValue("0:05:00")]
        public override InArgument<TimeSpan> Timeout { get; set; }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            await client.RecalculateSharepointWorkbook(
                CalculationType.FullRebuild,
                DriveItemReference,
                SessionConfiguration,
                RefreshIntervalValue,
                TimeoutValue,
                token
                );
            return ctx => { };
        }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            TimeoutValue = Timeout.Get(context);
            RefreshIntervalValue = RefreshInterval.Get(context);

        }
    }
}
