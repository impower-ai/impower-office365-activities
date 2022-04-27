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
    public class CreateSession : SharepointWorkbookActivity
    {
        [Category("Input")]
        [DefaultValue(true)]
        [DisplayName("Persist Changes")]
        public InArgument<bool> PersistChanges { get; set; }
        internal bool PersistChangesValue;
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            client.Cr
            return ctx => { };
        }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            PersistChangesValue = PersistChanges.Get(context);
            base.ReadContext(context);

        }
    }
}
