using Impower.Office365.Sharepoint.Models;
using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Sharepoint
{
    public class UploadListItems : SharepointListActivity
    {
        [Category("Input")]
        [RequiredArgument]       
        public InArgument<DataTable> DataArgument { get; set; }
        private DataTable Data;

        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            Data = DataArgument.Get(context);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            await client.UploadListItems(token, ListReference, Data);
            return ctx => { };
        }

    }
}
