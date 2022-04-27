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
    public class UploadListItems : SharepointSiteActivity
    {
        [RequiredArgument]
        [Category("Input")]
        public InArgument<ListLocator> List { get; set; }
        private string listID;
        [Category("Input")]
        [RequiredArgument]       
        public InArgument<DataTable> Data { get; set; }
        private DataTable data;

        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            listID = List.Get(context);
            data = Data.Get(context);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            await client.UploadListItems(token, SiteId, listID, data);
            return ctx => { };
        }
        protected override Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            return base.Initialize(client, context, token);
        }

    }
}
