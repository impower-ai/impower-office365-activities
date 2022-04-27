using Impower.Office365.Sharepoint.Models;
using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Sharepoint
{
    public abstract class SharepointListActivity : SharepointSiteActivity
    {
        [RequiredArgument]
        [DisplayName("List ID")]
        public InArgument<ListLocator> ListLocator { get; set; }

        protected ListReference ListReference;
        protected List List;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            ListReference = SiteReference.List(context.GetValue(ListLocator));
        }

        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            await base.Initialize(client, context, token);

            try
            {
                List = await ListReference.Get(client, token);
            }
            catch(Exception e)
            {
                throw new Exception("An Error Occured While Trying To Retrieve The Specified List.",e);
            }
        }
    }
}
