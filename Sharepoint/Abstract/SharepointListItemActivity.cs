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
    public abstract class SharepointListItemActivity : SharepointListActivity
    {
        [RequiredArgument]
        [DisplayName("ListItem ID")]
        public InArgument<ListItemLocator> ListItemLocator { get; set; }

        protected ListItemReference ListItemReference;
        protected ListItem ListItem;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            ListItemReference = ListReference.Item(context.GetValue(ListItemLocator));
        }

        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            await base.Initialize(client, context, token);
            try
            {
                ListItem = await ListItemReference.Get(client, token);
            }
            catch(Exception e)
            {
                throw new Exception("An Error Occured While Trying To Retrieve The Specified ListItem",e);
            }

        }
    }
}
