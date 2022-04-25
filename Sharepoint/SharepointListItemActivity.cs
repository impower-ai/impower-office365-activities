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
    public abstract class SharepointListItemActivity : SharepointSiteActivity
    {
        [RequiredArgument]
        [DisplayName("List ID")]
        public InArgument<ListLocator> ListLocator { get; set; }
        [RequiredArgument]
        [DisplayName("ListItem ID")]
        public InArgument<ListItemLocator> ListItemLocator { get; set; }

        protected string ListIdValue;
        protected string ListItemIdValue;
        protected List ListValue;
        protected ListItem ListItemValue;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            ListIdValue = context.GetValue(ListLocator);
            ListItemIdValue = context.GetValue(ListItemLocator);
        }

        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            await base.Initialize(client, context, token);

            try
            {
                ListValue = await client.GetSharepointList(token, SiteValue.Id, ListIdValue);
            }
            catch(Exception e)
            {
                throw new Exception("An Error Occured While Trying To Retrieve The Specified List.",e);
            }
            try
            {
                ListItemValue = await client.GetSharepointListItem(token, SiteValue.Id, ListIdValue, ListItemIdValue);
            }
            catch(Exception e)
            {
                throw new Exception("An Error Occured While Trying To Retrieve The Specified ListItem",e);
            }

        }
    }
}
