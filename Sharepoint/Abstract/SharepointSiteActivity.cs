using Microsoft.Graph;
using System;
using System.Activities;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using Impower.Office365.Sharepoint.Models;

namespace Impower.Office365.Sharepoint
{
    public abstract class SharepointSiteActivity : Office365Activity
    {
        [Category("Connection")]
        [DisplayName("Sharepoint URL")]
        public InArgument<string> WebURL { get; set; }
        protected SiteReference SiteReference => new SiteReference(Site.Id);
        protected string WebUrlValue;
        protected Site Site;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            WebUrlValue = context.GetValue(WebURL);
        }
        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            try
            {
                Site = await client.GetSharepointSiteFromUrl(token, WebUrlValue);
            }
            catch(Exception e)
            {
                throw new Exception("Error Occured While Retrieving Site From URL", e);
            }
        }
    }
}
