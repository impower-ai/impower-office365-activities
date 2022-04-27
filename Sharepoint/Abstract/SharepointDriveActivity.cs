using Microsoft.Graph;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using static Impower.Office365.Sharepoint.SharepointExtensions;
using Impower.Office365.Sharepoint.Models;

namespace Impower.Office365.Sharepoint
{
    public abstract class SharepointDriveActivity : SharepointSiteActivity
    {
        [Category("Connection")]
        [DisplayName("Sharepoint Drive")]
        [Description("The Target Drive Name. Defaults To The Documents Library")]
        public InArgument<string> DriveName { get; set; }
        protected string DriveNameValue;
        protected Drive Drive;
        protected DriveReference DriveReference => SiteReference.Drive(Drive);
        protected ListReference DefaultList => SiteReference.List(Drive.List);
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            DriveNameValue = context.GetValue(DriveName);
        }
        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            await base.Initialize(client, context, token);
            if (!String.IsNullOrWhiteSpace(DriveNameValue))
            {
                Drive = await client.GetSharepointDriveByName(token, SiteReference, DriveNameValue);
                if(Drive == null)
                {
                    throw new Exception("Error Occured While Retrieving Drive By Name");
                }
            }
        }
    }
}
