using Microsoft.Graph;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using static Impower.Office365.Sharepoint.SharepointExtensions;

namespace Impower.Office365.Sharepoint
{
    public abstract class SharepointDriveActivity : SharepointSiteActivity
    {
        [Category("Connection")]
        [DisplayName("Sharepoint Drive")]
        [Description("The Target Drive Name. Defaults To The Documents Library")]
        public InArgument<string> DriveName { get; set; }
        protected string DriveNameValue;
        protected Drive DriveValue;
        protected SharepointDrive Drive => new SharepointDrive()
        protected string DriveId => DriveValue?.Id;
        protected string ListId => DriveValue?.List?.Id;
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
                DriveValue = await client.GetSharepointDriveByName(token, SiteId, DriveNameValue);
                if(DriveValue == null)
                {
                    throw new Exception("Error Occured While Retrieving Drive By Name");
                }
            }
        }
    }
}
