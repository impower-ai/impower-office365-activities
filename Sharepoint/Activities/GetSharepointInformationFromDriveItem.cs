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
    [DisplayName("Get Sharepoint Information From Drive Item")]
    public class GetSharepointInformationFromDriveItem : Office365Activity
    {
        [RequiredArgument]
        [Category("Input")]
        [DisplayName("Drive Item")]
        public InArgument<DriveItem> DriveItemInput { get; set; }
        [Category("Output")]
        [DisplayName("Drive")]
        public OutArgument<Drive> DriveOutput { get; set; }
        [Category("Output")]
        [DisplayName("Site")]
        public OutArgument<Site> SiteOutput { get; set; }
        [Category("Output")]
        [DisplayName("Drive Name")]
        public OutArgument<string> DriveName { get; set; }
        [DisplayName("Site URL")]
        [Category("Output")]
        public OutArgument<string> SiteURL { get; set; }
        protected DriveItem DriveItem;
        protected Site SiteValue;
        protected Drive DriveValue;
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            SiteValue = await client.AttemptToRetreiveSiteFromDriveItem(token, DriveItem);
            DriveValue = await client.AttemptToRetrieveDriveFromDriveItem(DriveItem, SiteValue, token);

            return ctx =>
            {
                ctx.SetValue(DriveOutput, DriveValue);
                ctx.SetValue(SiteOutput, SiteValue);
                ctx.SetValue(DriveName, DriveValue.Name);
                ctx.SetValue(SiteURL, SiteValue.WebUrl);
            };
            

        }
        protected override Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            return Task.CompletedTask;
        }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            DriveItem = DriveItemInput.Get(context);
        }
    }
}
