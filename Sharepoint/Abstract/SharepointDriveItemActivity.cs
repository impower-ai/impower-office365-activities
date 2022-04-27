using Impower.Office365.Sharepoint.Models;
using Microsoft.Graph;
using System.Activities;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Sharepoint
{
    public abstract class SharepointDriveItemActivity : SharepointDriveActivity
    {
        [Category("Input")]
        [DisplayName("DriveItem ID")]
        [RequiredArgument]
        public InArgument<DriveItemLocator> DriveItemLocator { get; set; }

        protected DriveItemReference DriveItemReference => DriveReference.Item(DriveItemId);
        protected string DriveItemId;
        protected DriveItem DriveItem;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            DriveItemId = context.GetValue(DriveItemLocator);
        }
        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            await base.Initialize(client, context, token);
            DriveItem = await DriveItemReference.Get(client, token);
            
        }
    }
}
