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
    public class GetDriveItemFromSharingLink : Office365Activity
    {
        [RequiredArgument]
        [Category("Input")]
        [DisplayName("Sharing URL")]
        public InArgument<string> SharingURL { get; set; }
        internal string sharingUrl;

        [Category("Output")]
        public OutArgument<DriveItem> DriveItem { get; set; }
        [Category("Output")]
        public OutArgument<ListItem> ListItem { get; set; }
        [Category("Output")]
        public OutArgument<ItemReference> Parent { get; set; }
        //[Category("Output")]
        //public OutArgument<SharedDriveItem> SharedDriveItem { get; set; }
        //[Category("Output")]
        //public OutArgument<Site> Site { get; set; }
        protected override Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            return Task.CompletedTask;
        }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            sharingUrl = context.GetValue(SharingURL);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            Task<DriveItem> driveItemTask = client.GetDriveItemFromSharingUrl(token, sharingUrl);
            Task<ListItem> listItemTask  = client.GetListItemFromSharingUrl(token, sharingUrl);
            await Task.WhenAll(driveItemTask, listItemTask);
            var driveItem = driveItemTask.Result;
            var listItem = listItemTask.Result;

            return ctx =>
            {
                ctx.SetValue(DriveItem, driveItem);
                ctx.SetValue(ListItem, listItem);
                if(driveItem.ParentReference != null)
                {
                    ctx.SetValue(Parent, driveItem.ParentReference);
                }
            };

        }
    }
}
