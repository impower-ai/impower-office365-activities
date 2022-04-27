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
        protected string SharingStringValue;

        [Category("Output")]
        public OutArgument<DriveItem> DriveItemOutput { get; set; }
        [Category("Output")]
        public OutArgument<ListItem> ListItemOutput { get; set; }
        [Category("Output")]
        public OutArgument<ItemReference> Parent { get; set; }
        protected override Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            return Task.CompletedTask;
        }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            SharingStringValue = context.GetValue(SharingURL);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            Task<DriveItem> driveItemTask = client.GetDriveItemFromSharingUrl(token, SharingStringValue);
            Task<ListItem> listItemTask  = client.GetListItemFromSharingUrl(token, SharingStringValue);
            await Task.WhenAll(driveItemTask, listItemTask);
            var driveItem = await driveItemTask;
            var listItem = await listItemTask;

            return ctx =>
            {
                ctx.SetValue(DriveItemOutput, driveItem);
                ctx.SetValue(ListItemOutput, listItem);
                if(driveItem.ParentReference != null)
                {
                    ctx.SetValue(Parent, driveItem.ParentReference);
                }
            };

        }
    }
}
