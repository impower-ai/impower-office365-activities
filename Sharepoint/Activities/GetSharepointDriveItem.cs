using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
namespace Impower.Office365.Sharepoint
{
    [DisplayName("Get DriveItem")]
    public class GetSharepointDriveItem : SharepointDriveItemActivity
    {
        [Category("Output")]
        public OutArgument<ListItem> ListItemOutput { get; set; }
        [Category("Output")]
        public OutArgument<DriveItem> DriveItemOutput { get; set; }

        [Category("Output")]
        public OutArgument<Dictionary<string,object>> FieldsOutput { get; set; }
        [Category("Output")]
        public OutArgument<ItemReference> ReferenceOutput { get; set; }
        protected override Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(
          CancellationToken token,
          GraphServiceClient client
        )
        {
            return Task.FromResult<Action<AsyncCodeActivityContext>>(ctx =>
            {
                FieldsOutput.Set(ctx, new Dictionary<string, object>());
                DriveItemOutput.Set(ctx, DriveItem);
                if (DriveItem.ParentReference != null)
                {
                    ReferenceOutput.Set(ctx, DriveItem.ParentReference);
                }
                ListItemOutput.Set(ctx, DriveItem.ListItem);
                if (DriveItem.ListItem.AdditionalData != null)
                {
                    FieldsOutput.Set(ctx, DriveItem.ListItem.Fields.AdditionalData);
                }
            });
        }
    }
}
