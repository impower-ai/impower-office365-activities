using Impower.Office365.Sharepoint;
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

namespace Impower.Office365.Excel
{
    public abstract class SharepointWorkbookActivity : SharepointDriveActivity
    {
        [Category("Input")]
        [DisplayName("DriveItem ID")]
        [RequiredArgument]
        public InArgument<DriveItemLocator> DriveItem { get; set; }
        internal string DriveItemId;
        internal DriveItem DriveItemValue;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            DriveItemId = context.GetValue(DriveItem);
        }
        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            await base.Initialize(client, context, token);
            DriveItemValue = await client.GetSharepointWorkbook(token, SiteId, DriveId, DriveItemId);
        }
    }
}
