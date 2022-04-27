using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Impower.Office365;
using Microsoft.Graph;


namespace Impower.Office365.Sharepoint
{
    public class GetListItem : SharepointListItemActivity
    {
        [Category("Output")]
        [DisplayName("List Item")]
        public OutArgument<ListItem> ListItemOutput { get; set; }

        [Category("Output")]
        [DisplayName("List")]
        public OutArgument<List> ListOutput { get; set; }
        protected override Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken cancellationToken, GraphServiceClient client)
        {
            return Task.FromResult<Action<AsyncCodeActivityContext>>(ctx =>
            {
                ctx.SetValue(ListItemOutput, ListItem);
                ctx.SetValue(ListOutput, List);
            });
        }
    }
}
