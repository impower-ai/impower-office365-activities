using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Sharepoint
{
    public class GetListByName : SharepointSiteActivity
    {
        public InArgument<string> ListName { get; set; }
        private string listName;
        public OutArgument<List> List { get; set; }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            listName = ListName.Get(context);

        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            var matchingLists = await client.Sites[this.SiteId].Lists.Request().Filter($"contains(Name,'{listName}").GetAsync(token);
            if (matchingLists.Any())
            {
                return ctx =>
                {
                    ctx.SetValue(List, matchingLists.First());
                };
            }
            return null;

        }
    }
}
