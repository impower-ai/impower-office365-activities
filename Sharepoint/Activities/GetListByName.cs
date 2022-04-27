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
    public class GetListByName : SharepointSiteActivity
    {
        [RequiredArgument]
        [DisplayName("List Name")]
        [Category("Input")]
        public InArgument<string> ListName { get; set; }
        protected string ListNameValue;
        public OutArgument<List> ListOutput { get; set; }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            ListNameValue = ListName.Get(context);

        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            var matchingLists = await SiteReference.RequestBuilder(client).Lists.Request().Filter($"contains(Name,'{ListName}").GetAsync(token);
            if (matchingLists.Any())
            {
                return ctx =>
                {
                    ctx.SetValue(ListOutput, matchingLists.First());
                };
            }
            return ctx => { };

        }
    }
}
