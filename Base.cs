using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Runtime.ExceptionServices;
using System.ComponentModel;
using System.Threading;
using UiPath.Shared.Activities;
using Impower.Office365.Authentication;

namespace Impower.Office365
{
    public abstract class Office365Activity : AsyncTaskCodeActivity
    {
        [Category("Config")]
        [DisplayName("Timeout")]
        [DefaultValue("0:00:30")]
        public virtual InArgument<TimeSpan> Timeout { get; set; }
        public TimeSpan TimeoutValue;
        [Category("Connection")]
        [Description("Specify Client Object, Otherwise Uses Scope.")]
        [DisplayName("Graph Client")]
        public InArgument<GraphServiceClient> GraphClient { get; set; }

        protected abstract void ReadContext(AsyncCodeActivityContext context);
        protected abstract Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token);
        protected virtual Action<AsyncCodeActivityContext> Finalize()
        {
            return ctx => { };
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsync(
          AsyncCodeActivityContext context,
          CancellationToken token)
        {
            //HANDLE CLIENT
            TimeoutValue = context.GetValue(Timeout);
            var client = context.GetValue(GraphClient);
            if(client == null)
            {
                client = (GraphServiceClient)Extensions.GetClientFromScope(context);
            }
            if(client == null)
            {
                throw new Exception("Could not acquire Graph Client from context. Place activity in scope or pass in client directly.");
            }

            //BEGIN EXECUTION
            ReadContext(context);
            await Initialize(client, context, token);
            var actions = await AwaitWithTimeout(ExecuteAsyncWithClient(token, client), TimeoutValue);
            actions += Finalize();
            return actions;

        }

        private async Task<TResult> AwaitWithTimeout<TResult>(Task<TResult> task, TimeSpan timeout)
        {
            if (await Task.WhenAny(task, Task.Delay(timeout)) == task)
            {
                return await task;
            }
            else
            {
                throw new TimeoutException("Timeout Exceeded");
            }
        }
        protected abstract Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(
          CancellationToken token,
          GraphServiceClient client
        );
    }
}
