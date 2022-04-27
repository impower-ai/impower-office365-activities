using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Impower.Office365;
using System.ComponentModel;

namespace Impower.Office365.Mail
{
    [DisplayName("Get Message By ID")]
    public class GetMessage : MessageActivity
    {
        [DisplayName("Get Attachments?")]
        [Category("Config")]
        [Description("Retrieve Attachments? Set to 'False' for performance use-cases.")]
        [DefaultValue(true)]
        public InArgument<bool> GetAttachments { get; set; }
        [Category("Output")]
        [Description("Message object retrieved")]
        public OutArgument<Message> Message { get; set; }


        internal bool GetAttachmentsValue;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            GetAttachmentsValue = context.GetValue(GetAttachments);
            
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(
          CancellationToken cancellationToken,
          GraphServiceClient client
        )
        {
            var message = await client.GetMessage(cancellationToken, MessageIdValue, EmailValue);
            return ctx =>
            {
                Message.Set(ctx, message);
            };
        }
    }
}
