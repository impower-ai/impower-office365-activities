﻿using Microsoft.Graph;
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
    public class GetMessage : Office365Activity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("MessageID of the target message")]
        public InArgument<string> MessageID { get; set; }
        [Category("Input")]
        [Description("Email address of user the email is associated with")]
        [RequiredArgument]
        public InArgument<string> Email { get; set; }
        [Category("Misc")]
        [Description("Retrieve Attachments? Set to 'False' for performance use-cases.")]
        [DefaultValue(true)]
        public InArgument<bool> GetAttachments { get; set; }
        [Category("Output")]
        [Description("Message object retrieved")]
        public OutArgument<Message> Message { get; set; }


        internal string MessageIdValue;
        internal string EmailValue;
        internal bool GetAttachmentsValue;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            GetAttachmentsValue = context.GetValue(GetAttachments);
            MessageIdValue = context.GetValue(MessageID);
            EmailValue = context.GetValue(Email);
        }
        protected override Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token){ return Task.CompletedTask; }
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
