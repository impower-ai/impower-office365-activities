using Microsoft.Graph;
using System;
using System.Activities;
using System.ComponentModel;
using System.IO;
using System.Linq.Expressions;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Mail
{
    [DisplayName("Download Message as .EML")]
    public class GetMessageAsEML : MessageActivity
    {
        [DisplayName("File Path")]
        [Category("Input")]
        [Description("Where To Save The Email")]
        [DefaultValue(true)]
        public InArgument<string> FilePath { get; set; }
        protected string FilePathValue;

        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            var stream = await client.GetMessageAsEML(token, MessageIdValue, EmailValue);
            using (FileStream outputFileStream = new FileStream(FilePathValue, FileMode.Create))
            {
                stream.CopyTo(outputFileStream);
            }
            return ctx => { Expression.Empty(); };
        }

        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            FilePathValue = FilePath.Get(context);
        }
    }
}
