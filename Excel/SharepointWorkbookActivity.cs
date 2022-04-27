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
    public abstract class SharepointWorkbookActivity : SharepointDriveItemActivity
    {
        [Category("Config")]
        [DisplayName("Use Session?")]
        [DefaultValue(true)]
        public InArgument<bool> UseSessionArgument { get; set; }
        [Category("Config")]
        [DisplayName("Persist Changes?")]
        [DefaultValue(true)]
        public InArgument<bool> PersistChangesArgument { get; set; }
        [Category("Config")]
        [DisplayName("Session")]
        public InOutArgument<WorkbookSessionInfo> SessionArgument { get; set; }
        public WorkbookSessionConfiguration SessionConfiguration;
        private Workbook Workbook;
        public bool PersistChanges;
        public bool UseSession;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            var useSession = context.GetValue(UseSessionArgument);
            if (useSession)
            {
                var persistChanges = context.GetValue(PersistChangesArgument);
                var existingSession = context.GetValue(SessionArgument);
                if(existingSession != null && existingSession.PersistChanges.HasValue && existingSession.PersistChanges.Value != persistChanges)
                {
                    throw new ArgumentException("Passed Session Persist Settings Did Not Match Activity Arguments");
                }
                SessionConfiguration = new WorkbookSessionConfiguration(existingSession, useSession, persistChanges);
            }
            else
            {
                SessionConfiguration = WorkbookSessionConfiguration.CreateSessionlessConfiguration();
            }
        }
        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            await base.Initialize(client, context, token);
            if (SessionConfiguration.Session == null)
            {
                SessionConfiguration = await SessionConfiguration.NewSession(client, DriveItemReference, token);
            }
            Workbook = await client.GetSharepointWorkbook(token, DriveItemReference, SessionConfiguration);

        }
        protected override Action<AsyncCodeActivityContext> Finalize()
        {
            var action = base.Finalize();
            if(SessionConfiguration.Session != null)
            {
                action += ctx => ctx.SetValue(SessionArgument, SessionConfiguration.Session);
            }
            return action;
        }
    }
}
