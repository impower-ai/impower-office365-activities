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
    [DisplayName("Get Worksheet Name From Table")]
    public class GetSheetNameFromTable : SharepointWorkbookActivity
    {
        [RequiredArgument]
        [DisplayName("Table Name")]
        public InArgument<string> TableNameArgument { get; set; }
        public string TableName;

        [RequiredArgument]
        [DisplayName("Worksheet Name")]
        public OutArgument<string> WorksheetNameArgument { get; set; }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            TableName = TableNameArgument.Get(context);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            var sheetName = await client.GetWorksheetNameFromTable(DriveItemReference, TableName, SessionConfiguration, token);
            return ctx => {
                ctx.SetValue(WorksheetNameArgument, sheetName);
            };
        }
    }
}
