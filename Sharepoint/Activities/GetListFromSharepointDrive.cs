﻿using Microsoft.Graph;
using System;
using System.Activities;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
namespace Impower.Office365.Sharepoint
{
    [DisplayName("Get List From Sharepoint Drive")]
    public class GetListFromSharepointDrive : SharepointDriveActivity
    {
        [Category("Output")]
        [DisplayName("List ID")]
        public OutArgument<string> ListIdentifier { get; set; }
        [Category("Output")]
        public OutArgument<List> List { get; set; }
        [Category("Output")]
        [DisplayName("Writable Fields")]
        public OutArgument<string[]> Fields { get; set; }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            List list = await client.GetSharepointList(token, SiteId, ListId);
            string[] fields = list.Columns.Where(column => !(column.ReadOnly ?? false)).Select(column => column.Name).ToArray();
            return ctx =>
            {
                ctx.SetValue(ListIdentifier, ListId);
                ctx.SetValue(List, list);
                ctx.SetValue(Fields, fields);
            };
        }
    }
}
