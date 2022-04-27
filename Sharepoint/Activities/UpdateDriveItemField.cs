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

namespace Impower.Office365.Sharepoint
{
    [DisplayName("Update DriveItem Field")]
    public class UpdateDriveItemField : SharepointDriveItemActivity
    {
        [DisplayName("Field Name")]
        [RequiredArgument]
        [Category("Input")]
        public InArgument<string> FieldNameInput { get; set; }
        [DisplayName("Value")]
        [RequiredArgument]
        [Category("Input")]
        public InArgument<object> FieldInput { get; set; }
        protected string FieldName;
        protected object Field;
        protected Dictionary<string, object> fieldData = new Dictionary<string, object>();
        [DisplayName("Updated Fields")]
        [Category("Output")]
        public OutArgument<Dictionary<string,object>> UpdatedFields { get; set; }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            FieldName = FieldNameInput.Get(context);
            Field = FieldInput.Get(context);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            var driveItemReference = new DriveItemReference(Site, Drive, DriveItem.Id);
            var listId = DriveItem?.ListItem?.ParentReference?.Id ?? (await driveItemReference.Get(client, token)).ListItem.ParentReference.Id;
            var list = await SiteReference.List(listId).Get(client,token);

            //TODO - this logic is messy - potential collisions of internal names and display names could lead to unexpected behavior.
            var writeableColumns = list.Columns.Where(column => !(column.ReadOnly ?? false));
            var matchingColumns = writeableColumns.Where(column => column.Name.Equals(FieldName) || column.DisplayName.Equals(FieldName));
            if (matchingColumns.Any())
            {
                fieldData[matchingColumns.First().Name] = Field;
            }
            else
            {
                throw new Exception($"Could not find a field matching '{FieldName}' in the target list. Available fields are: {String.Join(",", writeableColumns.Select(column => column.Name))}");
            }
            var updatedFields = await client.UpdateSharepointDriveItemFields(token, driveItemReference, new FieldValueSet { AdditionalData = fieldData });
            return ctx =>
            {
                ctx.SetValue(UpdatedFields, updatedFields.AdditionalData);
            };
        }
    }
}
