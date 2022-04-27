using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Linq;

namespace Impower.Office365.Sharepoint
{
    [DisplayName("Update DriveItem Fields")]
    public class UpdateDriveItemFields : SharepointDriveItemActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<Dictionary<string, object>> FieldsInput { get; set; }
        [Category("Input")]
        [DisplayName("Use Display Names")]
        [Description("Allows referencing columns by their display name. If set, keys will be matching first against the internal name and then against the display name, as a fallback.")]
        public InArgument<bool> UseDisplayNames { get; set; }
        [Category("Output")]
        [DisplayName("Updated Fields")]
        public OutArgument<Dictionary<string, object>> UpdatedFields { get; set; }
        protected Dictionary<string, object> FieldsValue;
        protected bool UseDisplayNamesValue; 
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            FieldsValue = context.GetValue(FieldsInput);
            UseDisplayNamesValue = context.GetValue(UseDisplayNames);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            if (UseDisplayNamesValue)
            {
                var list = await SiteReference.List(DriveItem.ListItem.ParentReference.Id).Get(client, token);
                //TODO - this could be cleaned up.
                //This will throw if one of the display names resolves to a name that already exists in the dictionary.
                var newFieldsValue = new Dictionary<string, object>();
                var writeableColumns = list.Columns.Where(column => !(column.ReadOnly ?? false));
                foreach (var kvp in FieldsValue)
                {
                    var matchingColumns = writeableColumns.Where(column => column.Name.Equals(kvp.Key));
                    if(matchingColumns.Any())
                    {
                        newFieldsValue.Add(kvp.Key, kvp.Value);
                        break;
                    }
                    matchingColumns = writeableColumns.Where(column => column.DisplayName.Equals(kvp.Key));
                    if(matchingColumns.Any())
                    {
                        var matchingColumn = matchingColumns.First();
                        newFieldsValue.Add(matchingColumn.Name, kvp.Value);
                        break;
                    }
                    throw new Exception($"Could not find a field matching '{kvp.Key}' in the target list. Available fields are: {String.Join(",", writeableColumns.Select(column => column.Name))}");
                }
                FieldsValue = newFieldsValue;
            }

            FieldValueSet fieldValueSet = new FieldValueSet
            {
                AdditionalData = FieldsValue
            };
            FieldValueSet result = await client.UpdateSharepointDriveItemFields(token, DriveItemReference, fieldValueSet);
            return ctx =>
            {
                ctx.SetValue(UpdatedFields, result.AdditionalData);
            };
        }
    }
}
