using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Impower.Office365.Authentication
{
    public static class Extensions
    {
        public static GraphServiceClient GetClientFromScope(
            AsyncCodeActivityContext context
        )
        {
            var parentScope = context.DataContext.GetProperties()["ParentScope"];
            if(parentScope != null)
            {
                return parentScope.GetValue(context.DataContext) as GraphServiceClient;
            }
            return null;
        }
    }
}
