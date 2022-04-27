using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Impower.Office365.Excel
{
    public struct WorkbookSessionConfiguration
    {
 
        public bool UseSession => useSession;
        private readonly bool useSession;
        public bool PersistChanges => persistChanges;
        private readonly bool persistChanges;
        public WorkbookSessionInfo Session => session;
        private readonly WorkbookSessionInfo session;
        public WorkbookSessionConfiguration(WorkbookSessionInfo session, bool useSession, bool persistChanges)
        {
            this.session = session;
            this.useSession = useSession;
            this.persistChanges = persistChanges;
        }
        public static WorkbookSessionConfiguration CreateSessionlessConfiguration() => new WorkbookSessionConfiguration(null, false, false);
        public static implicit operator WorkbookSessionInfo(WorkbookSessionConfiguration config) => config.Session;
    }
}
