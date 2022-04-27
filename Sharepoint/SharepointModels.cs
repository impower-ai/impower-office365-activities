using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Impower.Office365.Sharepoint.Models
{
    public struct DriveItemLocator
    {
        public string Id { get; private set; }
        public DriveItemLocator(string id)
        {
            this.Id = id;
        }
        public static implicit operator string(DriveItemLocator l) => l.Id;
        public static implicit operator DriveItemLocator(string id) => new DriveItemLocator(id);
        public static implicit operator DriveItemLocator(DriveItem driveItem) => new DriveItemLocator(driveItem.Id);
    }
    public struct ListItemLocator
    {
        public string Id { get; private set; }
        public ListItemLocator(string id)
        {
                    this.Id = id;
        }
        public static implicit operator string(ListItemLocator l) => l.Id;
        public static implicit operator ListItemLocator(string id) => new ListItemLocator(id);
        public static implicit operator ListItemLocator(ListItem listItem) => new ListItemLocator(listItem.Id);

    }
    public struct ListLocator
    {
        public string Id { get; private set; }
        public ListLocator(string id)
        {
            this.Id = id;
        }
        public static implicit operator string(ListLocator l) => l.Id;
        public static implicit operator ListLocator(string id) => new ListLocator(id);
        public static implicit operator ListLocator(List list) => new ListLocator(list.Id);
    }
    public enum LinkType
    {
        view,
        edit
    }
}
