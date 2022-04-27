using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Sharepoint.Models
{
    public interface ITypedValueProvider
    {
        ITypedValueProvider GetValue<TValue>();
    }
    public abstract class BaseItemReference<TBuilder,TResult> 
        where TResult : BaseItem
        where TBuilder : IBaseItemRequestBuilder

    {
        protected readonly string id;
        public BaseItemReference(string id)
        {
            this.id = id;
        }
        public abstract TBuilder RequestBuilder(GraphServiceClient client);
        public virtual async Task<TResult> Get(GraphServiceClient client, CancellationToken token)
        {
            return (TResult)await RequestBuilder(client).Request().GetAsync(token);
        }
        public virtual async Task<TResult> Update(GraphServiceClient client, CancellationToken token, TResult instance)
        {
            return (TResult)await RequestBuilder(client).Request().UpdateAsync(instance);
        }
    }
    public class SiteReference : BaseItemReference<ISiteRequestBuilder,Site>
    {
        public string SiteId => id;
        public SiteReference(SiteLocator site) : base(site) { }
        public override ISiteRequestBuilder RequestBuilder(GraphServiceClient client) => client.Sites[SiteId];
        public DriveReference Drive(DriveLocator drive) => new DriveReference(this, drive);
        public ListReference List(ListLocator list) => new ListReference(this, list);
        public static implicit operator SiteReference(Site site) => new SiteReference(site);
    }
    public class DriveReference : BaseItemReference<IDriveRequestBuilder, Drive>
    {
        private readonly SiteReference site;
        public string DriveId => id;
        public string SiteId => site.SiteId;
        public DriveReference(SiteLocator site, DriveLocator drive) : base(drive)
        {
            this.site = new SiteReference(site);
        }
        public override IDriveRequestBuilder RequestBuilder(GraphServiceClient client) => String.IsNullOrEmpty(DriveId) ? site.RequestBuilder(client).Drive : site.RequestBuilder(client).Drives[DriveId];

        public override Task<Drive> Get(GraphServiceClient client, CancellationToken token)
        {
            return RequestBuilder(client).Request().Expand(drive => drive.List).GetAsync();
        }
        public DriveItemReference Item(DriveItemLocator driveItem) => new DriveItemReference(this, driveItem);
    }
    public class DriveItemReference : BaseItemReference<IDriveItemRequestBuilder, DriveItem>
    {
        private readonly DriveReference drive;
        public string ItemId => id;
        public string DriveId => drive.DriveId;
        public string SiteId => drive.SiteId;
        public DriveItemReference(DriveReference drive, DriveItemLocator item) : base(item)
        {
            this.drive = drive;
        }
        public DriveItemReference(SiteLocator site, DriveLocator drive, DriveItemLocator item) : this(new DriveReference(site, drive), item) { }
        public override IDriveItemRequestBuilder RequestBuilder(GraphServiceClient client) => drive.RequestBuilder(client).Items[ItemId];
        public override Task<DriveItem> Get(GraphServiceClient client, CancellationToken token)
        {
            return RequestBuilder(client).Request().Expand(item => item.ListItem).GetAsync();
        }
    }
    public class ListReference : BaseItemReference<IListRequestBuilder, List>
    {
        private readonly SiteReference site;
        public string SiteId => site.SiteId;
        public string ListId => id;
        public ListReference(SiteLocator site, ListLocator list) : base(list)
        {
            this.site = new SiteReference(site);
        }
        public override IListRequestBuilder RequestBuilder(GraphServiceClient client) => site.RequestBuilder(client).Lists[ListId];
        public ListItemReference Item(ListItemLocator listItem) => new ListItemReference(SiteId, ListId, listItem);
    }
    public class ListItemReference : BaseItemReference<IListItemRequestBuilder, ListItem>
    {
        private readonly ListReference list;
        public string SiteId => list.SiteId;
        public string ListId => list.ListId;
        public string ItemId => id;
        public ListItemReference(SiteLocator site, ListLocator list, ListItemLocator listItem) : base(listItem)
        {
            this.list = new ListReference(site, list);
        }
        public override IListItemRequestBuilder RequestBuilder(GraphServiceClient client) => list.RequestBuilder(client).Items[ItemId];
    }
    public struct SiteLocator
    {
        public string Id { get; private set; }
        public SiteLocator(string id) => Id = id;
        public static implicit operator string(SiteLocator s) => s.Id;
        public static implicit operator SiteLocator(string id) => new SiteLocator(id);
        public static implicit operator SiteLocator(Site site) => new SiteLocator(site.Id);
        public static implicit operator SiteLocator(SiteReference site) => new SiteLocator(site.SiteId);
    }
    public struct DriveLocator
    {
        public string Id { get; private set; }
        public DriveLocator(string id) => Id = id;
        public static implicit operator string(DriveLocator d) => d.Id;
        public static implicit operator DriveLocator(string id) => new DriveLocator(id);
        public static implicit operator DriveLocator(Drive drive) => new DriveLocator(drive.Id);
        public static implicit operator DriveLocator(DriveReference drive) => new DriveLocator(drive.DriveId);
    }
    public struct DriveItemLocator
    {
        public string Id { get; private set; }
        public DriveItemLocator(string id) => Id = id;
        public static implicit operator string(DriveItemLocator l) => l.Id;
        public static implicit operator DriveItemLocator(string id) => new DriveItemLocator(id);
        public static implicit operator DriveItemLocator(DriveItem driveItem) => new DriveItemLocator(driveItem.Id);
        public static implicit operator DriveItemLocator(DriveItemReference driveItem) => new DriveItemLocator(driveItem.DriveId);
    }
    public struct ListItemLocator
    {
        public string Id { get; private set; }
        public ListItemLocator(string id) => Id = id;
        public static implicit operator string(ListItemLocator l) => l.Id;
        public static implicit operator ListItemLocator(string id) => new ListItemLocator(id);
        public static implicit operator ListItemLocator(ListItem listItem) => new ListItemLocator(listItem.Id);
        public static implicit operator ListItemLocator(ListItemReference listItem) => new ListItemLocator(listItem.ItemId);

    }
    public struct ListLocator
    {
        public string Id { get; private set; }
        public ListLocator(string id) => Id = id;
        public static implicit operator string(ListLocator l) => l.Id;
        public static implicit operator ListLocator(string id) => new ListLocator(id);
        public static implicit operator ListLocator(List list) => new ListLocator(list.Id);
        public static implicit operator ListLocator(ListReference list) => new ListLocator(list.ListId);
    }
    public enum LinkType
    {
        view,
        edit
    }
}
