using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using Impower.Office365.Sharepoint.Models;
using Impower.Office365.Excel;
using System.Data;
using static Impower.Office365.Excel.ExcelExtensions;

namespace Impower.Office365.Sharepoint
{
    public static partial class SharepointExtensions
    {

        public static async Task<Permission> CreateSharingLinkForSharepointDriveItem(
            this GraphServiceClient client,
            CancellationToken token,
            DriveItemReference driveItem,
            LinkType type
        )
        {
            return await driveItem.RequestBuilder(client).CreateLink(type.ToString(), "organization").Request().PostAsync(token);
        }
        public static async Task<DriveItem> GetDriveItemFromSharingUrl(
            this GraphServiceClient client,
            CancellationToken token,
            string sharingUrl
        )
        {
            var encodedUrl = GetEncodedSharingUrl(sharingUrl);
            return await client.Shares[encodedUrl].DriveItem.Request().GetAsync(token);

        }
        public static async Task<ListItem> GetListItemFromSharingUrl(
            this GraphServiceClient client,
            CancellationToken token,
            string sharingUrl
        )
        {
            var encodedUrl = GetEncodedSharingUrl(sharingUrl);
            return await client.Shares[encodedUrl].ListItem.Request().GetAsync(token);

        }
        public static async Task<Site> GetSharepointSiteFromUrl(
            this GraphServiceClient client,
            CancellationToken token,
            string webUrl
        )
        {
            var hostName = GetSharepointHostNameFromUrl(webUrl);
            var sitePath = GetSharepointSitePathFromUrl(webUrl);
            try
            {
                return await client.Sites.GetByPath(sitePath, hostName).Request().GetAsync(token);
            }
            catch (Exception e)
            {
                throw new Exception($"Could not find a site for '{webUrl}'", e);
            }
        }

        public static async Task<Drive> AttemptToRetrieveDriveFromDriveItem(
            this GraphServiceClient client,
            DriveItem driveItem,
            SiteReference site,
            CancellationToken token
        )
        {
            Drive Drive = null;
            //Atttempt to grab DriveId from parent reference.
            string DriveId = driveItem.ParentReference?.DriveId;

            //As a fallback, attempt to get Drive from WebUrl
            if (String.IsNullOrWhiteSpace(DriveId) && !String.IsNullOrWhiteSpace(driveItem.WebUrl))
            {
                Drive = await client.GetSharepointDriveByUrl(token, site, driveItem.WebUrl);
                DriveId = Drive?.Id;
            }

            if (String.IsNullOrWhiteSpace(DriveId) && !string.IsNullOrEmpty(driveItem.Id) && !String.IsNullOrEmpty(driveItem.ETag))
            {
                var defaultDrive = await client.GetDefaultDriveForSite(token, site.SiteId);
                try
                {
                    var foundItem = await new DriveItemReference(site, defaultDrive, driveItem).Get(client, token);
                    DriveId = foundItem?.ETag == driveItem.ETag ? defaultDrive.Id : null;
                }
                catch
                {
                    //If "GetSharepointDriveItem" fails, that means that the given ItemID was not found in the default drive, so we can safely move on to throwing our error.
                }
            }
            if (string.IsNullOrWhiteSpace(DriveId))
            {
                throw new Exception("DriveItem provided did not have enough information to determine the drive.");
            }

            return Drive ?? await new DriveReference(site, DriveId).Get(client, token);
        }
        public static async Task<Site> AttemptToRetreiveSiteFromDriveItem(
            this GraphServiceClient client,
            CancellationToken token,
            DriveItem driveItem
        )
        {
            //Attempts to acquire from parent reference
            string siteId = driveItem.ParentReference?.SiteId;
            Site site = null;
            //As a fallback, attempts to acquire from url.
            if (string.IsNullOrWhiteSpace(siteId) && !String.IsNullOrWhiteSpace(driveItem.WebUrl))
            {
                var siteUrl = GetSharepointSiteUrlFromDriveItemWebUrl(driveItem.WebUrl);
                site = await client.GetSharepointSiteFromUrl(token, siteUrl);
                siteId = site.Id;
            }
            //At this point, if siteId is not set, we can conclude the above methods have failed.
            if (string.IsNullOrWhiteSpace(siteId))
            {
                throw new Exception("DriveItem provided did not have enough information to determine the site.");
            }
            if (site == null)
            {
                site = await new SiteReference(siteId).Get(client, token);
            }
            return site;
        }
        public static async Task<Drive> GetDefaultDriveForSite(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId
        )
        {
            return await client.Sites[siteId].Drive.Request().GetAsync(token);
        }
        public static async Task<Drive> GetSharepointDriveFromDriveItemWebUrl(
            this GraphServiceClient client,
            CancellationToken token,
            SiteReference site,
            string siteWebUrl,
            string driveItemWebUrl
        )
        {
            var driveUrlName = GetDriveUrlNameFromDriveItemWebUrl(driveItemWebUrl, siteWebUrl);
            var reconstructedDriveWebUrl = $"{siteWebUrl.TrimEnd('/')}/{driveUrlName}";
            return await client.GetSharepointDriveByUrl(token, site, reconstructedDriveWebUrl);

        }
        public static async Task<Drive> GetSharepointDriveByUrl(
            this GraphServiceClient client,
            CancellationToken token,
            SiteReference site,
            string driveUrl
        )
        {
            var allDrives = await site.RequestBuilder(client).Drives.Request().GetAsync(token);
            var matchingDrives = allDrives.Where(drive => drive.WebUrl == driveUrl);
            if (matchingDrives.Any())
            {
                return matchingDrives.First();
            }
            return null;
        }

        public static async Task<List> GetSharepointListFromDrive(
            this GraphServiceClient client,
            CancellationToken token,
            DriveReference drive
        )
        {
            var retrievedDrive = await drive.Get(client, token);
            return retrievedDrive.List;
        }
        public static async Task<Drive> GetSharepointDriveByName(
            this GraphServiceClient client,
            CancellationToken token,
            SiteReference site,
            string driveName
        )
        {
            var allDrives = await site.RequestBuilder(client).Drives.Request().GetAsync(token);
            var matchingDrives = allDrives.Where(drive => drive.Name == driveName);
            if (matchingDrives.Any())
            {
                return matchingDrives.First();
            }
            return null;
        }
        public static async Task<DriveItem> GetSharepointWorkbook(
            this GraphServiceClient client,
            CancellationToken token,
            DriveItemReference item,
            WorkbookSessionConfiguration session

        )
        {
            var driveItem = await item.RequestBuilder(client).Request().UpdateRequestWithSession(session).GetAsync(token);
            if (String.IsNullOrWhiteSpace(driveItem?.Workbook?.Application?.Id ?? String.Empty))
            {
                throw new Exception("This drive item does not appear to be a valid workbook.");
            }
            return driveItem;
        }
        public static async Task UploadListItems(this GraphServiceClient client, CancellationToken token, ListReference list, DataTable data)
        {

            foreach (DataRow row in data.AsEnumerable())
            {

                var listItem = new ListItem()
                {
                    Fields = new FieldValueSet()
                    {

                        AdditionalData = new Dictionary<string, object>()
                        {

                        }
                    }
                };
                foreach (DataColumn column in data.Columns)
                {
                    listItem.Fields.AdditionalData[column.ColumnName] = row[column.ColumnName].ToString();

                }
                _ = await list.RequestBuilder(client).Items.Request().AddAsync(listItem, cancellationToken: token);
            }
        }
        public static async Task<List> GetSharepointList(
            this GraphServiceClient client,
            CancellationToken token,
            ListReference list
        )
        {
            return await list.RequestBuilder(client).Request().Expand(l => l.Columns).GetAsync(token);

        }
        public static async Task<FieldValueSet> UpdateSharepointDriveItemFields(
            this GraphServiceClient client,
            CancellationToken token,
            DriveItemReference driveItem,
            FieldValueSet fieldValueSet
        )
        {
            return await driveItem.RequestBuilder(client).ListItem.Fields.Request().UpdateAsync(fieldValueSet, token);
        }
        public static async Task<List<DriveItem>> GetSharepointDriveItemsByPath(
            this GraphServiceClient client,
            CancellationToken token,
            DriveReference drive,
            string path
        )
        {
            IDriveItemChildrenCollectionRequest request;
            string folder = Path.GetDirectoryName(path);
            string filename = Path.GetFileName(path);

            if (String.IsNullOrWhiteSpace(folder))
            {
                request = drive.RequestBuilder(client).Root.Children.Request();
            }
            else
            {
                request = drive.RequestBuilder(client).Root.ItemWithPath(folder).Children.Request();
            }

            var items = await request.GetAsync(token);
            if (String.IsNullOrWhiteSpace(filename))
            {
                return items.ToList();
            }
            else
            {
                return items.Where(item => item.Name == filename).ToList();
            }

        }
    }
}
