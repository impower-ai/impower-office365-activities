using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using Newtonsoft.Json;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using Impower.Office365.Sharepoint.Models;
using Impower.Office365.Excel;
using System.Text.RegularExpressions;
using System.Data;
using static Impower.Office365.Excel.ExcelExtensions;

namespace Impower.Office365.Sharepoint
{
    public static class SharepointExtensions
    {
        public static async Task RecalculateSharepointWorkbook(this GraphServiceClient client, CalculationType type, TimeSpan interval, TimeSpan timeout, CancellationToken token, string siteId, string driveId, string itemId)
        {
            var driveItemRequestBuilder = client.Sites[siteId].Drives[driveId].Items[itemId];
            await client.RecalculateWorkbook(driveItemRequestBuilder, type, interval, timeout, token);
        }
        public static string GetDriveUrlNameFromDriveItemWebUrl(string driveItemWebUrl, string siteWebUrl)
        {
            if (!driveItemWebUrl.Contains(siteWebUrl))
            {
                throw new Exception("Could not find Site URL within DriveItem WebURL.");
            }
            return driveItemWebUrl
                .Replace(siteWebUrl, String.Empty)
                .TrimStart('/')
                .Split('/')[0];
        }
        public static string GetSharepointSiteUrlFromDriveItemWebUrl(string driveItemWebUrl)
        {
            return Regex.Match(driveItemWebUrl, ".*/sites/([^/]*(/|$)){1}").Value;
        }
        public static string GetSharepointHostNameFromUrl(string url)
        {
            return url.Replace("http://", String.Empty).Replace("https://", String.Empty).Split('/')[0];
        }
        public static string GetSharepointSitePathFromUrl(string url)
        {
            int index = url.IndexOf("/sites/");
            if (index < 0)
            {
                throw new Exception("Could not find site path from URL");
            }
            return url.Substring(index).TrimEnd('/');
        }
        public static string GetEncodedSharingUrl(string url)
        {
            string base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(url));
            string encodedUrl = "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
            return encodedUrl;
        }
        public static async Task<Permission> ShareDriveItem(
            this GraphServiceClient client,
            CancellationToken token,
            string driveItemId,
            string siteId,
            string driveId,
            LinkType type
        )
        {
            IDriveRequestBuilder drive;
            if (String.IsNullOrWhiteSpace(driveId))
            {
                drive = client.Sites[siteId].Drive;
            }
            else
            {
                drive = client.Sites[siteId].Drives[driveId];
            }
            return await drive.Items[driveItemId].CreateLink(type.ToString(), "organization").Request().PostAsync(token);
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
        public static async Task<Drive> GetDriveFromDriveId(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId
        )
        {
            return await client.Sites[siteId].Drives[driveId].Request().GetAsync(token);
        }
        public static async Task<Site> GetSiteFromSiteId(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId
        )
        {
            return await client.Sites[siteId].Request().GetAsync(token);

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
            CancellationToken token,
            DriveItem driveItem,
            string siteId
        )
        {
            Drive Drive = null;
            //Atttempt to grab DriveId from parent reference.
            string DriveId = driveItem.ParentReference?.DriveId;

            //As a fallback, attempt to get Drive from WebUrl
            if (String.IsNullOrWhiteSpace(DriveId) && !String.IsNullOrWhiteSpace(driveItem.WebUrl))
            {
                Drive = await client.GetSharepointDriveByUrl(token, siteId, driveItem.WebUrl);
                DriveId = Drive?.Id;
            }

            //As an additional fallback, attempt to get Drive by 
            if (String.IsNullOrWhiteSpace(DriveId) && !String.IsNullOrEmpty(driveItem.Id) && !String.IsNullOrEmpty(driveItem.ETag))
            {
                var defaultDrive = await client.GetDefaultDriveForSite(token, siteId);
                try
                {
                    var foundItem = await client.GetSharepointDriveItem(token, siteId, defaultDrive.Id, driveItem.Id);
                    DriveId = foundItem?.ETag == driveItem.ETag ? defaultDrive.Id : null;
                }
                catch
                {
                    //If "GetSharepointDriveItem" fails, that means that the given ItemID was not found in the default drive, so we can safely move on to throwing our error.
                }
            }
            if (String.IsNullOrWhiteSpace(DriveId))
            {
                throw new Exception("DriveItem provided did not have enough information to determine the drive.");
            }

            if (Drive == null)
            {
                Drive = await client.GetDriveFromDriveId(token, siteId, DriveId);
            }
            return Drive;
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
            if (String.IsNullOrWhiteSpace(siteId) && !String.IsNullOrWhiteSpace(driveItem.WebUrl))
            {
                var siteUrl = GetSharepointSiteUrlFromDriveItemWebUrl(driveItem.WebUrl);
                site = await client.GetSharepointSiteFromUrl(token, siteUrl);
                siteId = site.Id;
            }
            //At this point, if siteId is not set, we can conclude the above methods have failed.
            if (String.IsNullOrWhiteSpace(siteId))
            {
                throw new Exception("DriveItem provided did not have enough information to determine the site.");
            }
            if (site == null)
            {
                site = await client.GetSiteFromSiteId(token, siteId);
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
            string siteId,
            string siteWebUrl,
            string driveItemWebUrl
        )
        {
            var driveUrlName = GetDriveUrlNameFromDriveItemWebUrl(driveItemWebUrl, siteWebUrl);
            var reconstructedDriveWebUrl = $"{siteWebUrl.TrimEnd('/')}/{driveUrlName}";
            return await client.GetSharepointDriveByUrl(token, siteId, reconstructedDriveWebUrl);

        }
        public static async Task<Drive> GetSharepointDriveByUrl(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveUrl
        )
        {
            var site = await client.Sites[siteId].Request().GetAsync(token);
            var allDrives = await client.Sites[siteId].Drives.Request().GetAsync(token);
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
            string siteId,
            string driveId
        )
        {
            var drive = await client.Sites[siteId].Drives[driveId].Request().Expand(d => d.List).GetAsync(token);
            return drive.List;
        }
        public static async Task<Drive> GetSharepointDriveByName(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveName
        )
        {
            var site = await client.Sites[siteId].Request().GetAsync(token);
            var allDrives = await client.Sites[siteId].Drives.Request().Expand(drive => drive.List).GetAsync(token);
            var matchingDrives = allDrives.Where(drive => drive.Name == driveName);
            if (matchingDrives.Any())
            {
                return matchingDrives.First();
            }
            return null;
        }
        internal static IDriveItemRequest GetSharepointDriveItemRequest(
            this GraphServiceClient client,
           string siteId,
            string driveId,
            string itemId
        )
        {
            IDriveRequestBuilder drive;
            if (String.IsNullOrWhiteSpace(driveId))
            {
                drive = client.Sites[siteId].Drive;
            }
            else
            {
                drive = client.Sites[siteId].Drives[driveId];
            }
            return drive.Items[itemId].Request();
        }
        public static async Task<DriveItem> GetSharepointDriveItem(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId,
            string itemId
        )
        {
            return await client.GetSharepointDriveItemRequest(siteId, driveId, itemId).Expand(item => item.ListItem).GetAsync(token);
        }
        public static async Task<DriveItem> GetSharepointWorkbook(
                        this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId,
            string itemId
        )
        {
            var driveItem = await client.GetSharepointDriveItemRequest(siteId, driveId, itemId).Expand(item => item.Workbook).GetAsync(token);
            if (String.IsNullOrWhiteSpace(driveItem?.Workbook?.Application?.Id ?? String.Empty))
            {
                throw new Exception("This drive item does not appear to be a valid workbook.");
            }
            return driveItem;
        }
        public static async Task UploadListItems(this GraphServiceClient client, CancellationToken token, string siteId, string listId, DataTable data)
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
                _ = await client.Sites[siteId].Lists[listId].Items.Request().AddAsync(listItem, cancellationToken: token);
            }
        }
        public static async Task<List> GetSharepointList(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string listId
        )
        {
            return await client.Sites[siteId].Lists[listId].Request().Expand(list => list.Columns).GetAsync(token);

        }
        public static async Task<FieldValueSet> UpdateSharepointDriveItemFields(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId,
            string itemId,
            FieldValueSet fieldValueSet
        )
        {
            IDriveRequestBuilder drive;
            if (String.IsNullOrWhiteSpace(driveId))
            {
                drive = client.Sites[siteId].Drive;
            }
            else
            {
                drive = client.Sites[siteId].Drives[driveId];
            }

            return await drive.Items[itemId].ListItem.Fields.Request().UpdateAsync(fieldValueSet, token);

        }
        public static async Task<ListItem> GetSharepointListItem(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string listId,
            string itemId
        )
        {
            return await client.Sites[siteId].Lists[listId].Items[itemId].Request().GetAsync(token);
        }
        public static async Task<List<DriveItem>> GetSharepointDriveItemsByPath(
            this GraphServiceClient client,
            CancellationToken token,
            string siteId,
            string driveId,
            string path
        )
        {
            IDriveItemChildrenCollectionRequest request;
            IDriveRequestBuilder drive;
            string folder = Path.GetDirectoryName(path);
            string filename = Path.GetFileName(path);

            if (String.IsNullOrWhiteSpace(driveId))
            {
                drive = client.Sites[siteId].Drive;
            }
            else
            {
                drive = client.Sites[siteId].Drives[driveId];
            }

            if (String.IsNullOrWhiteSpace(folder))
            {
                request = drive.Root.Children.Request();
            }
            else
            {
                request = drive.Root.ItemWithPath(folder).Children.Request();
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
