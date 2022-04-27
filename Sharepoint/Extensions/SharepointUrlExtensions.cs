using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Impower.Office365.Sharepoint
{
    public static partial class SharepointExtensions
    {
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
            string base64Value = Convert.ToBase64String(Encoding.UTF8.GetBytes(url));
            string encodedUrl = "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
            return encodedUrl;
        }
    }
}
