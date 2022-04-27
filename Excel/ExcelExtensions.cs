using Impower.Office365.Sharepoint.Models;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Excel
{
    public static class ExcelExtensions
    {
        private const string SessionHeader = "workbook-session-id";

        public static async Task<WorkbookSessionConfiguration> NewSession(this WorkbookSessionConfiguration current, GraphServiceClient client, DriveItemReference driveItem, CancellationToken token)
        {
            if (!current.UseSession)
            {
                return WorkbookSessionConfiguration.CreateSessionlessConfiguration();
            }
            if(current.Session != null)
            {
                try
                {
                    await client.EndWorkbookSession(driveItem, current, token);
                }
                catch { Trace.WriteLine("Failed to end session - " + current.Session.Id);  }
            }
            return new WorkbookSessionConfiguration(
                await client.CreateSharepointWorkbookSession(driveItem, current.PersistChanges, token),
                current.UseSession,
                current.PersistChanges
            );
        }
        public static async Task<WorkbookSessionInfo> CreateSharepointWorkbookSession(this GraphServiceClient client, DriveItemReference driveItem, bool persistChanges, CancellationToken token)
        {
            return await driveItem.RequestBuilder(client).BeginWorkbookSession(persistChanges, token);
        }
        public static async Task RecalculateSharepointWorkbook(this GraphServiceClient client, CalculationType type, DriveItemReference driveItem, WorkbookSessionConfiguration session, TimeSpan interval, TimeSpan timeout, CancellationToken token)
        {
            await client.RecalculateWorkbook(driveItem.RequestBuilder(client), session, type, interval, timeout, token);
        }

        public static TRequest UpdateRequestWithSession<TRequest>(this TRequest requestBuilder, WorkbookSessionConfiguration config)
            where TRequest : IBaseRequest
        {
            return config.UseSession ? requestBuilder.Header(SessionHeader, config.Session.Id) : requestBuilder;
        }
        public static TRequest UpdateRequestWithSession<TRequest>(this TRequest requestBuilder, WorkbookSessionInfo session)
            where TRequest : IBaseRequest
        {
            return requestBuilder.Header(SessionHeader, session.Id);
        }
        public enum CalculationType
        {
            Recalculate,
            Full,
            FullRebuild
        }
        public static async Task ExecuteLongRunningRequest(this GraphServiceClient client, IBaseRequest requestBuilder, HttpMethod method, TimeSpan interval, TimeSpan timeout, CancellationToken token)
        {
            var pollingUri = await client.GetLongRunningRequestPollingUri(requestBuilder, method, token);
            await client.PollAsyncOperation(pollingUri, interval, token, timeout);
        }
        public static async Task<TResult> ExecuteLongRunningRequest<TResult>(this GraphServiceClient client, IBaseRequest requestBuilder,  HttpMethod method, TimeSpan interval, TimeSpan timeout, CancellationToken token)
            where TResult : class
        {
            var pollingUri = await client.GetLongRunningRequestPollingUri(requestBuilder, method, token);
            var resultUri = await client.PollAsyncOperation(pollingUri, interval, token, timeout);
            var result = await GetCompletedOperationResult<TResult>(client, resultUri, token);
            return result;

        }
        internal static async Task<Uri> GetLongRunningRequestPollingUri(this GraphServiceClient client, IBaseRequest requestBuilder, HttpMethod method, CancellationToken token)
        {
            var initialRequest = requestBuilder.Header("Prefer", "respond-async").GetHttpRequestMessage();
            initialRequest.Method = method;
            var initialResponse = await client.HttpProvider.SendAsync(initialRequest, HttpCompletionOption.ResponseHeadersRead, token);
            initialResponse.EnsureSuccessStatusCode();
            return initialResponse.Headers.Location;
        }
        internal static async Task<TResult> GetCompletedOperationResult<TResult>(this GraphServiceClient client, Uri uri, CancellationToken token)
            where TResult : class
        {
            HttpRequestMessage finalRequest = new HttpRequestMessage(HttpMethod.Get, uri);
            ResponseHandler handler = new ResponseHandler(client.HttpProvider.Serializer);

            var response = await client.HttpProvider.SendAsync(finalRequest, HttpCompletionOption.ResponseContentRead, token);
            response.EnsureSuccessStatusCode();
            return await handler.HandleResponse<TResult>(response);
        }
        internal static async Task<Uri> PollAsyncOperation(this GraphServiceClient client, Uri uri, TimeSpan interval, CancellationToken token, TimeSpan timeout)
        {
            var pollOperationTask = client.PollAsyncOperation(uri, interval, token);
            if (await Task.WhenAny(pollOperationTask, Task.Delay(timeout, token)) == pollOperationTask)
            {
                return await pollOperationTask;

            }
            else
            {
                throw new TimeoutException("Timeout Reached While Polling");
            }
        }
        internal static async Task<Uri> PollAsyncOperation(this GraphServiceClient client, Uri uri, TimeSpan interval, CancellationToken token)
        {
            var resourceLocationKey = "resourceLocation";
            string status;
            JObject responseObject;
            do
            {
                await Task.Delay(interval);
                HttpRequestMessage message = new HttpRequestMessage(HttpMethod.Get, uri);
                var response = await client.HttpProvider.SendAsync(message, HttpCompletionOption.ResponseContentRead, token);
                response.EnsureSuccessStatusCode();
                var responseString = await response.Content.ReadAsStringAsync();
                responseObject = JObject.Parse(responseString);
                status = responseObject["status"].ToString();
            } while (!status.Equals("succeeded"));
            if (responseObject.ContainsKey(resourceLocationKey))
            {
                return new Uri(responseObject[resourceLocationKey].ToString());
            }
            return null;
        }
        public static async Task<WorkbookSessionConfiguration> BeginSharepointWorkbookSession(this GraphServiceClient client, DriveItemReference driveItem, bool persistChanges, CancellationToken token)
        {
            return await BeginWorkbookSession(driveItem.RequestBuilder(client), persistChanges, token);
        }
        public static async Task RefreshWorkbookSession(this GraphServiceClient client, DriveItemReference driveItem, WorkbookSessionInfo session, CancellationToken token)
        {
            await RefreshWorkbookSession(driveItem.RequestBuilder(client), session, token);
        }
        public static async Task EndWorkbookSession(this GraphServiceClient client, DriveItemReference driveItem, WorkbookSessionConfiguration session, CancellationToken token)
        {
            await EndWorkbookSession(driveItem.RequestBuilder(client), session, token);
        }

        internal static async Task<WorkbookSessionConfiguration> BeginWorkbookSession(this IDriveItemRequestBuilder driveItemRequestBuilder, bool persistChanges, CancellationToken token)
        {
            var sessionInfo = await driveItemRequestBuilder.Workbook.CreateSession(persistChanges).Request().PostAsync(token);
            return new WorkbookSessionConfiguration(sessionInfo, true, true);
        }
        internal static async Task RefreshWorkbookSession(this IDriveItemRequestBuilder driveItemRequestBuilder, WorkbookSessionInfo session, CancellationToken token)
        {
            await driveItemRequestBuilder.Workbook.RefreshSession().Request().UpdateRequestWithSession(session).PostAsync(token);
        }
        internal static async Task EndWorkbookSession(this IDriveItemRequestBuilder driveItemRequestBuilder, WorkbookSessionInfo session, CancellationToken token)
        {
            await driveItemRequestBuilder.Workbook.CloseSession().Request().UpdateRequestWithSession(session).PostAsync(token);
        }
        internal static async Task RecalculateWorkbook(this GraphServiceClient client, IDriveItemRequestBuilder driveItemRequestBuilder, WorkbookSessionConfiguration session, CalculationType type, TimeSpan pollInterval, TimeSpan timeout, CancellationToken token)
        {
            Console.WriteLine("Polling status....");
            Trace.WriteLine("Polling status...");
            var request = driveItemRequestBuilder.Workbook.Application.Calculate(type.ToString()).Request().UpdateRequestWithSession(session);
            await client.ExecuteLongRunningRequest(request, HttpMethod.Post, pollInterval, timeout, token);
        }
        
    }
}
