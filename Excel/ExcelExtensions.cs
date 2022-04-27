using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Excel
{
    public static class ExcelExtensions
    {
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
        internal static async Task<WorkbookSessionInfo> BeginWorkbookSession(this IDriveItemRequestBuilder driveItemRequestBuilder)
        {
            return await driveItemRequestBuilder.Workbook.CreateSession(true).Request().PostAsync();
        }
        internal static async Task RecalculateWorkbook(this GraphServiceClient client, IDriveItemRequestBuilder driveItemRequestBuilder, CalculationType type, TimeSpan pollInterval, TimeSpan timeout, CancellationToken token)
        {
            var request = driveItemRequestBuilder.Workbook.Application.Calculate(type.ToString()).Request();
            await client.ExecuteLongRunningRequest(request, HttpMethod.Post, pollInterval, timeout, token);
        }
        
    }
}
