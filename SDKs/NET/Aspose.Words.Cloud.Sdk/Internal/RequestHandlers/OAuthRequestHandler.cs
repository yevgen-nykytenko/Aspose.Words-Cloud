// --------------------------------------------------------------------------------------------------------------------
// <copyright company="Aspose" file="OAuthRequestHandler.cs">
//   Copyright (c) 2016 Aspose.Words for Cloud
// </copyright>
// <summary>
//   Permission is hereby granted, free of charge, to any person obtaining a copy
//  of this software and associated documentation files (the "Software"), to deal
//  in the Software without restriction, including without limitation the rights
//  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//  copies of the Software, and to permit persons to whom the Software is
//  furnished to do so, subject to the following conditions:
// 
//  The above copyright notice and this permission notice shall be included in all
//  copies or substantial portions of the Software.
// 
//  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//  SOFTWARE.
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Aspose.Words.Cloud.Sdk.RequestHandlers
{
    using System.Collections.Generic;
    using System.IO;
    using System.Net;    
    using System.Text;

    using Newtonsoft.Json;

    internal class OAuthRequestHandler : IRequestHandler
    {        
        private readonly Configuration configuration;
        private readonly ApiInvoker apiInvoker;

        private string accessToken; ////= "1n6lfta7NUeAkgfu0JWnEIWdl4QEvECnUF810_9K709ZuMZofe9tneG-_yfTfHAEGuEX0TWk-WIp4tUUuoRBmeubucE_hNpF0zz6p38S73EHfNIMCVZ-drwvVJlDE2nMfX7jOrxDY652xJ5LZYt-41XUr0pV-o_6dXevtmK7xIPeUE1DsbLNIUILNfgebJkce3j6VwtvRQfUtniKVC1CU2ZOZwDEq-ZZr8IIROlJ1uUgX1uxIMCD14UyuX7rycPusGeCEmVK4Yz1nAMc6amfKZl38P071uzsPUCBrHOKY1DoyHJ-q9k7A3M5F75ihl_4AanFrH_7anH0lPlQvJcrnOtiSBEzuoI6TQLuSrpEeEDQ3QHtNZqe6Z6KdNER_6FMHosRDkiX2SiVMA45PtUnuVQyDl2IJBp5sMqs67Ib03XSy60qI";
        private string refreshToken;

        public OAuthRequestHandler(Configuration configuration)
        {
            this.configuration = configuration;

            var requestHandlers = new List<IRequestHandler>();
            requestHandlers.Add(new DebugLogRequestHandler(this.configuration));
            requestHandlers.Add(new ApiExceptionRequestHandler());
            this.apiInvoker = new ApiInvoker(requestHandlers);
        }

        public string ProcessUrl(string url)
        {
            return url;
        }

        public void BeforeSend(WebRequest request, Stream streamToSend)
        {
            if (this.configuration.AuthType != AuthType.OAuth2)
            {
                return;
            }

            if (string.IsNullOrEmpty(this.accessToken))
            {
                this.RequestToken();
            }

            request.Headers.Add("Authorization", "Bearer " + this.accessToken);
        }       

        public void ProcessResponse(HttpWebResponse response, Stream resultStream)
        {            
        }

        private void RequestToken()
        {
            var requestUrl = this.configuration.ApiBaseUrl + "/oauth2/token";

            var postData = "grant_type=client_credentials";
            postData += "&client_id=" + this.configuration.AppSid;
            postData += "&client_secret=" + this.configuration.AppKey;

            var responseString = this.apiInvoker.InvokeApi(
                requestUrl,
                "POST",
                postData,
                contentType: "application/x-www-form-urlencoded");

            var result =
                (GetAccessTokenResult)SerializationHelper.Deserialize(responseString, typeof(GetAccessTokenResult));

            this.accessToken = result.AccessToken;
            this.refreshToken = result.RefreshToken;
        }

        private class GetAccessTokenResult
        {
            [JsonProperty(PropertyName = "access_token")]
            public string AccessToken { get; set; }

            [JsonProperty(PropertyName = "refresh_token")]
            public string RefreshToken { get; set; }
        }        
    }
}