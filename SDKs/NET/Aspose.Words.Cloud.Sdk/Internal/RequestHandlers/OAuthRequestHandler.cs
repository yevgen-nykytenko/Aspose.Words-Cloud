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
    using System.IO;
    using System.Net;    
    using System.Text;

    using Newtonsoft.Json;

    internal class OAuthRequestHandler : IRequestHandler
    {        
        private readonly Configuration configuration;
        private string accessToken;
        private string refreshToken;

        public OAuthRequestHandler(Configuration configuration)
        {
            this.configuration = configuration;
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
            var request = WebRequest.Create(this.configuration.ApiBaseUrl + "/oauth2/token");
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";

            var postData = "grant_type=client_credentials";
            postData += "&client_id=" + this.configuration.AppSid;
            postData += "&client_secret=" + this.configuration.AppKey;

            var data = Encoding.ASCII.GetBytes(postData);            
            request.ContentLength = data.Length;

            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            var response = (HttpWebResponse)request.GetResponse();

            var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

            var result = (GetAccessTokenResult)SerializationHelper.Deserialize(responseString, typeof(GetAccessTokenResult));

            this.accessToken = result.AccessToken;
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