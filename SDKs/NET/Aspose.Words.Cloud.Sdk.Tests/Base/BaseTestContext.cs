// // --------------------------------------------------------------------------------------------------------------------
// // <copyright company="Aspose" file="BaseTestContext.cs">
// //   Copyright (c) 2016 Aspose.Words for Cloud
// // </copyright>
// // <summary>
// //   Permission is hereby granted, free of charge, to any person obtaining a copy
// //  of this software and associated documentation files (the "Software"), to deal
// //  in the Software without restriction, including without limitation the rights
// //  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// //  copies of the Software, and to permit persons to whom the Software is
// //  furnished to do so, subject to the following conditions:
// // 
// //  The above copyright notice and this permission notice shall be included in all
// //  copies or substantial portions of the Software.
// // 
// //  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// //  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// //  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// //  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// //  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// //  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// //  SOFTWARE.
// // </summary>
// // --------------------------------------------------------------------------------------------------------------------

namespace Aspose.Words.Cloud.Sdk.Tests.Base
{
    using Aspose.Words.Cloud.Sdk.Api;

    using Com.Aspose.Storage.Api;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Base class for all tests
    /// </summary>
    [TestClass]
    public abstract class BaseTestContext
    {
        private const string ApiKey = "0fbf678c5ecabdb5caca48452a736dd0";
        private const string ApiSid = "91a2fd07-bba1-4b32-9112-abfb1fe8aebd";
        private const string AppUrl = "http://api.aspose.cloud/v1.1";

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseTestContext"/> class.
        /// </summary>
        protected BaseTestContext()
        {
            this.WordsApi = new WordsApi(ApiKey, ApiSid, AppUrl);
            this.StorageApi = new StorageApi(ApiKey, ApiSid, AppUrl);
        }

        /// <summary>
        /// Storage API
        /// </summary>
        protected StorageApi StorageApi { get; }

        /// <summary>
        /// Words API
        /// </summary>
        protected WordsApi WordsApi { get; }
    }
}