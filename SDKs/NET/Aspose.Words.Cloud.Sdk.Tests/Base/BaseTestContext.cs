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

    /// <summary>
    /// Base class for all tests
    /// </summary>
    public abstract class BaseTestContext
    {
        private const string BaseProductUri = @"http://api-dev.aspose.cloud/v1.1";
        private const string AppSID = "78b637f6-b4cc-41de-a619-d8bd9fc2b6b6";
        private const string AppKey = "3d588eb82b3d5a634ad3141f09b03629";
        private const string DropBoxAppSid = "C821FA88-DA8B-4B97-925A-8D69A6B2FCD1";
        private const string DropBoxAppKey = "9be9d89b967e5f08d4bbddfa8f7cbcd0";

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseTestContext"/> class.
        /// </summary>
        protected BaseTestContext()
        {
            this.WordsApi = new WordsApi(AppKey, AppSID, BaseProductUri);
            this.StorageApi = new StorageApi(AppKey, AppSID, BaseProductUri);
            this.DropboxStorageApi = new StorageApi(DropBoxAppKey, DropBoxAppSid, BaseProductUri);
        }

        /// <summary>
        /// Base path to test data
        /// </summary>
        protected static string BaseTestDataPath
        {
            get
            {
                return "Temp/SdkTests/TestData";
            }
        }

        /// <summary>
        /// Base path to output data
        /// </summary>
        protected static string BaseTestOutPath
        {
            get
            {
                return "TestOut";
            }
        }

        /// <summary>
        /// Storage API
        /// </summary>
        protected StorageApi StorageApi { get; }

        /// <summary>
        /// Dropbox storage API
        /// </summary>
        protected StorageApi DropboxStorageApi { get; set; }

        /// <summary>
        /// Words API
        /// </summary>
        protected WordsApi WordsApi { get; }
    }
}