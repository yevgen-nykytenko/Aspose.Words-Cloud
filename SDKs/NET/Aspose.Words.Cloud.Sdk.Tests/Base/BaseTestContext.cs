// // --------------------------------------------------------------------------------------------------------------------
// // <copyright company="Aspose" file="BaseTestContext.cs">
// //   Copyright (c) 2017 Aspose.Words for Cloud
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
    using System.IO;
    
    using Com.Aspose.Storage.Api;

    /// <summary>
    /// Base class for all tests
    /// </summary>
    public abstract class BaseTestContext
    {
        // It is "test" credentials for "dev" server. Please, don't use them in youre application.
        protected const string BaseProductUri = @"http://api-dev.aspose.cloud";
        protected const string AppSID = "78b637f6-b4cc-41de-a619-d8bd9fc2b6b6";
        protected const string AppKey = "3d588eb82b3d5a634ad3141f09b03629";
        protected const string StorageAppSID = "ff470aee-3000-43dd-877d-e02e74499f18";
        protected const string StorageAppKey = "532a70d65e0a752d55673b86f10e53fc";

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseTestContext"/> class.
        /// </summary>
        protected BaseTestContext()
        {
            var configuration = new Configuration { ApiBaseUrl = BaseProductUri, AppKey = AppKey, AppSid = AppSID };
            this.WordsApi = new WordsApi(configuration);
            this.StorageApi = new StorageApi(AppKey, AppSID, BaseProductUri + "/v1.1");
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
        /// Returns common folder with source test files
        /// </summary>
        protected static string CommonFolder
        {
            get
            {
                return "Common/";
            }
        }

        /// <summary>
        /// Returns folder with source for document conversion tests
        /// </summary>
        protected static string ConvertFolder
        {
            get
            {
                return "ConvertDocument/";
            }
        }

        /// <summary>
        /// Returns folder with source for document protection tests
        /// </summary>
        protected static string ProtectFolder
        {
            get
            {
                return "DocumentProtection/";
            }
        }

        /// <summary>
        /// Returns folder with source for drawing objects tests
        /// </summary>
        protected static string DrawingFolder
        {
            get
            {
                return "Drawing/";
            }
        }

        /// <summary>
        /// Returns folder with source for fields tests
        /// </summary>
        protected static string FieldFolder
        {
            get
            {
                return "Field/";
            }
        }

        /// <summary>
        /// Returns folder with source for mail merge tests
        /// </summary>
        protected static string MailMergeFolder
        {
            get
            {
                return "MailMerge/";
            }
        }

        /// <summary>
        /// Returns folder with source for text tests
        /// </summary>
        protected static string TextFolder
        {
            get
            {
                return "Text/";
            }
        }

        /// <summary>
        /// Storage API
        /// </summary>
        protected StorageApi StorageApi { get; set; }

        /// <summary>
        /// Words API
        /// </summary>
        protected WordsApi WordsApi { get; set; }

        /// <summary>
        /// Returns test data path
        /// </summary>
        /// <param name="subfolder">subfolder for specific tests</param>
        /// <returns>test data path</returns>
        protected static string GetDataDir(string subfolder = null)
        {
            return Path.Combine("TestData", string.IsNullOrEmpty(subfolder) ? string.Empty : subfolder);
        }        
    }
}