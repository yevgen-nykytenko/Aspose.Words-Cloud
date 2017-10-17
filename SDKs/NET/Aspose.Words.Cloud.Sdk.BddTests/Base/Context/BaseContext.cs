// --------------------------------------------------------------------------------------------------------------------
// <copyright company="Aspose" file="BaseContext.cs">
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

namespace Aspose.Words.Cloud.Sdk.BddTests.Base.Context
{
    using System.IO;

    using Aspose.Words.Cloud.Sdk.Api;

    using Com.Aspose.Storage.Api;

    /// <summary>
    /// Base context for all tests.
    /// </summary>
    public class BaseContext
    {
        private const string BaseProductUri = @"http://api-dev.aspose.cloud/v1.1";
        private const string AppSID = "78b637f6-b4cc-41de-a619-d8bd9fc2b6b6";
        private const string AppKey = "3d588eb82b3d5a634ad3141f09b03629";

        private string testFolder;

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseContext"/>.
        /// </summary>
        protected BaseContext()
        {            
            this.WordsApi = new WordsApi(AppKey, AppSID, BaseProductUri);
            this.StorageApi = new StorageApi(AppKey, AppSID, BaseProductUri);
        }

        /// <summary>
        /// Storage API
        /// </summary>
        public StorageApi StorageApi { get; set; }

        /// <summary>
        /// Words API
        /// </summary>
        public WordsApi WordsApi { get; set; }

        /// <summary>
        /// Response.
        /// </summary>
        public object Response { get; set; }        

        /// <summary>
        /// Get path to test data
        /// </summary>
        public string TestDataPath
        {
            get
            {
                return this.testFolder ?? (this.testFolder = Path.Combine(DirectoryHelper.GetTestDataPath(), this.TestSubFolderInStorage));
            }
        }

        /// <summary>
        /// Folder name
        /// </summary>
        public string TestFolderInStorage
        {
            get
            {
                return "TempSDKTests/" + this.TestSubFolderInStorage;
            }
        }

        /// <summary>
        /// Subfolder name for specific test data
        /// </summary>
        public string TestSubFolderInStorage { get; set; }

        /// <summary>
        /// Is document with this name exist
        /// </summary>
        /// <param name="name">document name</param>
        /// <returns>is exist</returns>
        public bool FileWithNameExists(string name)
        {
            var isExists = this.StorageApi.GetIsExist(Path.Combine(this.TestFolderInStorage, name), null, null);
            if (isExists != null && isExists.FileExist != null)
            {
                return isExists.FileExist.IsExist;
            }

            return false;
        }
    }
}