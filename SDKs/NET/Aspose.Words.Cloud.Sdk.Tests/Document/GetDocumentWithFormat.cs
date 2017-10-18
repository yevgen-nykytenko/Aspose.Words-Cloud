// // --------------------------------------------------------------------------------------------------------------------
// // <copyright company="Aspose" file="GetDocumentWithFormat.cs">
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

namespace Aspose.Words.Cloud.Sdk.Tests.Document
{
    using Aspose.Words.Cloud.Sdk.Model.Requests;
    using Aspose.Words.Cloud.Sdk.Tests.Base;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Example about how to get document with different format
    /// </summary>
    public class GetDocumentWithFormat : BaseTestContext
    {
        /// <summary>
        /// Test for getting document with specified format
        /// </summary>
        [TestMethod]
        public void TestGetDocumentWithFormat()
        {
            string name = "test_multi_pages.docx";
            string format = "text";

            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentWithFormatRequest(name, format);
            var result = this.WordsApi.GetDocumentWithFormat(request);
            Assert.IsTrue(result.Length > 0, "Conversion has failed");
        }

        /// <summary>
        /// Test for getting document with specified format and outPath
        /// </summary>
        [TestMethod]
        public void TestGetDocumentWithFormatAndOutPath()
        {
            string name = "test_multi_pages.docx";
            string format = "text";
            string outPath = "out/test_multi_pages.text";

            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentWithFormatRequest(name, format, outPath: outPath);
            this.WordsApi.GetDocumentWithFormat(request);
            var result = this.StorageApi.GetIsExist(outPath, null, null);
            Assert.IsNotNull(result, "Cannot download document from storage");
            Assert.IsTrue(result.FileExist.IsExist, "File doesn't exist on storage");
        }

        /// <summary>
        /// Test for getting document with specified format and storage
        /// </summary>
        [TestMethod]
        public void TestGetDocumentFormatUsingStorage()
        {
            string name = "test_multi_pages.docx";
            string format = "text";
            string storage = "dropboxstorage";

            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentWithFormatRequest(name, format, storage: storage);
            var result = this.WordsApi.GetDocumentWithFormat(request);
            Assert.IsTrue(result.Length > 0, "Conversion has failed");
        }
    }
}