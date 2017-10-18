// // --------------------------------------------------------------------------------------------------------------------
// // <copyright company="Aspose" file="DocumentProperties.cs">
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

namespace Aspose.Words.Cloud.Sdk.Tests.DocumentProperties
{
    using Aspose.Words.Cloud.Sdk.Model;
    using Aspose.Words.Cloud.Sdk.Model.Requests;
    using Aspose.Words.Cloud.Sdk.Tests.Base;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Example about how to get document properties
    /// </summary>
    public class DocumentProperties : BaseTestContext
    {
        /// <summary>
        /// A test for GetDocumentProperties
        /// </summary>
        [TestMethod]
        public void TestGetDocumentProperties()
        {
            string name = "test_multi_pages.docx";
            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentPropertiesRequest(name);
            var actual = this.WordsApi.GetDocumentProperties(request);
            Assert.AreEqual(200, actual.Code);
        }


        /// <summary>
        /// A test for GetDocumentProperty
        /// </summary>
        [TestMethod]
        public void TestGetDocumentProperty()
        {
            string name = "test_multi_pages.docx";
            string propertyName = "Author";
            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentPropertyRequest(name, propertyName);
            var actual = this.WordsApi.GetDocumentProperty(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// Test for deleting document property
        /// </summary>
        [TestMethod]
        public void TestDeleteDocumentProperty()
        {
            string name = "test_multi_pages.docx";
            string propertyName = "AsposeAuthor";
            string filename = "test_multi_pages.docx";

            var body = new DocumentProperty { Name = "AsposeAuthor", Value = "Imran Anwar", BuiltIn = false };

            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            // setting a property
            var updateRequest = new CreateOrUpdateDocumentPropertyRequest(name, propertyName, body, destFileName: filename);
            this.WordsApi.CreateOrUpdateDocumentProperty(updateRequest);

            var deleteRequest = new DeleteDocumentPropertyRequest(name, propertyName, destFileName: filename);
            var actual = this.WordsApi.DeleteDocumentProperty(deleteRequest);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// Test for updating document property
        /// </summary>
        [TestMethod]
        public void TestPutUpdateDocumentProperty()
        {
            string name = "test_multi_pages.docx";
            string propertyName = "Author";
            string filename = "test_multi_pages.docx";
            DocumentProperty body = new DocumentProperty { Name = "Author", Value = "Imran Anwar" };

            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new CreateOrUpdateDocumentPropertyRequest(name, propertyName, body, destFileName: filename);
            var actual = this.WordsApi.CreateOrUpdateDocumentProperty(request);
            Assert.AreEqual(200, actual.Code);
        }
    }
}