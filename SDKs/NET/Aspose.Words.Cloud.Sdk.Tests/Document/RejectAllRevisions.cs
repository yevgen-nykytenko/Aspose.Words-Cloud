// // --------------------------------------------------------------------------------------------------------------------
// // <copyright company="Aspose" file="RejectAllRevisions.cs">
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
    /// Example about how to reject all revisions in document
    /// </summary>
    public class RejectAllRevisions : BaseTestContext
    {
        /// <summary>
        /// Test for rejecting revisions in document
        /// </summary>
        [TestMethod]
        public void TestRejectAllRevisions()
        {
            string name = "test_multi_pages.docx";
            string filename = "test_multi_pages.docx";

            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new RejectAllRevisionsRequest(name, destFileName: filename);
            var actual = this.WordsApi.RejectAllRevisions(request);

            Assert.AreEqual(200, actual.Code);
        }
    }
}