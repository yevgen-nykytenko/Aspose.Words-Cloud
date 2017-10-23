// // --------------------------------------------------------------------------------------------------------------------
// // <copyright company="Aspose" file="ExecuteMailMergeTest.cs">
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
// //  --------------------------------------------------------------------------------------------------------------------
namespace Aspose.Words.Cloud.Sdk.Tests.MailMerge
{
    using System.IO;

    using Aspose.Words.Cloud.Sdk.Model.Requests;
    using Aspose.Words.Cloud.Sdk.Tests.Base;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Example about how to perform mail merge
    /// </summary>
    [TestClass]
    public class ExecuteMailMergeTest : BaseTestContext
    {
        private readonly string dataFolder = Path.Combine(BaseTestDataPath, "DocumentActions/MailMerge");

        /// <summary>
        /// Test for executing mail merge online
        /// </summary>
        [TestMethod]
        public void TestPutExecuteMailMergeOnline()
        {
            using (var file = System.IO.File.OpenRead(Common.GetDataDir() + "SampleExecuteTemplate.docx"))
            {
                using (var data = System.IO.File.OpenRead(Common.GetDataDir() + "SampleExecuteTemplateData.txt"))
                {
                    var request = new PutExecuteMailMergeOnlineRequest(file, data);
                    var result = this.WordsApi.PutExecuteMailMergeOnline(request);
                    Assert.IsTrue(result.Length > 0, "Error occured while executing mail merge");
                }
            }
        }

        /// <summary>
        /// Test for executing mail merge
        /// </summary>
        [TestMethod]
        public void TestPostDocumentExecuteMailMerge()
        {
            var localName = "SampleMailMergeTemplate.docx";
            var remoteName = "TestPostDocumentExecuteMailMerge.docx";
            var fullName = Path.Combine(this.dataFolder, remoteName);
            var destFileName = Path.Combine(BaseTestOutPath, remoteName);
            var data = System.IO.File.ReadAllText(Common.GetDataDir() + "SampleMailMergeTemplateData.txt");

            this.StorageApi.PutCreate(fullName, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + localName));

            // TODO check why when destFileName contains folder, folder is doubled while saving
            var request = new PostDocumentExecuteMailMergeRequest(remoteName, false, data, this.dataFolder, destFileName: destFileName);
            var actual = this.WordsApi.PostDocumentExecuteMailMerge(request);

            Assert.AreEqual(200, actual.Code);
        }
    }
}