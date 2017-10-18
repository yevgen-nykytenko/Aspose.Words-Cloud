// // --------------------------------------------------------------------------------------------------------------------
// // <copyright company="Aspose" file="DrawingObjects.cs">
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
namespace Aspose.Words.Cloud.Sdk.Tests.Drawing
{
    using Aspose.Words.Cloud.Sdk.Model;
    using Aspose.Words.Cloud.Sdk.Model.Requests;
    using Aspose.Words.Cloud.Sdk.Tests.Base;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Example about how to get drawing objects
    /// </summary>
    public class DrawingObjects : BaseTestContext
    {
        /// <summary>
        /// Test fir getting drawing objects from document
        /// </summary>
        [TestMethod]
        public void TestGetDocumentDrawingObjects()
        {
            string name = "test_multi_pages.docx";

            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentDrawingObjectsRequest(name, nodePath: "sections/0");
            var actual = this.WordsApi.GetDocumentDrawingObjects(request);

            Assert.AreEqual(200, actual.Code);
        }


        /// <summary>
        /// Test for getting drawing object by specified index
        /// </summary>
        [TestMethod]
        public void TestGetDocumentDrawingObjectByIndex()
        {
            string name = "test_multi_pages.docx";
            int objectIndex = 0;

            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentDrawingObjectByIndexRequest(name, objectIndex, nodePath: "sections/0");
            DrawingObjectResponse actual = this.WordsApi.GetDocumentDrawingObjectByIndex(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// Test for getting drawing object by specified index and format
        /// </summary>
        [TestMethod]
        public void TestGetDocumentDrawingObjectByIndexWithFormat()
        {
            string name = "test_multi_pages.docx";
            int objectIndex = 0;
            string format = "png";

            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new RenderDrawingObjectRequest(name, format, objectIndex, nodePath: "sections/0");
            var result = this.WordsApi.RenderDrawingObject(request);
            Assert.IsTrue(result.Length > 0, "Error occured while getting drawing object");
        }

        /// <summary>
        /// Test for reading drawing object's image data
        /// </summary>
        [TestMethod]
        public void TestGetDocumentDrawingObjectImageData()
        {
            string name = "test_multi_pages.docx";
            int objectIndex = 0;
            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentDrawingObjectImageDataRequest(name, objectIndex, nodePath: "sections/0");
            var result = this.WordsApi.GetDocumentDrawingObjectImageData(request);
            Assert.IsTrue(result.Length > 0, "Error occured while getting drawing object");
        }

        /// <summary>
        /// Test for getting drawing object OLE data
        /// </summary>
        [TestMethod]
        public void TestGetDocumentDrawingObjectOleData()
        {
            string name = "sample_EmbeddedOLE.docx";
            int objectIndex = 0;
            this.StorageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentDrawingObjectOleDataRequest(name, objectIndex, nodePath: "sections/0");
            this.WordsApi.GetDocumentDrawingObjectOleData(request);
        }
    }
}