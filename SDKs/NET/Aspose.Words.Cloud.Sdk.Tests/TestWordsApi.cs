// --------------------------------------------------------------------------------------------------------------------
// <copyright company="Aspose" file="TestWordsApi.cs">
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

namespace Aspose.Words.Cloud.Sdk.Tests
{
    using System.Diagnostics;

    using Aspose.Words.Cloud.Sdk;
    using Aspose.Words.Cloud.Sdk.Api;
    using Aspose.Words.Cloud.Sdk.Model;
    using Aspose.Words.Cloud.Sdk.Model.Requests;

    using Com.Aspose.Storage.Api;

    using Microsoft.VisualStudio.TestTools.UnitTesting;

    using NMock;

    /// <summary>
    /// This is a test class for TestWordsApi and is intended
    /// to contain all TestWordsApi Unit Tests
    /// </summary>
    [TestClass]
    [DeploymentItem("Data", "Data")]
    public class TestWordsApi
    {
        private const string ApiKey = "0fbf678c5ecabdb5caca48452a736dd0";
        private const string ApiSid = "91a2fd07-bba1-4b32-9112-abfb1fe8aebd";
        private const string AppUrl = "http://api.aspose.cloud/v1.1";

        private readonly WordsApi wordsApi;
        private readonly StorageApi storageApi;

        /// <summary>
        /// A test for WordsApi Constructor
        /// </summary>
        public TestWordsApi()
        {
            this.wordsApi = new WordsApi(ApiKey, ApiSid, AppUrl);
            this.storageApi = new StorageApi(ApiKey, ApiSid, AppUrl);
        }

        /// <summary>
        /// Gets or sets the test context which provides
        /// information about and functionality for the current test run.
        /// </summary>
        public TestContext TestContext { get; set; }

        /// <summary>
        /// A test for AcceptAllRevisions
        /// </summary>
        [TestMethod]
        public void TestAcceptAllRevisions()
        {
            string name = "test_multi_pages.docx";
            string filename = "Test2.docx";
         
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new AcceptAllRevisionsRequest(name, destFileName: filename);
            var actual = this.wordsApi.AcceptAllRevisions(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for DeleteComment
        /// </summary>
        [TestMethod]
        public void TestDeleteComment()
        {
            string name = "test_multi_pages.docx";
            int commentIndex = 0;           
            string fileName = "test_multi_pages.docx";
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new DeleteCommentRequest(name, commentIndex, destFileName: fileName);
            var actual = this.wordsApi.DeleteComment(request);
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for DeleteDocumentFields
        /// </summary>
        [TestMethod]
        public void TestDeleteDocumentFields()
        {
            string name = "test_multi_pages.docx";
         
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));
            
            var request = new DeleteFieldsRequest(name);
            var actual = this.wordsApi.DeleteFields(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for DeleteDocumentMacros
        /// </summary>
        [TestMethod]
        public void TestDeleteDocumentMacros()
        {
            string name = "test_multi_pages.docx";
         
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new DeleteDocumentMacrosRequest(name);
            var actual = this.wordsApi.DeleteDocumentMacros(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for DeleteDocumentProperty
        /// </summary>
        [TestMethod]
        public void TestDeleteDocumentProperty()
        {
            string name = "test_multi_pages.docx";
            string propertyName = "AsposeAuthor";
            string filename = "test_multi_pages.docx";
          
            var body = new DocumentProperty();
            body.Name = "AsposeAuthor";
            body.Value = "Imran Anwar";
            body.BuiltIn = false;

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            // setting a property
            var updateRequest = new CreateOrUpdateDocumentPropertyRequest(name, propertyName, body, destFileName: filename);
            this.wordsApi.CreateOrUpdateDocumentProperty(updateRequest);

            var deleteRequest = new DeleteDocumentPropertyRequest(name, propertyName, destFileName: filename);
            var actual = this.wordsApi.DeleteDocumentProperty(deleteRequest);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for DeleteDocumentWatermark
        /// </summary>
        [TestMethod]
        public void TestDeleteDocumentWatermark()
        {
            string name = "test_multi_pages.docx";
            string filename = "test.docx";
         
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new DeleteDocumentWatermarkRequest(name, destFileName: filename);
            var actual = this.wordsApi.DeleteDocumentWatermark(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for DeleteFormField
        /// </summary>
        [TestMethod]
        public void TestDeleteFormField()
        {
            string name = "FormFilled.docx";            
            int formfieldIndex = 0;
            string destFileName = "FormFilledTest.docx";
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new DeleteFormFieldRequest(name, formfieldIndex, nodePath: "sections/0", destFileName: destFileName);
            SaaSposeResponse actual = this.wordsApi.DeleteFormField(request);

            Assert.AreEqual(200, actual.Code);                        
        }

        /// <summary>
        /// A test for DeleteField
        /// </summary>
        [TestMethod]
        public void TestDeleteField()
        {
            string name = "GetField.docx";         
            int fieldIndex = 0;           
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new DeleteFieldRequest(name, fieldIndex, nodePath: "sections/0/paragraphs/0");
            var actual = this.wordsApi.DeleteField(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for DeleteHeadersFooters
        /// </summary>
        [TestMethod]
        public void TestDeleteHeadersFooters()
        {
            string name = "test_multi_pages.docx";
            string filename = "test.docx";

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new DeleteHeadersFootersRequest(name, sectionPath: "sections/0", destFileName: filename);
            SaaSposeResponse actual = this.wordsApi.DeleteHeadersFooters(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for DeleteParagraphFields
        /// </summary>
        [TestMethod]
        public void TestDeleteParagraphFields()
        {
            string name = "test_multi_pages.docx";
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new DeleteFieldsRequest(name, nodePath: "paragraphs/0");
            var actual = this.wordsApi.DeleteFields(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for DeleteSectionFields
        /// </summary>
        [TestMethod]
        public void TestDeleteSectionFields()
        {
            string name = "test_multi_pages.docx";            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new DeleteFieldsRequest(name, nodePath: "sections/0");
            var actual = this.wordsApi.DeleteFields(request);

            Assert.AreEqual(200, actual.Code);            
        }

        /// <summary>
        /// A test for GetRenderPage
        /// </summary>
        [TestMethod]
        public void TestGetRenderPage()
        {
            string name = "SampleWordDocument.docx";
            int pageNumber = 1;
            string format = "bmp";
           
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new RenderPageRequest(name, pageNumber, format);
            this.wordsApi.RenderPage(request);

            // TODO: check response file
        }

        /// <summary>
        /// A test for DeleteSectionParagraphFields
        /// </summary>
        [TestMethod]
        public void TestDeleteSectionParagraphFields()
        {
            string name = "test_multi_pages.docx";
         
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new DeleteFieldsRequest(name, nodePath: "sections/0/paragraphs/0");
            var actual = this.wordsApi.DeleteFields(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for DeleteUnprotectDocument
        /// </summary>
        [TestMethod]
        public void TestDeleteUnprotectDocument()
        {
            string name = "SampleProtectedBlankWordDocument.docx";
            
            var body = new ProtectionRequest();
            body.Password = "aspose";

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new DeleteUnprotectDocumentRequest(name, body);
            ProtectionDataResponse actual = this.wordsApi.DeleteUnprotectDocument(request);
            
            Assert.AreEqual(200, actual.Code);            
        }

        /// <summary>
        /// A test for GetComment
        /// </summary>
        [TestMethod]
        public void TestGetComment()
        {
            string name = "test_multi_pages.docx";
            int commentIndex = 0;
         
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetCommentRequest(name, commentIndex);
            var actual = this.wordsApi.GetComment(request);

            Assert.AreEqual(200, actual.Code);            
        }

        /// <summary>
        /// A test for GetComments
        /// </summary>
        [TestMethod]
        public void TestGetComments()
        {
            string name = "test_multi_pages.docx";
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetCommentsRequest(name);
            var actual = this.wordsApi.GetComments(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocument
        /// </summary>
        [TestMethod]
        public void TestGetDocument()
        {            
            string name = "test_multi_pages.docx";          
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));
            
            var request = new GetDocumentRequest(name);
            var actual = this.wordsApi.GetDocument(request);

           Assert.AreEqual(200, actual.Code);            
        }

        /// <summary>
        /// A test for GetDocumentBookmarkByName
        /// </summary>
        [TestMethod]
        public void TestGetDocumentBookmarkByName()
        {
            string name = "test_multi_pages.docx";
            string bookmarkName = "aspose";           
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentBookmarkByNameRequest(name, bookmarkName);
            var actual = this.wordsApi.GetDocumentBookmarkByName(request);
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentBookmarks
        /// </summary>
        [TestMethod]
        public void TestGetDocumentBookmarks()
        {
            string name = "test_multi_pages.docx";
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentBookmarksRequest(name);
            var actual = this.wordsApi.GetDocumentBookmarks(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentDrawingObjectByIndex
        /// </summary>
        [TestMethod]
        public void TestGetDocumentDrawingObjectByIndex()
        {
            string name = "test_multi_pages.docx";
            int objectIndex = 0;
          
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentDrawingObjectByIndexRequest(name, objectIndex, nodePath: "sections/0");
            DrawingObjectResponse actual = this.wordsApi.GetDocumentDrawingObjectByIndex(request);

            Assert.AreEqual(200, actual.Code);                        
        }

        /// <summary>
        /// A test for GetDocumentDrawingObjectByIndexWithFormat
        /// </summary>
        [TestMethod]
        public void TestGetDocumentDrawingObjectByIndexWithFormat()
        {
            string name = "test_multi_pages.docx";
            int objectIndex = 0;
            string format = "png";
           
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new RenderDrawingObjectRequest(name, format, objectIndex, nodePath: "sections/0");
            this.wordsApi.RenderDrawingObject(request);                        
        }

        /// <summary>
        /// A test for GetDocumentDrawingObjectImageData
        /// </summary>
        [TestMethod]
        public void TestGetDocumentDrawingObjectImageData()
        {
            string name = "test_multi_pages.docx";
            int objectIndex = 0;            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentDrawingObjectImageDataRequest(name, objectIndex, nodePath: "sections/0");
            this.wordsApi.GetDocumentDrawingObjectImageData(request);            
        }

        /// <summary>
        /// A test for GetDocumentDrawingObjectOleData
        /// </summary>
        [TestMethod]
        public void TestGetDocumentDrawingObjectOleData()
        {
            string name = "sample_EmbeddedOLE.docx";
            int objectIndex = 0;             
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentDrawingObjectOleDataRequest(name, objectIndex, nodePath: "sections/0");
            this.wordsApi.GetDocumentDrawingObjectOleData(request);            
        }

        /// <summary>
        /// A test for GetDocumentDrawingObjects
        /// </summary>
        [TestMethod]
        public void TestGetDocumentDrawingObjects()
        {
            string name = "test_multi_pages.docx";
         
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentDrawingObjectsRequest(name, nodePath: "sections/0");
            var actual = this.wordsApi.GetDocumentDrawingObjects(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentFieldNames
        /// </summary>
        [TestMethod]
        public void TestGetDocumentFieldNames()
        {
            string name = "test_multi_pages.docx";
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentFieldNamesRequest(name);
            var actual = this.wordsApi.GetDocumentFieldNames(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentHyperlinkByIndex
        /// </summary>
        [TestMethod]
        public void TestGetDocumentHyperlinkByIndex()
        {
            string name = "test_doc.docx";
            int hyperlinkIndex = 0;         
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentHyperlinkByIndexRequest(name, hyperlinkIndex);
            var actual = this.wordsApi.GetDocumentHyperlinkByIndex(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentHyperlinks
        /// </summary>
        [TestMethod]
        public void TestGetDocumentHyperlinks()
        {
            string name = "test_doc.docx";           
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentHyperlinksRequest(name);
            var actual = this.wordsApi.GetDocumentHyperlinks(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentParagraph
        /// </summary>
        [TestMethod]
        public void TestGetDocumentParagraph()
        {
            string name = "test_multi_pages.docx";
            int index = 0;

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentParagraphRequest(name, index, nodePath: "sections/0");
            var actual = this.wordsApi.GetDocumentParagraph(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentParagraph
        /// </summary>
        [TestMethod]
        public void TestGetDocumentParagraphWithoutNodePath()
        {
            string name = "test_multi_pages.docx";
            int index = 0;

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentParagraphRequest(name, index);
            var actual = this.wordsApi.GetDocumentParagraph(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentParagraphRun
        /// </summary>
        [TestMethod]
        public void TestGetDocumentParagraphRun()
        {
            string name = "test_multi_pages.docx";            
            int runIndex = 0;             

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentParagraphRunRequest(name, "paragraphs/0", runIndex);
            var actual = this.wordsApi.GetDocumentParagraphRun(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentParagraphRunFont
        /// </summary>
        [TestMethod]
        public void TestGetDocumentParagraphRunFont()
        {
            string name = "test_multi_pages.docx";            
            int runIndex = 0;           

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentParagraphRunFontRequest(name, "paragraphs/0", runIndex);
            var actual = this.wordsApi.GetDocumentParagraphRunFont(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentParagraphs
        /// </summary>
        [TestMethod]
        public void TestGetDocumentParagraphs()
        {
            string name = "test_multi_pages.docx";
         
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentParagraphsRequest(name, nodePath: "sections/0");
            var actual = this.wordsApi.GetDocumentParagraphs(request);

            Assert.AreEqual(200, actual.Code);            
        }

        /// <summary>
        /// A test for GetDocumentProperties
        /// </summary>
        [TestMethod]
        public void TestGetDocumentProperties()
        {
            string name = "test_multi_pages.docx";            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentPropertiesRequest(name);
            var actual = this.wordsApi.GetDocumentProperties(request);
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
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentPropertyRequest(name, propertyName);
            var actual = this.wordsApi.GetDocumentProperty(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentProtection
        /// </summary>
        [TestMethod]
        public void TestGetDocumentProtection()
        {
            string name = "test_multi_pages.docx";
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));
            
            var request = new GetDocumentProtectionRequest(name);
            var actual = this.wordsApi.GetDocumentProtection(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentStatistics
        /// </summary>
        [TestMethod]
        public void TestGetDocumentStatistics()
        {
            string name = "test_multi_pages.docx";
          
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentStatisticsRequest(name);
            var actual = this.wordsApi.GetDocumentStatistics(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentTextItems
        /// </summary>
        [TestMethod]
        public void TestGetDocumentTextItems()
        {
            string name = "test_multi_pages.docx";
          
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));
            
            var request = new GetDocumentTextItemsRequest(name);
            var actual = this.wordsApi.GetDocumentTextItems(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetDocumentWithFormat
        /// </summary>
        [TestMethod]
        public void TestGetDocumentWithFormat()
        {
            string name = "test_multi_pages.docx";
            string format = "text";
     
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentWithFormatRequest(name, format);
            this.wordsApi.GetDocumentWithFormat(request);

            // TODO: add case with specified out path            
        }

        /// <summary>
        /// A test for GetField
        /// </summary>
        [TestMethod]
        public void TestGetField()
        {
            string name = "GetField.docx";         
            int fieldIndex = 0;

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetFieldRequest(name, fieldIndex, nodePath: "sections/0/paragraphs/0");
            var actual = this.wordsApi.GetField(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetFields
        /// </summary>
        [TestMethod]
        public void TestGetFields()
        {
            string name = "GetField.docx";
           
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetFieldsRequest(name, nodePath: "sections/0");
            FieldsResponse actual = this.wordsApi.GetFields(request);

           Assert.AreEqual(200, actual.Code);           
        }

        /// <summary>
        /// A test for TestGetParagraphRuns
        /// </summary>
        [TestMethod]
        public void TestGetParagraphRuns()
        {
            string name = "GetField.docx";
          
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetDocumentParagraphRunsRequest(name, "sections/0/paragraphs/0");
            RunsResponse actual = this.wordsApi.GetDocumentParagraphRuns(request);

            Assert.AreEqual(200, actual.Code);            
        }

        /// <summary>
        ///  A test for GetFormFields
        /// </summary>        
        [TestMethod]
        public void TestGetFormFields()
        {
            string name = "FormFilled.docx";
          
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetFormFieldsRequest(name, nodePath: "sections/0");
            FormFieldsResponse actual = this.wordsApi.GetFormFields(request);

           Assert.AreEqual(200, actual.Code);            
        }

        /// <summary>
        /// A test for GetFormField
        /// </summary>
        [TestMethod]
        public void TestGetFormField()
        {
            string name = "FormFilled.docx";
            int formfieldIndex = 0;

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetFormFieldRequest(name, formfieldIndex, nodePath: "sections/0");
            FormFieldResponse actual = this.wordsApi.GetFormField(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetSection
        /// </summary>
        [TestMethod]
        public void TestGetSection()
        {
            string name = "test_multi_pages.docx";
            int sectionIndex = 0;
           
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetSectionRequest(name, sectionIndex);
            var actual = this.wordsApi.GetSection(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetSectionPageSetup
        /// </summary>
        [TestMethod]
        public void TestGetSectionPageSetup()
        {
            string name = "test_multi_pages.docx";
            int sectionIndex = 0;
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetSectionPageSetupRequest(name, sectionIndex);
            var actual = this.wordsApi.GetSectionPageSetup(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for GetSections
        /// </summary>
        [TestMethod]
        public void TestGetSections()
        {
            string name = "test_multi_pages.docx";
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new GetSectionsRequest(name);
            var actual = this.wordsApi.GetSections(request);

            Assert.AreEqual(200, actual.Code);            
        }

        /// <summary>
        /// A test for PostAppendDocument
        /// </summary>
        [TestMethod]
        public void TestPostAppendDocument()
        {
            string name = "test_multi_pages.docx";
            string filename = "test_multi_pages.docx";
           
            var body = new DocumentEntryList();
            System.Collections.Generic.List<DocumentEntry> docEntries = new System.Collections.Generic.List<DocumentEntry>();

            DocumentEntry docEntry = new DocumentEntry();
            docEntry.Href = "test_multi_pages.docx";
            docEntry.ImportFormatMode = "KeepSourceFormatting";
            docEntries.Add(docEntry);
            body.DocumentEntries = docEntries;            

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostAppendDocumentRequest(name, body, destFileName: filename);
            var actual = this.wordsApi.PostAppendDocument(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostChangeDocumentProtection
        /// </summary>
        [TestMethod]
        public void TestPostChangeDocumentProtection()
        {
            string name = "test_multi_pages.docx";
            
            var body = new ProtectionRequest();
            body.NewPassword = string.Empty;

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostChangeDocumentProtectionRequest(name, body);
            var actual = this.wordsApi.PostChangeDocumentProtection(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostComment
        /// </summary>
        [TestMethod]
        public void TestPostComment()
        {
            string name = "test_multi_pages.docx";
            int commentIndex = 0;
           
            var body = new Comment();

            var dpdto = new DocumentPosition();
            NodeLink nodeLink = new NodeLink();
            
            dpdto.Node = nodeLink;
            dpdto.Offset = 0;
            
            nodeLink.NodeId = "0.0.3";

            body.RangeStart = dpdto;
            body.RangeEnd = dpdto;

            body.Initial = "IA";
            body.Author = "Imran Anwar";
            body.Text = "A new Comment";

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostCommentRequest(name, commentIndex, body);
            var actual = this.wordsApi.PostComment(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostDocumentExecuteMailMerge
        /// </summary>
        [TestMethod]
        public void TestPostDocumentExecuteMailMerge()
        {
            string name = "SampleMailMergeTemplate.docx";                       
            string filename = "SampleMailMergeResult.docx";            
            
            var data = System.IO.File.ReadAllText(Common.GetDataDir() + "SampleMailMergeTemplateData.txt");

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostDocumentExecuteMailMergeRequest(name, false, data, destFileName: filename);
            var actual = this.wordsApi.PostDocumentExecuteMailMerge(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostDocumentParagraphRunFont
        /// </summary>
        [TestMethod]
        public void TestPostDocumentParagraphRunFont()
        {
            string name = "test_multi_pages.docx";            
            int runIndex = 0;
            string filename = "test.docx";
            var body = new Font();
            body.Bold = true;
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostDocumentParagraphRunFontRequest(name, body, "paragraphs/0", runIndex, destFileName: filename);
            var actual = this.wordsApi.PostDocumentParagraphRunFont(request);
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostDocumentSaveAs
        /// </summary>
        [TestMethod]
        public void TestPostDocumentSaveAs()
        {
            string name = "test_multi_pages.docx";
           
            var body = new SaveOptionsData();

            body.SaveFormat = "pdf";
            body.FileName = "output.pdf";

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostDocumentSaveAsRequest(name, body);
            var actual = this.wordsApi.PostDocumentSaveAs(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostExecuteTemplate
        /// </summary>
        [TestMethod]
        public void TestPostExecuteTemplate()
        {
            string name = "TestExecuteTemplate.doc";
            
            string destFileName = "TestExecuteResult.doc";
                        
            string data = System.IO.File.ReadAllText(Common.GetDataDir() + "TestExecuteTemplateData.txt");

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostExecuteTemplateRequest(name, data, destFileName: destFileName);
            var actual = this.wordsApi.PostExecuteTemplate(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostField
        /// </summary>
        [TestMethod]
        public void TestPostField()
        {
            string name = "GetField.docx";         
            int fieldIndex = 0;
            string destFileName = "newSampleWordDocument.docx";
            
            var body = new Field();
            body.Result = "3";
            body.FieldCode = "{ NUMPAGES }";
            
            body.NodeId = "0.0.3";
                        
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostFieldRequest(name, body, fieldIndex, nodePath: "sections/0/paragraphs/0", destFileName: destFileName);
            var actual = this.wordsApi.PostField(request);
            
            Assert.AreEqual(200, actual.Code);            
        }

        /// <summary>
        /// A test for PostFormField
        /// </summary>
        [TestMethod]
        public void TestPostFormField()
        {
            // Arrange
            string name = "FormFilled.docx";           
            int formfieldIndex = 0;
            string destFileName = "newFormFilled.docx";
            
            FormFieldTextInput body = new FormFieldTextInput();

            body.Name = "FullName";
            body.Enabled = true;
            body.CalculateOnExit = true;
            body.StatusText = string.Empty;
                     
            body.TextInputType = FormFieldTextInput.TextInputTypeEnum.Regular;
            body.TextInputDefault = string.Empty;

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostFormFieldRequest(name, body, formfieldIndex, nodePath: "sections/0", destFileName: destFileName);

            // Act
            FormFieldResponse actual = this.wordsApi.PostFormField(request);
            
            // Assert
            Assert.AreEqual(200, actual.Code);
            Assert.AreEqual("FullName", actual.FormField.Name);
            Assert.AreEqual(true, actual.FormField.Enabled);

            var formFieldTextInput = actual.FormField as FormFieldTextInput;
            Assert.IsTrue(formFieldTextInput != null, "Incorrect type of formfield: {0} instead of {1}", actual.FormField.GetType(), typeof(FormFieldTextInput));
            Assert.AreEqual(FormFieldTextInput.TextInputTypeEnum.Regular, formFieldTextInput.TextInputType);
        }

        /// <summary>
        /// A test for PostInsertDocumentWatermarkImage
        /// </summary>
        [TestMethod]
        public void TestPostInsertDocumentWatermarkImage()
        {
            string name = "test_multi_pages.docx";
            string filename = "test.docx";
            double rotationAngle = 0F;
            string image = "aspose-cloud.png";
           
            using (var file = System.IO.File.OpenRead(Common.GetDataDir() + image))
            {
                this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

                var request = new PostInsertDocumentWatermarkImageRequest(name,
                    file,
                    rotationAngle: rotationAngle,
                    destFileName: filename);

                var actual = this.wordsApi.PostInsertDocumentWatermarkImage(request);

                Assert.AreEqual(200, actual.Code);
            }
        }

        /// <summary>
        /// A test for PostInsertDocumentWatermarkText
        /// </summary>
        [TestMethod]
        public void TestPostInsertDocumentWatermarkText()
        {
            string name = "test_multi_pages.docx";
            string filename = "test.docx";
            
            var body = new WatermarkText();
            body.Text = "The watermark of Aspose";            

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostInsertDocumentWatermarkTextRequest(name, body, destFileName: filename);
            var actual = this.wordsApi.PostInsertDocumentWatermarkText(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostInsertPageNumbers
        /// </summary>
        [TestMethod]
        public void TestPostInsertPageNumbers()
        {
            string name = "test_multi_pages.docx";
            string filename = "test_multi_pages.docx";
            
            var body = new PageNumber();
            body.Alignment = "center";
            body.Format = "{PAGE} of {NUMPAGES}";
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostInsertPageNumbersRequest(name, body, destFileName: filename);
            var actual = this.wordsApi.PostInsertPageNumbers(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostInsertWatermarkImage
        /// </summary>
        [TestMethod]
        public void TestPostInsertWatermarkImage()
        {
            string name = "test_multi_pages.docx";
            string filename = "TestPostInsertWatermarkImageOut.docx";
            double rotationAngle = 0F; // TODO: Initialize to an appropriate value
            string image = "aspose-cloud.png";
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));
            this.storageApi.PutCreate(image, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + image));

            var request = new PostInsertDocumentWatermarkImageRequest(name, image: image, rotationAngle: rotationAngle, destFileName: filename);
            var actual = this.wordsApi.PostInsertDocumentWatermarkImage(request);
            
            Assert.AreEqual(200, actual.Code);            
        }

        /// <summary>
        /// A test for PostInsertWatermarkText
        /// </summary>
        [TestMethod]
        public void TestPostInsertWatermarkText()
        {
            string name = "test_multi_pages.docx";
                       
            string filename = "test_multi_pages.docx";
           
            var body = new WatermarkText();
            body.Text = "This is the text";
            body.RotationAngle = 90.0f;
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostInsertDocumentWatermarkTextRequest(name, body, destFileName: filename);
            var actual = this.wordsApi.PostInsertDocumentWatermarkText(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostLoadWebDocument
        /// </summary>
        [TestMethod]
        public void TestPostLoadWebDocument()
        {
            var body = new LoadWebDocumentData();
            var soptions = new SaveOptionsData();
            soptions.FileName = "google.doc";
            soptions.SaveFormat = "doc";
            soptions.ColorMode = "1";
            soptions.DmlEffectsRenderingMode = "1";
            soptions.DmlRenderingMode = "1";
            soptions.UpdateSdtContent = false;
            soptions.ZipOutput = false;

            body.LoadingDocumentUrl = "http://google.com";
            body.SaveOptions = soptions;

            var request = new PostLoadWebDocumentRequest(body);
            var actual = this.wordsApi.PostLoadWebDocument(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostReplaceText
        /// </summary>
        [TestMethod]
        public void TestPostReplaceText()
        {
            string name = "test_multi_pages.docx";
            string filename = "test_multi_pages_result.docx";         
            var body = new ReplaceTextRequest();
            body.OldValue = "aspose";
            body.NewValue = "aspose new";

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostReplaceTextRequest(name, body, destFileName: filename);
            var actual = this.wordsApi.PostReplaceText(request);
            
            Assert.AreEqual(200, actual.Code);
        }
        
        /// <summary>
        /// A test for PostSplitDocument
        /// </summary>
        [TestMethod]
        public void TestPostSplitDocument()
        {
            string name = "test_multi_pages.docx";
            string format = "text";
            int from = 1; 
            int to = 2; 
           
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostSplitDocumentRequest(name, format: format, from: from, to: to);
            var actual = this.wordsApi.PostSplitDocument(request);
            
            Assert.AreEqual(200, actual.Code);            
        }

        /// <summary>
        /// A test for PostUpdateDocumentBookmark
        /// </summary>
        [TestMethod]
        public void TestPostUpdateDocumentBookmark()
        {
            string name = "test_multi_pages.docx";
            string bookmarkName = "aspose";
            string filename = "test.docx";          
            var body = new BookmarkData();
            body.Name = "aspose";
            body.Text = "This will be the text for Aspose";

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostUpdateDocumentBookmarkRequest(name, body, bookmarkName, destFileName: filename);
            var actual = this.wordsApi.PostUpdateDocumentBookmark(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PostUpdateDocumentFields
        /// </summary>
        [TestMethod]
        public void TestPostUpdateDocumentFields()
        {
            string name = "test_multi_pages.docx";
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PostUpdateDocumentFieldsRequest(name);
            var actual = this.wordsApi.PostUpdateDocumentFields(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PutComment
        /// </summary>
        [TestMethod]
        public void TestPutComment()
        {
            string name = "test_multi_pages.docx";           
            var body = new Comment();

            var dpdto = new DocumentPosition();
            NodeLink nodeLink = new NodeLink();
            
            dpdto.Node = nodeLink;
            dpdto.Offset = 0;            
            nodeLink.NodeId = "0.0.3";

            body.RangeStart = dpdto;
            body.RangeEnd = dpdto;

            body.Initial = "IA";
            body.Author = "Imran Anwar";
            body.Text = "A new Comment";

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));
                        
            var request = new PutCommentRequest(name, body);
            var actual = this.wordsApi.PutComment(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PutConvertDocument
        /// </summary>
        [TestMethod]
        public void TestPutConvertDocument()
        {
            string format = "pdf";            
            using (var fileStream = System.IO.File.OpenRead(Common.GetDataDir() + "test_uploadfile.docx"))
            {
                var request = new PutConvertDocumentRequest(fileStream, format);
                this.wordsApi.PutConvertDocument(request);             
            }
        }

        /// <summary>
        /// A test for PutDocumentFieldNames
        /// </summary>
        [TestMethod]
        public void TestPutDocumentFieldNames()
        {            
            using (var fileStream = System.IO.File.OpenRead(Common.GetDataDir() + "SampleExecuteTemplate.docx"))
            {
                var request = new PutDocumentFieldNamesRequest(fileStream, true);
                FieldNamesResponse actual = this.wordsApi.PutDocumentFieldNames(request);

                Assert.AreEqual(200, actual.Code);
            }
        }

        /// <summary>
        /// A test for PutDocumentSaveAsTiff
        /// </summary>
        [TestMethod]
        public void TestPutDocumentSaveAsTiff()
        {
            string name = "test_multi_pages.docx";
            string resultFile = "test.docx";
         
            TiffSaveOptionsData body = new TiffSaveOptionsData();
            body.FileName = "abc.tiff";
            body.SaveFormat = "tiff";

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));
            var request = new PutDocumentSaveAsTiffRequest(name,
                body,
                destFileName: resultFile);
            SaveResponse actual = this.wordsApi.PutDocumentSaveAsTiff(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PutExecuteMailMergeOnline
        /// </summary>
        [TestMethod]
        public void TestPutExecuteMailMergeOnline()
        {          
            using (var file = System.IO.File.OpenRead(Common.GetDataDir() + "SampleExecuteTemplate.docx"))
            {
                using (var data = System.IO.File.OpenRead(Common.GetDataDir() + "SampleExecuteTemplateData.txt"))
                {
                    var request = new PutExecuteMailMergeOnlineRequest(file, data);
                    this.wordsApi.PutExecuteMailMergeOnline(request);                   
                }
            }
        }

        /// <summary>
        /// A test for PutExecuteTemplateOnline
        /// </summary>
        [TestMethod]
        public void TestPutExecuteTemplateOnline()
        {           
            using (var file = System.IO.File.OpenRead(Common.GetDataDir() + "SampleMailMergeTemplate.docx"))
            {
                using (var data = System.IO.File.OpenRead(Common.GetDataDir() + "SampleExecuteTemplateData.txt"))
                {
                    var request = new PutExecuteTemplateOnlineRequest(file, data);
                    this.wordsApi.PutExecuteTemplateOnline(request);                    
                }
            }
        }

        /// <summary>
        /// A test for PutField
        /// </summary>
        [TestMethod]
        public void TestPutField()
        {
            string name = "SampleWordDocument.docx";
           
            Field body = new Field();
            body.Result = "3";
            body.FieldCode = "{ NUMPAGES }";
            
            body.NodeId = "0.0.3";
            
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PutFieldRequest(name, body, nodePath: "sections/0/paragraphs/0");
            var actual = this.wordsApi.PutField(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PutFormField
        /// </summary>
        [TestMethod]
        public void TestPutFormField()
        {
            string name = "test_multi_pages.docx";            
            string destFileName = "test.docx";
           
            var body = new FormFieldTextInput();

            body.Name = "FullName";
            body.Enabled = true;
            body.CalculateOnExit = true;
            body.StatusText = string.Empty;
            body.TextInputType = FormFieldTextInput.TextInputTypeEnum.Regular;
            body.TextInputDefault = "123";
            body.TextInputFormat = "UPPERCASE";

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PutFormFieldRequest(name, body, nodePath: "sections/0/paragraphs/0", destFileName: destFileName);
            var actual = this.wordsApi.PutFormField(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PutProtectDocument
        /// </summary>
        [TestMethod]
        public void TestPutProtectDocument()
        {
            string name = "test_multi_pages.docx";
            string filename = "test_multi_pages.docx";            
            ProtectionRequest body = new ProtectionRequest(); 

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new PutProtectDocumentRequest(name, body, destFileName: filename);
            var actual = this.wordsApi.PutProtectDocument(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for PutUpdateDocumentProperty
        /// </summary>
        [TestMethod]
        public void TestPutUpdateDocumentProperty()
        {
            string name = "test_multi_pages.docx";
            string propertyName = "Author";
            string filename = "test_multi_pages.docx";          
            DocumentProperty body = new DocumentProperty();
            body.Name = "Author";
            body.Value = "Imran Anwar";

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new CreateOrUpdateDocumentPropertyRequest(name, propertyName, body, destFileName: filename);
            var actual = this.wordsApi.CreateOrUpdateDocumentProperty(request);
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for RejectAllRevisions
        /// </summary>
        [TestMethod]
        public void TestRejectAllRevisions()
        {
            string name = "test_multi_pages.docx";
            string filename = "test_multi_pages.docx";
           
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new RejectAllRevisionsRequest(name, destFileName: filename);
            var actual = this.wordsApi.RejectAllRevisions(request);

            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for Search
        /// </summary>
        [TestMethod]
        public void TestSearch()
        {
            string name = "SampleWordDocument.docx";
            string pattern = "aspose";
         
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new SearchRequest(name, pattern);
            var actual = this.wordsApi.Search(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// A test for UpdateSectionPageSetup
        /// </summary>
        [TestMethod]
        public void TestUpdateSectionPageSetup()
        {
            string name = "test_multi_pages.docx";
            int sectionIndex = 0; 
            
            var body = new PageSetup();
            body.RtlGutter = true;
            body.LeftMargin = 10.0f;
            body.Orientation = PageSetup.OrientationEnum.Landscape;
            body.PaperSize = PageSetup.PaperSizeEnum.A5;

            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));

            var request = new UpdateSectionPageSetupRequest(name, sectionIndex, body);
            var actual = this.wordsApi.UpdateSectionPageSetup(request);
            
            Assert.AreEqual(200, actual.Code);
        }

        /// <summary>
        /// If file does not exist, 400 response should be returned with message "Error while loading file ".
        /// </summary>
        [TestMethod]
        public void TestHandleErrors()
        {
            string name = "noFileWithThisName.docx";
            
            try
            {
                var request = new GetSectionsRequest(name);
                this.wordsApi.GetSections(request);

                Assert.Fail("Excpected exception has not been throwed");
            }
            catch (ApiException apiException)
            {
                Assert.AreEqual(400, apiException.ErrorCode);
                Assert.IsTrue(apiException.Message.StartsWith("Error while loading file 'noFileWithThisName.docx' from storage:"), "Current message: " + apiException.Message);
            }
        }

        /// <summary>
        /// If user set the "Debug" option, request and response should be writed to trace
        /// </summary>
        [TestMethod]
        public void IfUserSetDebugOptionRequestAndErrorsShouldBeWritedToTrace()
        {
            // Arrange
            string name = "test_multi_pages.docx";
            this.storageApi.PutCreate(name, null, null, System.IO.File.ReadAllBytes(Common.GetDataDir() + name));
            var request = new DeleteFieldsRequest(name);
            var api = new WordsApi(ApiKey, ApiSid, AppUrl, true);

            var mockFactory = new MockFactory();
            var traceListenerMock = mockFactory.CreateMock<TraceListener>();
            Trace.Listeners.Add(traceListenerMock.MockObject);

            traceListenerMock.Expects.One.Method(p => p.WriteLine(string.Empty)).With(Is.StringContaining("DELETE: http://api.aspose.cloud/v1.1/words/test_multi_pages.docx/fields"));
            traceListenerMock.Expects.One.Method(p => p.WriteLine(string.Empty)).With(Is.StringContaining("Response 200: OK"));
            traceListenerMock.Expects.One.Method(p => p.WriteLine(string.Empty)).With(Is.StringContaining("{\"Code\":200,\"Status\":\"OK\"}"));

            traceListenerMock.Expects.AtLeastOne.Method(p => p.WriteLine(string.Empty)).With(Is.Anything);

            // Act
            api.DeleteFields(request);
            
            // Assert                    
            mockFactory.VerifyAllExpectationsHaveBeenMet();
        }
    }
}
