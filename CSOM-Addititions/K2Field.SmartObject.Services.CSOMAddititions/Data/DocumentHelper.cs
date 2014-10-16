using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSOM = Microsoft.SharePoint.Client;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using Attributes = SourceCode.SmartObjects.Services.ServiceSDK.Attributes;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using IO = System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using SourceCode.SmartObjects.Services.ServiceSDK;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace K2Field.SmartObject.Services.CSOMAddititions.Data
{
    [Attributes.ServiceObject("DocumentHelper", "Document Helper", "Service Object to assist in the creation of new documents")]
    class DocumentHelper
    {
        public DocumentHelper()
        {        }

        public ServiceConfiguration ServiceConfiguration
        { get; set;}

        #region Private Variables
        private int position;
        private int bufferSize;
        private string _ControlName = "";
        private string _ControlNames = "";
        private string _ControlValue = "";
        private string _ControlValues = "";
        private string _SplitCharacter = "";
        private string _ImgLocation = "";
        private string _ImgUrl = "";
        private string _SharePointURL = "";
        private string _LibraryName = "";
        private string _FolderName = "";
        private string _NewDocName = "";
        private string _FullDocumentURL = "";
        private bool _OverwriteExisting = false;
        private string _SrcSharePointURL = "";
        private string _SrcLibraryName = "";
        private string _SrcFolderName = "";
        private string _SrcDocName = "";
        private CSOM.List _WorkingList = null;

        #endregion

        #region Properties

        [Attributes.Property("ControlName", SoType.Text, "Control Name", "Control Name")]
        public string ControlName
        {
            get { return _ControlName; }
            set { _ControlName = value; }
        }

        [Attributes.Property("ControlValue", SoType.Text, "Control Value", "Control Value")]
        public string ControlValue
        {
            get { return _ControlValue; }
            set { _ControlValue = value; }
        }

        [Attributes.Property("ControlNames", SoType.Text, "Control Names", "Control Names")]
        public string ControlNames
        {
            get { return _ControlNames; }
            set { _ControlNames = value; }
        }

        [Attributes.Property("ControlValues", SoType.Text, "Control Values", "Control Values")]
        public string ControlValues
        {
            get { return _ControlValues; }
            set { _ControlValues = value; }
        }

        [Attributes.Property("SplitCharacter", SoType.Text, "Split Character", "Split Character")]
        public string SplitCharacter
        {
            get { return _SplitCharacter; }
            set { _SplitCharacter = value; }
        }

        [Attributes.Property("ImageLocation", SoType.Text, "Image Location", "Local Image Location")]
        public string ImageLocation
        {
            get { return _ImgLocation; }
            set { _ImgLocation = value; }
        }
       

        [Attributes.Property("SharePointURL", SoType.Text, "SharePoint URL", "SharePoint URL")]
        public string SharePointURL
        {
            get { return _SharePointURL; }
            set { _SharePointURL = value; }
        }

        [Attributes.Property("LibraryName", SoType.Text, "Library Name", "Library Name")]
        public string LibraryName
        {
            get { return _LibraryName; }
            set { _LibraryName = value; }
        }

        [Attributes.Property("FolderName", SoType.Text, "Folder Name", "Folder Name")]
        public string FolderName
        {
            get { return _FolderName; }
            set { _FolderName = value; }
        }

        [Attributes.Property("DocumentName", SoType.Text, "Document Name", "Document Name")]
        public string DocumentName
        {
            get { return _NewDocName; }
            set { _NewDocName = value; }
        }

        [Attributes.Property("SrcSharePointURL", SoType.Text, "Source Document SharePoint URL", "Source Document SharePoint URL")]
        public string SrcSharePointURL
        {
            get { return _SrcSharePointURL; }
            set { _SrcSharePointURL = value; }
        }

        [Attributes.Property("SrcLibraryName", SoType.Text, "Source Document Library Name", "Source Document Library Name")]
        public string SrcLibraryName
        {
            get { return _SrcLibraryName; }
            set { _SrcLibraryName = value; }
        }

        [Attributes.Property("SrcFolderName", SoType.Text, "Source Document Folder Name", "Source Document Folder Name")]
        public string SrcFolderName
        {
            get { return _SrcFolderName; }
            set { _SrcFolderName = value; }
        }

        [Attributes.Property("SrcDocumentName", SoType.Text, "Source Document Name", "Source Document Name")]
        public string SrcDocumentName
        {
            get { return _SrcDocName; }
            set { _SrcDocName = value; }
        }

        [Attributes.Property("FullDocumentURL", SoType.Text, "Full Document URL", "Full Document URL")]
        public string FullDocumentURL
        {
            get { return _FullDocumentURL; }
            set { _FullDocumentURL = value; }
        }

        [Attributes.Property("OverwriteExisting", SoType.YesNo, "Overwrite Existing", "Overwrite item if it exist")]
        public bool OverwriteExisting
        {
            get { return _OverwriteExisting; }
            set { _OverwriteExisting = value; }
        }

        #endregion

        #region Private Methods

        private void UpdateElements(Document document, string ContentControlName, string SourceText)
        {
            // Get all controls by name
            var sdtRuns = document.Descendants<SdtRun>()
                           .Where(run => run.SdtProperties.GetFirstChild<Tag>().Val.Value == this.ControlName);

            // Iterate the list
            foreach (SdtRun sdtRun in sdtRuns)
            {
                // Array used for line breaks
                string[] srcArr = SourceText.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

                // Set control's value
                sdtRun.Descendants<Text>().First().Text = this.ControlValue;

                Text xText = sdtRun.Descendants<Text>().FirstOrDefault();
                if (xText == null)
                {
                    throw new NullReferenceException("Object Reference: Unable to find xText object");
                }

                xText.Text = srcArr[0];
                Text newRow;
                
                // Handle line breaks
                for (int i = 1; i < srcArr.Length; i++)
                {
                    xText.Parent.AppendChild(new Break());
                    newRow = (Text)xText.Clone();
                    newRow.Text = srcArr[i];
                    xText.Parent.AppendChild(newRow);
                }
            }

            //foreach (SdtElement sdtElement in document.Descendants<SdtElement>().ToList())
            //{
            //    SdtAlias alias = sdtElement.Descendants<SdtAlias>().FirstOrDefault();

            //    if (alias != null)
            //    {
            //        string elementName = alias.Val.Value;
            //        var element = sdtElement;
            //        Text xText = element.Descendants<Text>().FirstOrDefault();
            //        if (xText == null)
            //        {
            //            throw new NullReferenceException("Object Reference: Unable to find xText object");
            //        }

            //        string[] srcArr = SourceText.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

            //        if (elementName == ContentControlName)
            //        {
            //            xText.Text = srcArr[0];
            //            Text newRow;

            //            for (int i = 1; i < srcArr.Length; i++)
            //            {
            //                xText.Parent.AppendChild(new Break());
            //                newRow = (Text)xText.Clone();
            //                newRow.Text = srcArr[i];
            //                xText.Parent.AppendChild(newRow);
            //            }
            //        }
            //    }
            //}
        }

        private CSOM.ClientContext InitClientContext(string sharePointURL)
        {
            // Instanciate client context
            CSOM.ClientContext clientContext = new CSOM.ClientContext(sharePointURL);

            // Switch to static credentials if provided
            if (this.ServiceConfiguration.ServiceAuthentication.AuthenticationMode == AuthenticationMode.Static)
            {
                var username = this.ServiceConfiguration.ServiceAuthentication.UserName; // "justine@pwcec.onmicrosoft.com";
                var password = this.ServiceConfiguration.ServiceAuthentication.Password;
                var securePassword = new System.Security.SecureString();

                foreach (char c in password.ToCharArray())
                    securePassword.AppendChar(c);

                clientContext.Credentials = new CSOM.SharePointOnlineCredentials(username, securePassword);
            }

            clientContext.PendingRequest.RequestExecutor.WebRequest.KeepAlive = false;

            return clientContext;
        }

        private void SetWorkingList(CSOM.ClientContext ctx, string libName)
        {
            // Get current web and load
            CSOM.Web site = ctx.Web;
            ctx.Load(site);

            // Trim Library name
            if (libName.EndsWith("/"))
                libName = libName.TrimEnd(new char[] { '/' });

            // Set working list
            _WorkingList = site.Lists.GetByTitle(libName);
        }

        private IO.MemoryStream GetDocumentStreamFromSP(CSOM.ClientContext clientCntxt, string libraryName, string folderName, string documentName )
        {
            // Counters for loading binary data
            position = 1;
            bufferSize = 200000;

            IO.MemoryStream imageMemStream = null;

            // Set working list
            SetWorkingList(clientCntxt, libraryName);

            CSOM.CamlQuery query = new CSOM.CamlQuery();
            // Query to test if the document exists
            query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/>" +
                "<Value Type='Text'>" + documentName + "</Value></Eq></Where></Query><RowLimit>2</RowLimit></View>";

            // Handle Folders if used
            if (folderName != null && folderName.Length > 0)
            {
                clientCntxt.Load(_WorkingList,
               list => list.RootFolder.ServerRelativeUrl);
                clientCntxt.ExecuteQuery();

                query.FolderServerRelativeUrl = _WorkingList.RootFolder.ServerRelativeUrl + "/" + folderName;
            }

            // Run Query
            CSOM.ListItemCollection collListItem = _WorkingList.GetItems(query);
            clientCntxt.Load(collListItem);
            clientCntxt.ExecuteQuery();
            
            string source = string.Empty;
            // Set source document location
            if (collListItem.Count == 1)
            {
                source = collListItem[0].FieldValues["FileRef"].ToString();
            }
            else
            {
                throw new Exception(String.Format("Document {0} not found in Library {1}", documentName, libraryName));
            }

            // Open the file and read it to memory
            using (CSOM.FileInformation fileInfo = CSOM.File.OpenBinaryDirect(clientCntxt, source))
            {
                Byte[] readBuffer = new Byte[bufferSize];

                imageMemStream = new IO.MemoryStream();
                while (position > 0)
                {
                    position = fileInfo.Stream.Read(readBuffer, 0, bufferSize);
                    imageMemStream.Write(readBuffer, 0, position);
                    readBuffer = new Byte[bufferSize];
                }

                fileInfo.Stream.Flush();
                imageMemStream.Flush();
            }

            // Return file in memory stream
            return imageMemStream;
        }

        private IO.MemoryStream GetTemplateStreamFromSP(CSOM.ClientContext clientContxt, string libraryName)
        {
            // Counters for loadin binary data
            position = 1;
            bufferSize = 200000;

            // Set working list
            SetWorkingList(clientContxt, libraryName);

            // Get template URL
            clientContxt.Load(_WorkingList,
                list => list.DocumentTemplateUrl);
            clientContxt.ExecuteQuery();

            // Set relative URL
            var source = _WorkingList.DocumentTemplateUrl;

            IO.MemoryStream docMemStream;
            // Open the template and load to memory
            using (CSOM.FileInformation fileInfo = CSOM.File.OpenBinaryDirect(clientContxt, source))
            {
                Byte[] readBuffer = new Byte[bufferSize];

                docMemStream = new IO.MemoryStream();
                while (position > 0)
                {
                    position = fileInfo.Stream.Read(readBuffer, 0, bufferSize);
                    docMemStream.Write(readBuffer, 0, position);
                    readBuffer = new Byte[bufferSize];
                }

                fileInfo.Stream.Flush();
                docMemStream.Flush();

                // Reset memory stream position
                docMemStream.Position = 0L;
            }

            return docMemStream;
        }

        private string PutDocumentStreamToSP(CSOM.ClientContext clientContxt, IO.MemoryStream docMemStream, string folderName, string docName, bool overwriteExisting)
        {
            string sUploadedDocURL = string.Empty;

            // Ensure memory stream is at the beggining of the file
            docMemStream.Position = 0L;
            
            // Load root URL
            clientContxt.Load(_WorkingList,
                list => list.RootFolder.ServerRelativeUrl);
            clientContxt.ExecuteQuery();

            string sRelativeURL = _WorkingList.RootFolder.ServerRelativeUrl;

            CSOM.CamlQuery query = new CSOM.CamlQuery();

            // Check if a subfolder has been specified
            if (this.FolderName != null && this.FolderName.Length > 0)
            {
                sRelativeURL += "/" + this.FolderName;
                query.FolderServerRelativeUrl = sRelativeURL;
            }

            try
            {
                // Save binary file
                CSOM.File.SaveBinaryDirect(clientContxt, sRelativeURL + "/" + docName, docMemStream, overwriteExisting);
            }
            catch (Exception ex)
            {
                // Check for document locking
                if (ex.Message.CompareTo("The underlying connection was closed: The connection was closed unexpectedly.") == 0)
                {
                    throw new Exception("Unable to to upload the document to SharePoint, most likely caused by the document being locked: " + ex.Message, ex);
                }
                else
                {
                    throw ex;
                }
            }

            // Query to check if the file has been uploaded
            query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='FileLeafRef'/>" +
                "<Value Type='Text'>" + docName + "</Value></Eq></Where></Query><RowLimit>2</RowLimit></View>";

            CSOM.ListItemCollection collListItem = _WorkingList.GetItems(query);

            clientContxt.Load(collListItem);

            clientContxt.ExecuteQuery();

            if (collListItem.Count == 1)
            {
                sUploadedDocURL = clientContxt.Url + collListItem[0].FieldValues["FileRef"].ToString();
            }
            else
            {
                throw new Exception("Upload failed: Unable to retrieve new document");
            }

            // Return document's URL
            return sUploadedDocURL;
        }

        private void ReplaceTextContentInDocumentv1(Document doc, string controlName, string controlValue)
        {
            string[] srcArr = controlValue.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
            string elementName;
            SdtElement element;
            Text xText;

            // Loop through nodes and update all with name provided
            foreach (SdtElement sdtElement in doc.Descendants<SdtElement>().ToList())
            {
                SdtAlias alias = sdtElement.Descendants<SdtAlias>().FirstOrDefault();

                if (alias != null)
                {
                    elementName = alias.Val.Value;
                    element = sdtElement;
                    xText = element.Descendants<Text>().FirstOrDefault();

                    // Valid Control found
                    if (xText != null && elementName == controlName)
                    {
                        // Add first row
                        xText.Text = srcArr[0];
                        Text newRow;

                        // Add additional rows
                        for (int i = 1; i < srcArr.Length; i++)
                        {
                            xText.Parent.AppendChild(new Break());
                            newRow = (Text)xText.Clone();
                            newRow.Text = srcArr[i];
                            xText.Parent.AppendChild(newRow);
                        }
                    }
                }
            }
        }

        private void ReplaceTextContentInDocument(Document doc, string controlName, string controlValue)
        {
            string[] srcArr = controlValue.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
            string elementName;
            SdtElement element;
            Text xText;

            // Loop through nodes and update all with name provided
            List<SdtElement> controlBlocks = doc.Body
                .Descendants<SdtElement>()
                    .Where
                    (r =>
                        r.SdtProperties.GetFirstChild<Tag>().Val == controlName
                    ).ToList<SdtElement>();

            // Replace image for all
            foreach (SdtElement controlBlock in controlBlocks)
            {
                SdtAlias alias = controlBlock.Descendants<SdtAlias>().FirstOrDefault();

                if (alias != null)
                {
                    elementName = alias.Val.Value;
                    element = controlBlock;
                    xText = element.Descendants<Text>().FirstOrDefault();

                    // Valid Control found
                    if (xText != null && elementName == controlName)
                    {
                        // Add first row
                        xText.Text = srcArr[0];
                        Text newRow;

                        // Add additional rows
                        for (int i = 1; i < srcArr.Length; i++)
                        {
                            xText.Parent.AppendChild(new Break());
                            newRow = (Text)xText.Clone();
                            newRow.Text = srcArr[i];
                            xText.Parent.AppendChild(newRow);
                        }
                    }
                }
            }
        }

        private void ReplaceTextContentInDocumentOrg(Document doc, string controlName, string controlValue)
        {
            // Iterate items with specified name and replace their content
            var sdtRuns = doc.Descendants<SdtRun>()
                       .Where(run => run.SdtProperties.GetFirstChild<Tag>().Val.Value == controlName);

            foreach (SdtRun sdtRun in sdtRuns)
            {
                sdtRun.Descendants<Text>().First().Text = controlValue;
            }
        }

        private IO.MemoryStream GetDocumentStreamFromFileSystem(string filePath)
        {
            // Read a file from file system to memory stream
            IO.MemoryStream memStream;
            using (IO.FileStream fileStream = IO.File.OpenRead(filePath))
            {
                memStream = new IO.MemoryStream();
                memStream.SetLength(fileStream.Length);
                fileStream.Read(memStream.GetBuffer(), 0, (int)fileStream.Length);
                fileStream.Flush();
                memStream.Flush();
            }
            // Reset stream to start from beginning
            memStream.Position = 0;
            return memStream;
        }

        private WordprocessingDocument UpdateImageContentControl(IO.MemoryStream documentStream, IO.MemoryStream imageStream)
        {
            // Create a editable document object from memory stream
            WordprocessingDocument finalDoc = WordprocessingDocument.Open(documentStream, true);

            // Get the relevant parts of the doc
            Body body = finalDoc.MainDocumentPart.Document.Body;

            MainDocumentPart mainPart = finalDoc.MainDocumentPart;
            Document document = mainPart.Document;

            // Get list of controls that has current name
            List<SdtElement> controlBlocks = mainPart.Document.Body
               .Descendants<SdtElement>()
                   .Where
                   (r =>
                       r.SdtProperties.GetFirstChild<Tag>().Val == this.ControlName
                   ).ToList<SdtElement>();

            // Replace image for all
            foreach (SdtElement controlBlock in controlBlocks)
            {

                // Find the Blip element of the content control.
                A.Blip blip = controlBlock.Descendants<A.Blip>().FirstOrDefault();

                // Remove all placeholders for the control, else the Insert Image icon persists 
                foreach (var rem in controlBlock.Descendants<ShowingPlaceholder>())
                {
                    rem.Remove();
                }

                // Get the image portion of the control
                ImagePart imagePart = (ImagePart)mainPart.GetPartById(blip.Embed);

                System.Drawing.Bitmap image = new System.Drawing.Bitmap(imageStream);

                // Write the source image to the destination
                using (IO.BinaryWriter writer = new IO.BinaryWriter(imagePart.GetStream()))
                {
                    writer.Write(imageStream.ToArray());
                    imageStream.Flush();
                    writer.Close();
                }

                // Adjust image to its proper size
                DW.Inline inline = controlBlock.Descendants<DW.Inline>().FirstOrDefault();
                // 9525 = pixels to points
                inline.Extent.Cy = image.Size.Height * 9525;
                inline.Extent.Cx = image.Size.Width * 9525;
                PIC.Picture pic = inline
                    .Descendants<PIC.Picture>().FirstOrDefault();
                pic.ShapeProperties.Transform2D.Extents.Cy
                    = image.Size.Height * 9525;
                pic.ShapeProperties.Transform2D.Extents.Cx
                    = image.Size.Width * 9525;
            }

            // Save changes made
            document.Save();

            return finalDoc;
        }

        #endregion

        #region Methods

        [Attributes.Method("CreateNewDocument", MethodType.Read, "Create New Document", "Create a New Document using the template in the Library specified",
           new string[] { "SharePointURL", "LibraryName", "DocumentName" },
           new string[] { "SharePointURL", "LibraryName", "FolderName", "DocumentName", "OverwriteExisting" }, 
           new string[] { "FullDocumentURL" })] 
        public DocumentHelper CreateNewDocument()
        {
            // Prep context object
            CSOM.ClientContext clientContext = InitClientContext(this.SharePointURL);

            try
            {
                // Get document library template
                using (IO.MemoryStream templateMemStream = GetTemplateStreamFromSP(clientContext, this.LibraryName))
                {
                    // This test always seem to fail for O365, ignore for now
                    //using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(templateMemStream, false))
                    //{
                    //    // Validate whether the new document is good to go
                    //    OpenXmlValidator validator = new OpenXmlValidator();
                    //    var errors = validator.Validate(wordDoc);
                    //    if (errors.Count() == 0)
                    //    {
                    // No errors downloading, save the new document

                    // Instanciate editable document object
                    using (WordprocessingDocument finalDoc = WordprocessingDocument.Open(templateMemStream, true))
                    {
                        // Change the type from template to final document
                        finalDoc.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                        MainDocumentPart mainPart = finalDoc.MainDocumentPart;
                        mainPart.DocumentSettingsPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate",
                           new Uri(this.DocumentName, UriKind.Relative));

                        mainPart.Document.Save();
                    }
                    //}
                    //else
                    //{
                    //    StringBuilder sInnerExc = new StringBuilder();

                    //    foreach (var error in errors)
                    //    {
                    //        sInnerExc.AppendLine(String.Format("Error description: {0}", error.Description));
                    //        sInnerExc.AppendLine(String.Format("Content type of part with error: {0}",
                    //           error.Part.ContentType));
                    //        sInnerExc.AppendLine(String.Format("Location of error: {0}", error.Path.XPath));
                    //    }

                    //    throw new Exception("Unable to create a valid document from the template", new Exception(sInnerExc.ToString()));
                    //}

                    //}

                    // Upload final document back to document library
                    this.FullDocumentURL = PutDocumentStreamToSP(clientContext, templateMemStream, this.FolderName, this.DocumentName, this.OverwriteExisting);
                }
            }
            finally
            {
                // Clean up connection context
                clientContext.Dispose();
            }

            return this;
        }

        [Attributes.Method("UpdateDocumentFieldValue", MethodType.Read, "Update a Document Field", "Update a Document Field's Value",
           new string[] { "SharePointURL", "LibraryName", "DocumentName", "ControlName", "ControlValue" },
           new string[] { "SharePointURL", "LibraryName","FolderName", "DocumentName", "ControlName", "ControlValue" },
           new string[] { "FullDocumentURL" })]
        public DocumentHelper UpdateDocumentFieldValue()
        {
            // Prep context object
            CSOM.ClientContext clientContext = InitClientContext(this.SharePointURL);

            try
            {
                // Read original document from SP
                using (IO.MemoryStream documentMemStream = GetDocumentStreamFromSP(clientContext, this.LibraryName, this.FolderName, this.DocumentName))
                {
                    documentMemStream.Position = 0L;

                    // Create editable document instance
                    using (WordprocessingDocument finalDoc = WordprocessingDocument.Open(documentMemStream, true))
                    {
                        Body body = finalDoc.MainDocumentPart.Document.Body;

                        MainDocumentPart mainPart = finalDoc.MainDocumentPart;
                        Document document = mainPart.Document;

                        // Replace control values
                        ReplaceTextContentInDocument(document, this.ControlName, this.ControlValue);

                        // Save
                        document.Save();
                    }

                    documentMemStream.Position = 0L;

                    // Write updated document back to SP
                    this.FullDocumentURL = PutDocumentStreamToSP(clientContext, documentMemStream, this.FolderName, this.DocumentName, true);
                }
            }
            finally
            {
                // Clean up connection context
                clientContext.Dispose();
            }

            return this;
        }

        [Attributes.Method("UpdateMultipleDocumentFields", MethodType.Read, "Update Multiple Document Fields", "Update Multiple Document Fields",
         new string[] { "SharePointURL", "LibraryName", "DocumentName", "ControlNames", "ControlValues", "SplitCharacter"  },
         new string[] { "SharePointURL", "LibraryName", "FolderName", "DocumentName", "ControlNames", "ControlValues", "SplitCharacter" },
         new string[] { "FullDocumentURL" })]
        public DocumentHelper UpdateMultipleDocumentFields()
        {
            // Prep context object
            CSOM.ClientContext clientContext = InitClientContext(this.SharePointURL);

            try
            {
                // Read original document from SP
                using (IO.MemoryStream documentMemStream = GetDocumentStreamFromSP(clientContext, this.LibraryName, this.FolderName, this.DocumentName))
                {
                    documentMemStream.Position = 0L;

                    // Create editable document instance
                    using (WordprocessingDocument finalDoc = WordprocessingDocument.Open(documentMemStream, true))
                    {
                        Body body = finalDoc.MainDocumentPart.Document.Body;

                        MainDocumentPart mainPart = finalDoc.MainDocumentPart;
                        Document document = mainPart.Document;

                        // Split inputs
                        string[] ControlNames = this.ControlNames.Split(new string[] { this.SplitCharacter }, StringSplitOptions.RemoveEmptyEntries);
                        string[] ControlValues = this.ControlValues.Split(new string[] { this.SplitCharacter }, StringSplitOptions.None);

                        // Loop the collection
                        for (int i = 0; i < ControlNames.Length; i++)
                        {
                            // Check for potentially empty entries
                            if (ControlNames[i].Length == 0)
                                continue;

                            // Replace control values
                            ReplaceTextContentInDocument(document, ControlNames[i], ControlValues[i]);
                        }

                        // Save
                        document.Save();
                    }

                    documentMemStream.Position = 0L;

                    // Write updated document back to SP
                    this.FullDocumentURL = PutDocumentStreamToSP(clientContext, documentMemStream, this.FolderName, this.DocumentName, true);
                }
            }
            finally
            {
                // Clean up connection context
                clientContext.Dispose();
            }

            return this;
        }

        [Attributes.Method("UpdateImageContentControlFileSystem", MethodType.Read, "Update Image Content Control from File System", "Update Image Content Control from File System",
          new string[] { "SharePointURL", "LibraryName", "DocumentName", "ControlName", "ImageLocation" },
          new string[] { "SharePointURL", "LibraryName", "FolderName", "DocumentName", "ControlName", "ImageLocation" },
          new string[] { "FullDocumentURL" })]
        public DocumentHelper UpdateImageContentControlFileSystem()
        {
            // Init the CSOM connection
            CSOM.ClientContext clientContext = InitClientContext(this.SharePointURL);

            try
            {
                // Load the source document
                using (IO.MemoryStream documentMemStream = GetDocumentStreamFromSP(clientContext, this.LibraryName, this.FolderName, this.DocumentName))
                {
                    // Ensure the stream is at the beginning
                    documentMemStream.Position = 0L;

                    // Get the image to be inserted and update the image control
                    using (IO.MemoryStream imageMemStream = GetDocumentStreamFromFileSystem(this.ImageLocation))
                    {
                        UpdateImageContentControl(documentMemStream, imageMemStream);

                        // Reset stream location
                        documentMemStream.Position = 0L;

                        // Upload the updated document
                        this.FullDocumentURL = PutDocumentStreamToSP(clientContext, documentMemStream, this.FolderName, this.DocumentName, true);
                    }
                }
            }
            finally
            {
                // Clean up connection context
                clientContext.Dispose();
            }

            return this;
        }

        [Attributes.Method("UpdateImageContentControlSP", MethodType.Read, "Update Image Content Control from SharePoint", "Update Image Content Control from SharePoint",
         new string[] { "SharePointURL", "LibraryName", "DocumentName", "ControlName", "ImageURL", "SrcSharePointURL", "SrcLibraryName", "SrcDocumentName" },
         new string[] { "SharePointURL", "LibraryName", "FolderName", "DocumentName", "ControlName", "SrcSharePointURL", "SrcLibraryName", "SrcFolderName", "SrcDocumentName" },
         new string[] { "FullDocumentURL" })]
        public DocumentHelper UpdateImageContentControlSP()
        {
            // Prep context objects
            // Master document's context
            CSOM.ClientContext clientContext = InitClientContext(this.SharePointURL);
            // Image to be used as replacement's context
            CSOM.ClientContext srcClientCntxt = InitClientContext(this.SrcSharePointURL);
            
            try
            {
                // Load the source document
                using (IO.MemoryStream documentMemStream = GetDocumentStreamFromSP(clientContext, this.LibraryName, this.FolderName, this.DocumentName))
                {
                    // Ensure the stream is at the beginning
                    documentMemStream.Position = 0L;

                    // Get the image to be inserted and update the image control
                    using (IO.MemoryStream imageMemStream = GetDocumentStreamFromSP(srcClientCntxt, this.SrcLibraryName, this.SrcFolderName, this.SrcDocumentName))
                    {
                        // Reset working lib after retrieval
                        SetWorkingList(clientContext, this.LibraryName);

                        // Update image content control using both streams
                        UpdateImageContentControl(documentMemStream, imageMemStream);

                        // Reset stream location
                        documentMemStream.Position = 0L;

                        // Upload the updated document
                        this.FullDocumentURL = PutDocumentStreamToSP(clientContext, documentMemStream, this.FolderName, this.DocumentName, true);
                    }
                }
            }
            finally
            {
                // Clean up connection contexts
                clientContext.Dispose();
                srcClientCntxt.Dispose();
            }

            return this;
        }

        #endregion
    }
}
