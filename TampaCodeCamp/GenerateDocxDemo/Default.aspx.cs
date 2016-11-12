using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Text;

namespace GenerateDocxDemo
{
    public partial class _Default : System.Web.UI.Page
    {
        #region State Management
        /// <summary>
        /// The document to that will be included in our generated document.
        /// </summary>
        private byte[] UploadedDocx
        {
            get
            {
                if (ViewState["UploadedDocx"] != null)
                {
                    return (byte[])ViewState["UploadedDocx"];
                }

                return null;
            }
            set
            {
                ViewState["UploadedDocx"] = value;
            }
        }

        /// <summary>
        /// The HTML content that will be included in our generated document.
        /// </summary>
        private string HTMLContent
        {
            get
            {
                if (ViewState["HTMLContent"] != null)
                {
                    return ViewState["HTMLContent"].ToString();
                }

                return null;
            }
            set
            {
                ViewState["HTMLContent"] = value;
            }

        }
        #endregion

        #region Event Handling
        /// <summary>
        /// On page load if document is to be loaded then load it, and then end the response
        /// to stop any further page processing.
        /// </summary>
        protected void Page_Load(object sender, EventArgs e)
        {
            this.btnMixedContent.Click += new EventHandler(btnMixedContent_Click);
            this.btnExtra.Click += new EventHandler(btnExtra_Click);
        }

        /// <summary>
        /// Generate the Extra Document.
        void btnExtra_Click(object sender, EventArgs e)
        {
            //Generate the document.
            WriteDocumentToResponse("RandomNumbers", BuildExtraDocument());
        }

        /// <summary>
        /// Generate the Mixed Content Document.
        /// </summary>
        void btnMixedContent_Click(object sender, EventArgs e)
        {
            HttpPostedFile file = fileUploadDocx.PostedFile;

            if (file != null && file.ContentType == "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            {
                //Convert the uploaded document to a byte array, then store it can be used in our output.
                byte[] bytes;
                using (BinaryReader br = new BinaryReader(file.InputStream))
                {
                    bytes = br.ReadBytes((int)file.InputStream.Length);
                    UploadedDocx = bytes;
                }

                //Store the html so that it can be used in our output.
                HTMLContent = freeTextoBox.Text;
            }

            //Generate the document.
            WriteDocumentToResponse("MixedContent", BuildMixedContentDocument());
        }
        #endregion
  
        #region Document Generation
        /// <summary>
        /// Build the rich content document.
        /// </summary>
        /// <returns>A byte array, which can later be used to write the document to the Response.</returns>
        byte[] BuildMixedContentDocument()
        {
            DocumentSpecification spec = new DocumentSpecification();
            spec.HTMLContent = HTMLContent;
            spec.DocumentToMerge = UploadedDocx;
            MixedContentDocument doc = new MixedContentDocument(spec);
            return doc.GenerateDocument();
        }

        /// <summary>
        /// Build the extra document.
        /// </summary>
        /// <returns>A byte array, which can later be used to write the document to the Response.</returns>
        byte[] BuildExtraDocument()
        {
            ExtraDocument doc = new ExtraDocument();
            return doc.GenerateDocument(Int32.Parse(txtNumPeople.Text), Int32.Parse(txtNumGiveAways.Text));
        }
        #endregion

        #region Helper Methods
        /// <summary>
        /// Write a generated document to the response.
        /// </summary>
        /// <param name="title">The file name that should be used for the generated docx.</param>
        /// <param name="docBytes">The content of the docx that needs to be generated.</param>
        private void WriteDocumentToResponse(string title, byte[] docBytes)
        {
            System.Web.HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

            //Set the filename and extension in the header response.
            System.Web.HttpContext.Current.Response.AppendHeader("Content-Disposition", string.Format("attachment;filename={0}.docx", title));

            //Write the bytes of the document to the response.
            System.Web.HttpContext.Current.Response.BinaryWrite(docBytes);

            //Stop the response so that there will be no further page processing.
            System.Web.HttpContext.Current.Response.End();
        }
        #endregion
    }
}
