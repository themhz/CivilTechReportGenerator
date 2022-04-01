using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using ReportGenerator;
using ReportGenerator.DataSources;
using ReportGenerator.Helpers;
using ReportGenerator_v1.DataSources;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO;
using System.Diagnostics;
using Newtonsoft.Json.Linq;
using System.Configuration;
using System.Text.Json.Nodes;

namespace ReportGenerator_v1.System {

    class DevExpressDocX003 : IReport {

        public RichEditDocumentServer wordProcessor { set; get; }
        public RichEditDocumentServer tempWordProcessor { set; get; }
        public IDataSource datasource { set; get; }
        public String template { set; get; }
        public String generatedfile { set; get; }
        private DocumentRange sourceRange { set; get; }
        private DocumentRange targetRange { set; get; }
        public DevExpressDocX003(RichEditDocumentServer _wordProcessor, IDataSource _datasource) {
            this.wordProcessor = _wordProcessor;
            this.datasource = _datasource;
        }
        public IReport create() {
            Console.WriteLine("creating file report");
            using (this.wordProcessor) {
                this.load();
                this.parse();
                this.save();
                this.openfile();
            }            
            return this;
        }
        public void load() {
            this.wordProcessor.Document.BeginUpdate();
            this.wordProcessor.LoadDocument(this.template);
        }
        public void save() {
            this.wordProcessor.Document.EndUpdate();
            Console.WriteLine("Saving file report");
            this.wordProcessor.SaveDocument(this.generatedfile, DocumentFormat.OpenXml);
            Console.WriteLine("Report saved in :" + this.template);
            Console.WriteLine("file report created");            
        }
        public void openfile() {
            Process.Start(new ProcessStartInfo(this.generatedfile) { UseShellExecute = true });
        }
        public void delete() {
            throw new NotImplementedException();
        }

        public void parse() {                        
            List<Comment> lcomments = new List<Comment>();
            foreach (Comment comment in getAllComments().ToList()) {
                this.parseCommentTypes(comment);
            }
            this.createPageBreak();
        }

        ///<summary>
        ///This function is used to check the type of parsing that will be used        
        ///</summary>
        ///<param name="json">the json string to be parsed as jobject and get the parse type</param>
        ///<param name="comment">the comment object that is passsed</param>
        private void parseCommentTypes(Comment comment, string id = "", RichEditDocumentServer wp = null) {
            string json = this.getCommentText(comment);
            try {
                JObject jsonObject = JObject.Parse(json);
                switch (jsonObject.GetValue("type").ToString()) {
                    case "field":
                        this.parseField(jsonObject, comment, id, wp);
                        break;
                    case "image":
                        this.parseImage(jsonObject, comment, id, wp);
                        break;
                    case "template":
                        this.parseTemplate(jsonObject, comment, id, wp);                       
                        break;
                    case "table":
                        //this.parseTable(jsonObject, comment, id);
                        break;
                }
            } catch (Exception ex) {

            }


        }




        /// <summary>
        /// Μαζεύει όλα τα σχόλια από τον word document
        /// </summary>
        /// <returns>Επιστρέφει τα σχόλια σε μορφή CommentCollection</returns>
        public CommentCollection getAllComments(RichEditDocumentServer wp = null) {
            //Collect all comments and loop through them in order to parse each one of them individually and create the document elemtnts            
            return this.wordProcessor.Document.Comments;
        }

        /// <summary>
        /// Creates a page break on the document to seperate pages between them
        /// </summary>
        public void createPageBreak(RichEditDocumentServer wp = null) {
            Regex r = new Regex("{PBR}");
            this.wordProcessor.Document.ReplaceAll(r, DevExpress.Office.Characters.PageBreak.ToString());
        }


        /// <summary>
        /// gets comment text
        /// </summary>
        /// <param name="comment">the comment object</param>
        /// <returns></returns>
        public string getCommentText(Comment comment) {
            SubDocument doc = comment.BeginUpdate();
            string commentText = doc.GetText(doc.Range).Replace("”", "\"").Replace("{{", "{").Replace("}}", "}");
            comment.EndUpdate(doc);

            return commentText;
        }

        /// <summary>
        /// replaces the targeted range that the comment points to with the text that relates to the name
        /// </summary>
        ///<param name="json">the json string to be parsed as jobject and get the parse type</param>
        ///<param name="comment">the comment object that is passsed</param>
        /// <param name="id"></param>
        public void parseField(JObject jo, Comment comment, string id = "", RichEditDocumentServer wp = null) {
            Console.WriteLine(jo + " is field");
            this.replaceRangeWithNewText(comment.Range, datasource.GetValue(jo.GetValue("name").ToString()).ToString());
        }

    
        /// <summary>
        /// Gets a range and replaces it with some text
        /// </summary>
        /// <param name="sourceRange">The range as DocumentRange object</param>
        /// <param name="targetText">some string</param>
        public void replaceRangeWithNewText(DocumentRange sourceRange, string targetText, RichEditDocumentServer wp = null) {
            //this.mainWordProcessor.Document.BeginUpdate();
            this.targetRange = sourceRange;
            if (this.targetRange != null)
                this.wordProcessor.Document.Replace(targetRange, targetText);
        }


        /// <summary>
        /// Gets the image and prints in the document. Notice that the image is in binary format not an actual jpg and its embedid in the xml
        /// </summary>
        /// <param name="jo">the json object</param>
        /// <param name="comment">the comment object</param>
        /// <param name="id">the id as primary key</param>
        public void parseImage(JObject jo, Comment comment, string id = "", RichEditDocumentServer wp = null) {
            Console.WriteLine(jo + " is image");
            XmlNodeList DetailList = ((Xml)datasource).getList("PageA", "ID", id);
            this.replaceTextWithImage(comment.Range, DetailList[0]["Image"].InnerText, wp);
        }

        /// <summary>
        /// Gets the image as 
        /// </summary>
        /// <param name="sourceRange">the document range that will be replaced by the image</param>
        /// <param name="imageInText">the text of the image withing the document</param>       
        public void replaceTextWithImage(DocumentRange sourceRange, string imageInText, RichEditDocumentServer wp = null) {            
            
            this.targetRange = sourceRange;
            if (this.targetRange != null) {

                byte[] bytes = Convert.FromBase64String(this.datasource.GetValue(imageInText).ToString());
                bytes = ImageResizer.resize(bytes, 700, 700);
                using (MemoryStream ms = new MemoryStream(bytes)) {
                    DocumentImageSource image = DocumentImageSource.FromStream(ms);

                    if (wp == null) {
                        this.wordProcessor.Document.Unit = DevExpress.Office.DocumentUnit.Inch;
                        this.wordProcessor.Document.Images.Insert(this.targetRange.Start, image);
                    } else {
                        wp.Document.Unit = DevExpress.Office.DocumentUnit.Inch;
                        wp.Document.Images.Insert(this.targetRange.Start, image);
                    }
                    
                }
                //this.delete();
            }            
        }

        public void replaceTextWithNewText(string findtext, string replacewithtext, RichEditDocumentServer wp = null) {
            //this.mainWordProcessor.Document.BeginUpdate();
            this.targetRange = this.getTextRange(findtext);
            if (this.targetRange != null) {
                if(wp == null)
                    this.wordProcessor.Document.Replace(targetRange, replacewithtext);
                else
                    wp.Document.Replace(targetRange, replacewithtext);
            }   
                
        }


        /// <summary>
        /// Parses the template in the comment
        /// Finds the looptable attribute in the json comment and loops so it prints the table as many times as it is found in the xml
        /// else it replaces with a template
        /// </summary>
        /// <param name="jo">the json object</param>
        /// <param name="comment">the comment object</param>
        /// <param name="id">the id as primary key</param>
        public void parseTemplate(JObject jo, Comment comment, string id = "", RichEditDocumentServer wp = null) {
            Console.WriteLine(jo + " is template");
            //Repeate for each table
            //if (jo.ContainsKey("table")) {
            //    //Φέρνει μια λίστα από πίνακες και για κάθε πίνακα φέρνει το κλειδί 
            //    XmlNodeList tables = ((Xml)datasource).getList(jo.GetValue("table").ToString(), jo.GetValue("foreignKey").ToString(), id);
            //    foreach (XmlNode table in tables) {
            //        this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), table[jo.GetValue("id").ToString()].InnerText, table[jo.GetValue("foreignKey").ToString()].InnerText);
            //    }
            //} else {
            //    this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), id);
            //}

            this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), id + 1);
            this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), id + 2);
            this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), id + 3);
            this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), id + 4);
            this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), id + 5);
            this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), id + 6);
            this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), id + 7);
            this.replaceTextWithTemplate(comment, jo.GetValue("name").ToString(), id + 8);
        }

        private int countering = 0;
        private void replaceTextWithTemplate(Comment comment, string file, string id, string foreightKey = "", RichEditDocumentServer wp = null) {
            string documentTemplate = Path.Combine(ConfigurationManager.AppSettings["templates"] + file);

            //if(countering < 4) {
                //create child wordprocessor
                using (RichEditDocumentServer childWordPrecessor = new RichEditDocumentServer()) {
                    //load document to child wordprocessor
                    childWordPrecessor.LoadDocumentTemplate(documentTemplate);

                //loop through child processor comments
                //foreach (Comment c in childWordPrecessor.Document.Comments) {
                //    SubDocument doc = c.BeginUpdate();
                //    //String field = doc.GetText(doc.Range).Replace("”", "\"").Replace("“", "\"");
                //    //if (field != "") {
                //        //this.parseCommentTypes(c, id, childWordPrecessor);
                //    //}
                //    doc.EndUpdate();
                //}
                //get the main wordprocessor back

                    
                    this.wordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}").Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
                    
                    //this.replaceTextWithNewText("Διατομή", id);
                }
                
            //}
            
            countering++;
        }

        /// <summary>
        /// Searches for a text in a document and gets its range. The search happens from top to bottom
        /// </summary>
        /// <param name="search">the text to search</param>
        /// <returns></returns>
        public DocumentRange getTextRange(string search, RichEditDocumentServer wp = null) {
            try {
                Regex myRegEx = new Regex(search);
                if(wp==null)
                    return this.wordProcessor.Document.FindAll(myRegEx).First();
                else
                    return wp.Document.FindAll(myRegEx).First();

            } catch (Exception ex) {
                return null;
            }
        }


        public void test() {            
            using (RichEditDocumentServer parentWordProcessor = new RichEditDocumentServer()) {
                parentWordProcessor.Document.BeginUpdate();
                parentWordProcessor.LoadDocument("..\\..\\DataSources\\files\\report_template2.docx");

                //create child wordprocessor
                string documentTemplate = Path.Combine("..\\..\\DataSources\\files\\templates\\template_part4.docx");
                using (RichEditDocumentServer childWordPrecessor = new RichEditDocumentServer()) {
                    //load document to child wordprocessor
                    childWordPrecessor.LoadDocumentTemplate(documentTemplate);

                    

                    //Appears ok
                    parentWordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}", parentWordProcessor).Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
                    parentWordProcessor.Document.EndUpdate();
                    //Appears ok
                    parentWordProcessor.Document.BeginUpdate();
                    parentWordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}", parentWordProcessor).Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
                    parentWordProcessor.Document.EndUpdate();
                    //Doesnt Appears
                    parentWordProcessor.Document.BeginUpdate();
                    parentWordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}", parentWordProcessor).Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
                    parentWordProcessor.Document.EndUpdate();

                    parentWordProcessor.Document.BeginUpdate();
                    parentWordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}", parentWordProcessor).Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
                    parentWordProcessor.Document.EndUpdate();

                    parentWordProcessor.Document.BeginUpdate();
                    parentWordProcessor.Document.InsertDocumentContent(this.getTextRange("{{Test}}", parentWordProcessor).Start, childWordPrecessor.Document.Range, InsertOptions.KeepSourceFormatting);
                    parentWordProcessor.Document.EndUpdate();

                }

                parentWordProcessor.Document.EndUpdate();
                parentWordProcessor.SaveDocument(this.generatedfile, DocumentFormat.OpenXml);
                this.openfile();
            }
        }
    }
}
