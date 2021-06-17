using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Core;
using System;
using System.Windows.Forms;

namespace test
{
    public partial class ThisAddIn
    {
      
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Word.Application wb;
            wb = this.Application;
            if (wb.Documents.Count > 0)
            {
                String queryResult_StartPosition = String.Empty;
                String queryResult_EndPosition = String.Empty;
                String queryResult_btnName = String.Empty;
                String queryResult_btnText = String.Empty;
                Microsoft.Office.Core.DocumentProperties properties = (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;

                //get doc info about start pos
                if (properties.Cast<DocumentProperty>().Where(c => c.Name == "startPosition").Count() > 0)
                {
                    queryResult_StartPosition = properties["startPosition"].Value;
                }
                else
                {
                    queryResult_StartPosition = String.Empty;
                }
                //get doc info about end pos
                
                if (properties.Cast<DocumentProperty>().Where(c => c.Name == "endPosition").Count() > 0)
                {
                    queryResult_EndPosition = properties["endPosition"].Value;
                }
                else
                {
                    queryResult_EndPosition = String.Empty;
                }
                //get info about btn name
                if (properties.Cast<DocumentProperty>().Where(c => c.Name == "btnName").Count() > 0)
                {
                    queryResult_btnName = properties["btnName"].Value;
                }
                else
                {
                    queryResult_btnName = String.Empty;
                }

                //get info about btn text
                if (properties.Cast<DocumentProperty>().Where(c => c.Name == "btnText").Count() > 0)
                {
                    queryResult_btnText = properties["btnText"].Value;
                }
                else
                {
                    queryResult_btnText = String.Empty;
                }


                Document vstoDocument = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
                Word.Range rng = vstoDocument.Range(queryResult_StartPosition, queryResult_EndPosition);
                Word.Selection selection = this.Application.Selection;
                if (selection != null && rng != null)
                {
                    Button button = new Button();
                    button.Click += new EventHandler(Generatedbtn_Click);
                    button = vstoDocument.Controls.AddButton(rng, 100, 30, queryResult_btnName);        
                    button.Text = queryResult_btnText;
                    button.Name = queryResult_btnName;
                }
            }//end of wb.docment.count
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        internal void DebugStartup()
        {
            //the same code as the ThisAddIn_Startup Function to simulate
            Word.Application wb;
            wb = this.Application;
            if (wb.Documents.Count > 0)
            {
                String queryResult_StartPosition = String.Empty;
                String queryResult_EndPosition = String.Empty;
                String queryResult_btnName = String.Empty;
                String queryResult_btnText = String.Empty;
                Microsoft.Office.Core.DocumentProperties properties = (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;

                //get doc info about start pos
                queryResult_StartPosition = loadInfoInProp("startPosition", properties);
                //get doc info about end pos
                queryResult_EndPosition = loadInfoInProp("endPosition", properties);
                //get info about btn name
                queryResult_btnName = loadInfoInProp("btnName", properties);
                //get info about btn text
                queryResult_btnText = loadInfoInProp("btnText", properties);
     

                Document vstoDocument = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
                Word.Range rng = vstoDocument.Range(queryResult_StartPosition, queryResult_EndPosition);
                Word.Selection selection = this.Application.Selection;
                if (selection != null && rng != null)
                {
                    Button button = new Button();
                    button.Click += new EventHandler(Generatedbtn_Click);
                    button = vstoDocument.Controls.AddButton(rng, 100, 30, queryResult_btnName);
                    button.Click += Generatedbtn_Click;
                    button.Text = queryResult_btnText;
                    button.Name = queryResult_btnName;
                }
            }//end of wb.docment.count
        }

        internal void WhenRibionBtnIsClicked()
        {
            Document vstoDocument = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
            Word.Selection selection = this.Application.Selection;
            if (selection != null && selection.Range != null)
            {
                string name = "myBtn";
                string text = "I am A Generated Button";
                Button button = new Button();
                button.Click += new EventHandler(Generatedbtn_Click);
                button = vstoDocument.Controls.AddButton(selection.Range, 100, 30, name);
                button.Click += Generatedbtn_Click; //for the click function
                button.Text = text;
                button.Name = name;

                /*this part is done so that when the document is closed the button state is saved in the document property*/
                string startPosition = selection.Range.Start.ToString();
                string endPosition = selection.Range.End.ToString();
                //create a custom property to save infornmation needed to recreate the button
                Microsoft.Office.Core.DocumentProperties properties = (DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties;

                //save the start range pos
                saveInfoInProp("startPosition", startPosition, properties);
                //save the End range pos
                saveInfoInProp("endPosition", endPosition, properties);
                // Store Button Info
                saveInfoInProp("btnName", name, properties);
                saveInfoInProp("btnText", text, properties);

            }//end of if selection != null && selection.Range != null

        }//end of WhenRibionBtnIsClicked Method

        void saveInfoInProp(string name,string value, Microsoft.Office.Core.DocumentProperties properties)
        {
            //save the start range pos
            if (properties.Cast<DocumentProperty>().Where(c => c.Name == "startPosition").Count() == 0)
            {
                properties.Add(name, false, MsoDocProperties.msoPropertyTypeString, value);
            }
            else
            {
                properties[name].Value = value;
                Globals.ThisAddIn.Application.ActiveDocument.Saved = false; //important somtimes
            }
        }

        string loadInfoInProp(string name, Microsoft.Office.Core.DocumentProperties properties)
        {
            string value;
         
            if (properties.Cast<DocumentProperty>().Where(c => c.Name == "startPosition").Count() > 0)
            {
                value = properties[name].Value;
            }
            else
            {
                value = String.Empty;
            }            
            return value;
        }

        void Generatedbtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("I have Been Clicked :-O");
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
