**This code demonstrates how to save WinForms controls in Word**
*the official [documentation](https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-adding-controls-to-a-document-at-run-time-in-a-vsto-add-in?view=vs-2019) for Microsoft does not give the code for how to save winform controls when the word document is saved as said by Microsoft "Windows Forms controls are not persisted when the document is saved and then closed"*

firstly create a button on the ribbon to generate a button on runtime.
and dont forget to add refernce to `Microsoft.Office.Tools.Word.v4.0.Utilities.dl`
![enter image description here](https://i.stack.imgur.com/BwwZl.png)
then on the button click event we added this piece of code to reference a method in `ThisAddsIn.cs` Class

    private void btnTest_Click(object sender, RibbonControlEventArgs e)
    {
      Globals.ThisAddIn.WhenRibionBtnIsClicked();
    }
now we go to the `ThisAddsIn.cs` Class and Make a method called `WhenRibionBtnIsClicked()`

        internal  void  WhenRibionBtnIsClicked()
        {
        Document vstoDocument = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
         Word.Selection selection = this.Application.Selection;
	      if (selection != null && selection.Range != null)
		    {
		    //Generate button on run time
		    	string name = "myBtn";
		    	Button button = new Button();
		    	button.Click += new EventHandler(Generatedbtn_Click);
		    	button = vstoDocument.Controls.AddButton(selection.Range, 100, 30, name);
		    	button.Click += Generatedbtn_Click; //for the click function
		    	button.Text = "I am A Generated Button";
		    	button.Name = name;
		        }

next we need to save the button state to a location you can do this multiple ways the best way to do it is store it in the documents custom properties 
what i did was create a function to achieve this 

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
the way we used this was

    string startPosition = selection.Range.Start.ToString();//get start position
    string endPosition = selection.Range.End.ToString();//get endposition
        //save the start range pos
        saveInfoInProp("startPosition", startPosition, properties);
    this function saves the start position in to the properties of the document so we can later make use of them when the document starts up

so now we are done with creating the button, next we need to recreate this button once the Word document starts with so this needs to be coded in the `ThisAddIn_Startup` function 

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

there is of course a function to load what ever was stored in the property 
```
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
```
and now to code a click event handler for this is easy since you attached a event handle

    void Generatedbtn_Click(object sender, EventArgs e)
    {
     MessageBox.Show("I have Been Clicked :-O");
    }

