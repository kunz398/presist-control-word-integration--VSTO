using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace test
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnTest_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.WhenRibionBtnIsClicked();
        }

        private void btnDebugStart_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.DebugStartup();
        }
    }
}
