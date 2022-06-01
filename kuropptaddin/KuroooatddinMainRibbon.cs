using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace kuropptaddin
{
    public partial class KuroooatddinMainRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void OpenEditorBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ShowPanel();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CreateNonVoicePPTX();
        }


        private void btn_merge_note_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.MergeNote();

        }

        private void GetInfoBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.getPresentationInfo();
        }
    }
}
