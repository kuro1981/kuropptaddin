using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace kuropptaddin
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Tools.CustomTaskPane myPane;
        public RehearsalTiming rehaCls;
        public MergeNoteForm mergeNoteForm;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            myPane = this.CustomTaskPanes.Add(new TagEditor(), "eLearning Reharsal tools");

        }
        public void ShowPanel()
        {
            myPane.Visible = true;
            myPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionBottom;
            myPane.Height = 450;
            rehaCls = new RehearsalTiming(Application.ActiveWindow.View.Slide);
        }

        public void CreateNonVoicePPTX()
        {
            var result = MessageBox.Show(
                "ナレーションなしPPTXを作成しますか？",
                "PPTXの作成",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.None);
            if (result == DialogResult.Yes)
            {
                foreach (PowerPoint.Slide slide in Application.ActivePresentation.Slides)
                {
                    removeNaration(slide);
                }
            }
        }

        public void removeNaration(PowerPoint.Slide slide)
        {
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.Type == Office.MsoShapeType.msoMedia && shape.MediaType == PowerPoint.PpMediaType.ppMediaTypeSound)
                {
                    shape.Delete();
                }
            }

        }

        public void CreateNonVoicePDF()
        {

        }

        public void getPresentationInfo()
        {
            long noteStringCount = 0;
            long objectStringCount = 0;
            foreach (PowerPoint.Slide slide in Application.ActivePresentation.Slides)
            {
                foreach(PowerPoint.Shape shape in slide.Shapes)
                {
                    if(shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        objectStringCount += shape.TextFrame.TextRange.Text.Count();
                    }
                }
                if (slide.HasNotesPage == Office.MsoTriState.msoTrue)
                {
                    noteStringCount += slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text.Count();
                }
            }
            var messageLines = new string[]
            {
                "スライド内の文字数 = " + objectStringCount.ToString(),
                "ノートの文字数 = " + noteStringCount.ToString(),
            };
            var result = MessageBox.Show(
                string.Join(Environment.NewLine, messageLines),
                "プレゼンテーションの情報",
                MessageBoxButtons.OK,
                MessageBoxIcon.None);

        }
            

        public void MergeNote()
        {
            mergeNoteForm = new MergeNoteForm();
            mergeNoteForm.Show();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        //protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() 
        //{
        //    return new Ribbon2();
        //}


        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }


}
