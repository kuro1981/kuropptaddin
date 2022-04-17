using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace kuropptaddin
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Tools.CustomTaskPane myPane;
        public RehearsalTiming rehaCls;

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
