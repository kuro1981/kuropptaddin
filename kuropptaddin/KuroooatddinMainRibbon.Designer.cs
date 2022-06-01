namespace kuropptaddin
{
    partial class KuroooatddinMainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public KuroooatddinMainRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.OpenEditorBtn = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btn_merge_note = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.GetInfoBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Kurodapp";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.OpenEditorBtn);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // OpenEditorBtn
            // 
            this.OpenEditorBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.OpenEditorBtn.Label = "OpenEditor";
            this.OpenEditorBtn.Name = "OpenEditorBtn";
            this.OpenEditorBtn.OfficeImageId = "FormControlEditBox";
            this.OpenEditorBtn.ShowImage = true;
            this.OpenEditorBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenEditorBtn_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button1);
            this.group2.Items.Add(this.btn_merge_note);
            this.group2.Items.Add(this.separator2);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // button1
            // 
            this.button1.Label = "ナレーション削除版を作成";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // btn_merge_note
            // 
            this.btn_merge_note.Label = "Note を一つにまとめる";
            this.btn_merge_note.Name = "btn_merge_note";
            this.btn_merge_note.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_merge_note_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // group3
            // 
            this.group3.Items.Add(this.GetInfoBtn);
            this.group3.Label = "group3";
            this.group3.Name = "group3";
            // 
            // GetInfoBtn
            // 
            this.GetInfoBtn.Label = "プレゼンテーションの情報";
            this.GetInfoBtn.Name = "GetInfoBtn";
            this.GetInfoBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetInfoBtn_Click);
            // 
            // KuroooatddinMainRibbon
            // 
            this.Name = "KuroooatddinMainRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenEditorBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_merge_note;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetInfoBtn;
    }

    partial class ThisRibbonCollection
    {
        internal KuroooatddinMainRibbon Ribbon1
        {
            get { return this.GetRibbon<KuroooatddinMainRibbon>(); }
        }
    }
}
