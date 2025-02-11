using CharExtractorRibbon;

namespace ExcelCharExtractor
{
    partial class CharExtractorRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        /// 
        private System.ComponentModel.IContainer components = null;

        public CharExtractorRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CharExtractorRibbon));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.提取 = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.删除字符 = this.Factory.CreateRibbonButton();
            this.提取字符 = this.Factory.CreateRibbonButton();
            this.选择类型 = this.Factory.CreateRibbonDropDown();
            this.常用工具 = this.Factory.CreateRibbonGroup();
            this.按颜色汇总 = this.Factory.CreateRibbonButton();
            this.筛选选定值 = this.Factory.CreateRibbonButton();
            this.单元格内每行重新编号 = this.Factory.CreateRibbonButton();
            this.btnDeleteSymbol = this.Factory.CreateRibbonButton();
            this.询价文件 = this.Factory.CreateRibbonGroup();
            this.txtApiUrl = this.Factory.CreateRibbonEditBox();
            this.txtApiKey = this.Factory.CreateRibbonEditBox();
            this.txtModelName = this.Factory.CreateRibbonEditBox();
            this.btnSaveSettings = this.Factory.CreateRibbonButton();
            this.AI功能区 = this.Factory.CreateRibbonGroup();
            this.cmbTaskType = this.Factory.CreateRibbonComboBox();
            this.btnSendRequest = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.提取.SuspendLayout();
            this.box1.SuspendLayout();
            this.常用工具.SuspendLayout();
            this.询价文件.SuspendLayout();
            this.AI功能区.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.提取);
            this.tab1.Groups.Add(this.常用工具);
            this.tab1.Groups.Add(this.询价文件);
            this.tab1.Groups.Add(this.AI功能区);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // 提取
            // 
            this.提取.Items.Add(this.box1);
            this.提取.Items.Add(this.选择类型);
            this.提取.Label = "提取";
            this.提取.Name = "提取";
            // 
            // box1
            // 
            this.box1.Items.Add(this.删除字符);
            this.box1.Items.Add(this.提取字符);
            this.box1.Name = "box1";
            // 
            // 删除字符
            // 
            this.删除字符.Label = "删除字符";
            this.删除字符.Name = "删除字符";
            this.删除字符.OfficeImageId = "Delete";
            this.删除字符.ShowImage = true;
            this.删除字符.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // 提取字符
            // 
            this.提取字符.Label = "提取字符";
            this.提取字符.Name = "提取字符";
            this.提取字符.OfficeImageId = "Filter";
            this.提取字符.ShowImage = true;
            this.提取字符.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // 选择类型
            // 
            this.选择类型.Image = ((System.Drawing.Image)(resources.GetObject("选择类型.Image")));
            ribbonDropDownItemImpl1.Label = "汉字";
            ribbonDropDownItemImpl2.Label = "英文";
            ribbonDropDownItemImpl3.Label = "数字";
            ribbonDropDownItemImpl4.Label = "英文标点";
            ribbonDropDownItemImpl5.Label = "中文标点";
            this.选择类型.Items.Add(ribbonDropDownItemImpl1);
            this.选择类型.Items.Add(ribbonDropDownItemImpl2);
            this.选择类型.Items.Add(ribbonDropDownItemImpl3);
            this.选择类型.Items.Add(ribbonDropDownItemImpl4);
            this.选择类型.Items.Add(ribbonDropDownItemImpl5);
            this.选择类型.Label = "选择类型";
            this.选择类型.Name = "选择类型";
            this.选择类型.OfficeImageId = "ilterBySelection";
            this.选择类型.ShowImage = true;
            this.选择类型.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown1_SelectionChanged);
            // 
            // 常用工具
            // 
            this.常用工具.Items.Add(this.按颜色汇总);
            this.常用工具.Items.Add(this.筛选选定值);
            this.常用工具.Items.Add(this.单元格内每行重新编号);
            this.常用工具.Items.Add(this.btnDeleteSymbol);
            this.常用工具.Label = "常用功能";
            this.常用工具.Name = "常用工具";
            // 
            // 按颜色汇总
            // 
            this.按颜色汇总.Image = ((System.Drawing.Image)(resources.GetObject("按颜色汇总.Image")));
            this.按颜色汇总.Label = "按颜色汇总";
            this.按颜色汇总.Name = "按颜色汇总";
            this.按颜色汇总.OfficeImageId = "ColorsGallery";
            this.按颜色汇总.ShowImage = true;
            this.按颜色汇总.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // 筛选选定值
            // 
            this.筛选选定值.Image = ((System.Drawing.Image)(resources.GetObject("筛选选定值.Image")));
            this.筛选选定值.Label = "筛选选定值";
            this.筛选选定值.Name = "筛选选定值";
            this.筛选选定值.OfficeImageId = "FilterBySelection";
            this.筛选选定值.ShowImage = true;
            this.筛选选定值.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.筛选选定值_Click);
            // 
            // 单元格内每行重新编号
            // 
            this.单元格内每行重新编号.Image = ((System.Drawing.Image)(resources.GetObject("单元格内每行重新编号.Image")));
            this.单元格内每行重新编号.Label = "单元格内每行重新编号";
            this.单元格内每行重新编号.Name = "单元格内每行重新编号";
            this.单元格内每行重新编号.ShowImage = true;
            this.单元格内每行重新编号.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.单元格内每行重新编号_Click);
            // 
            // btnDeleteSymbol
            // 
            this.btnDeleteSymbol.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteSymbol.Image")));
            this.btnDeleteSymbol.Label = "删除符号/文字";
            this.btnDeleteSymbol.Name = "btnDeleteSymbol";
            this.btnDeleteSymbol.ShowImage = true;
            this.btnDeleteSymbol.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteSymbol_Click);
            // 
            // 询价文件
            // 
            this.询价文件.Items.Add(this.txtApiUrl);
            this.询价文件.Items.Add(this.txtApiKey);
            this.询价文件.Items.Add(this.txtModelName);
            this.询价文件.Items.Add(this.btnSaveSettings);
            this.询价文件.Label = "AI配置区";
            this.询价文件.Name = "询价文件";
            // 
            // txtApiUrl
            // 
            this.txtApiUrl.Label = "API URL";
            this.txtApiUrl.Name = "txtApiUrl";
            this.txtApiUrl.Tag = "";
            this.txtApiUrl.Text = null;
            // 
            // txtApiKey
            // 
            this.txtApiKey.Label = "API Key";
            this.txtApiKey.Name = "txtApiKey";
            this.txtApiKey.Text = null;
            // 
            // txtModelName
            // 
            this.txtModelName.Label = "模型名称";
            this.txtModelName.Name = "txtModelName";
            this.txtModelName.Text = null;
            // 
            // btnSaveSettings
            // 
            this.btnSaveSettings.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveSettings.Image")));
            this.btnSaveSettings.Label = "保存设置";
            this.btnSaveSettings.Name = "btnSaveSettings";
            this.btnSaveSettings.ShowImage = true;
            this.btnSaveSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveSettings_Click);
            // 
            // AI功能区
            // 
            this.AI功能区.Items.Add(this.cmbTaskType);
            this.AI功能区.Items.Add(this.btnSendRequest);
            this.AI功能区.Label = "AI功能区";
            this.AI功能区.Name = "AI功能区";
            // 
            // cmbTaskType
            // 
            this.cmbTaskType.Label = "任务类型";
            this.cmbTaskType.Name = "cmbTaskType";
            this.cmbTaskType.Text = null;
            // 
            // btnSendRequest
            // 
            this.btnSendRequest.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSendRequest.Image = ((System.Drawing.Image)(resources.GetObject("btnSendRequest.Image")));
            this.btnSendRequest.Label = "问AI";
            this.btnSendRequest.Name = "btnSendRequest";
            this.btnSendRequest.ShowImage = true;
            this.btnSendRequest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSendRequest_Click);
            // 
            // CharExtractorRibbon
            // 
            this.Name = "CharExtractorRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CharExtractorRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.提取.ResumeLayout(false);
            this.提取.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.常用工具.ResumeLayout(false);
            this.常用工具.PerformLayout();
            this.询价文件.ResumeLayout(false);
            this.询价文件.PerformLayout();
            this.AI功能区.ResumeLayout(false);
            this.AI功能区.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 提取;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 提取字符;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 删除字符;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown 选择类型;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 常用工具;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 按颜色汇总;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 筛选选定值;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 单元格内每行重新编号;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 询价文件;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtApiUrl;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtApiKey;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtModelName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSendRequest;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbTaskType;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AI功能区;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteSymbol;
    }

    
}
