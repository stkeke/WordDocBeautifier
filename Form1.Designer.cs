namespace Word_Doc_Beautifier
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnMultiChoice_FindAndConvertNumberIndex = new Button();
            groupBox1 = new GroupBox();
            btnFindAndFormat3HChoices = new Button();
            btnFindAndFormat4HChoices = new Button();
            btnPageSetNarrorwMarginFooter = new Button();
            btnConvertToDocx = new Button();
            lblWordDoc = new Label();
            btnSaveAsPDF = new Button();
            lblCompatibility = new Label();
            tbHeader = new TextBox();
            btnSetHeader = new Button();
            lbHeaderHistory = new ListBox();
            lblLocalPath = new Label();
            label1 = new Label();
            label2 = new Label();
            tbFileName = new TextBox();
            btnRename = new Button();
            btnSaveAs = new Button();
            btnRefresh = new Button();
            lbDocuments = new ListBox();
            btnCloseAllWordApps = new Button();
            lblStatus = new Label();
            tbLog = new TextBox();
            btn = new Button();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // btnMultiChoice_FindAndConvertNumberIndex
            // 
            btnMultiChoice_FindAndConvertNumberIndex.AutoSize = true;
            btnMultiChoice_FindAndConvertNumberIndex.Location = new Point(30, 64);
            btnMultiChoice_FindAndConvertNumberIndex.Name = "btnMultiChoice_FindAndConvertNumberIndex";
            btnMultiChoice_FindAndConvertNumberIndex.Size = new Size(296, 35);
            btnMultiChoice_FindAndConvertNumberIndex.TabIndex = 0;
            btnMultiChoice_FindAndConvertNumberIndex.Text = "Find and Convert Numbered Index";
            btnMultiChoice_FindAndConvertNumberIndex.UseVisualStyleBackColor = true;
            btnMultiChoice_FindAndConvertNumberIndex.Click += btnMultiChoice_FindAndConvertNumberIndex_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(btn);
            groupBox1.Controls.Add(btnFindAndFormat3HChoices);
            groupBox1.Controls.Add(btnFindAndFormat4HChoices);
            groupBox1.Controls.Add(btnMultiChoice_FindAndConvertNumberIndex);
            groupBox1.Location = new Point(49, 551);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(495, 465);
            groupBox1.TabIndex = 1;
            groupBox1.TabStop = false;
            groupBox1.Text = "Multiple Choices";
            // 
            // btnFindAndFormat3HChoices
            // 
            btnFindAndFormat3HChoices.AutoSize = true;
            btnFindAndFormat3HChoices.Location = new Point(30, 211);
            btnFindAndFormat3HChoices.Name = "btnFindAndFormat3HChoices";
            btnFindAndFormat3HChoices.Size = new Size(296, 35);
            btnFindAndFormat3HChoices.TabIndex = 0;
            btnFindAndFormat3HChoices.Text = "Find and Format 3H Choices";
            btnFindAndFormat3HChoices.UseVisualStyleBackColor = true;
            btnFindAndFormat3HChoices.Click += btnFindAndFormat3HChoices_Click;
            // 
            // btnFindAndFormat4HChoices
            // 
            btnFindAndFormat4HChoices.AutoSize = true;
            btnFindAndFormat4HChoices.Location = new Point(30, 157);
            btnFindAndFormat4HChoices.Name = "btnFindAndFormat4HChoices";
            btnFindAndFormat4HChoices.Size = new Size(296, 35);
            btnFindAndFormat4HChoices.TabIndex = 0;
            btnFindAndFormat4HChoices.Text = "Find and Format 4H Choices";
            btnFindAndFormat4HChoices.UseVisualStyleBackColor = true;
            btnFindAndFormat4HChoices.Click += btnFindAndFormat4HChoices_Click;
            // 
            // btnPageSetNarrorwMarginFooter
            // 
            btnPageSetNarrorwMarginFooter.Location = new Point(61, 370);
            btnPageSetNarrorwMarginFooter.Name = "btnPageSetNarrorwMarginFooter";
            btnPageSetNarrorwMarginFooter.Size = new Size(341, 45);
            btnPageSetNarrorwMarginFooter.TabIndex = 2;
            btnPageSetNarrorwMarginFooter.Text = "Page Set Narrow Margin and Footer";
            btnPageSetNarrorwMarginFooter.UseVisualStyleBackColor = true;
            btnPageSetNarrorwMarginFooter.Click += btnPageSetNarrorwMarginFooter_Click;
            // 
            // btnConvertToDocx
            // 
            btnConvertToDocx.Location = new Point(61, 433);
            btnConvertToDocx.Name = "btnConvertToDocx";
            btnConvertToDocx.Size = new Size(341, 42);
            btnConvertToDocx.TabIndex = 4;
            btnConvertToDocx.Text = "Convert To Current Word .docx";
            btnConvertToDocx.UseVisualStyleBackColor = true;
            btnConvertToDocx.Click += btnConvertToDocx_Click;
            // 
            // lblWordDoc
            // 
            lblWordDoc.ImageAlign = ContentAlignment.BottomLeft;
            lblWordDoc.Location = new Point(58, 9);
            lblWordDoc.Name = "lblWordDoc";
            lblWordDoc.Size = new Size(969, 59);
            lblWordDoc.TabIndex = 3;
            lblWordDoc.Text = "Active Document";
            // 
            // btnSaveAsPDF
            // 
            btnSaveAsPDF.AutoSize = true;
            btnSaveAsPDF.Location = new Point(58, 303);
            btnSaveAsPDF.Name = "btnSaveAsPDF";
            btnSaveAsPDF.Size = new Size(341, 45);
            btnSaveAsPDF.TabIndex = 5;
            btnSaveAsPDF.Text = "Save as PDF";
            btnSaveAsPDF.UseVisualStyleBackColor = true;
            btnSaveAsPDF.Click += btnSaveAsPDF_Click;
            // 
            // lblCompatibility
            // 
            lblCompatibility.AutoSize = true;
            lblCompatibility.Location = new Point(1064, 74);
            lblCompatibility.Name = "lblCompatibility";
            lblCompatibility.Size = new Size(167, 25);
            lblCompatibility.TabIndex = 6;
            lblCompatibility.Text = "Word Compatibility";
            lblCompatibility.Click += label1_Click;
            // 
            // tbHeader
            // 
            tbHeader.Location = new Point(58, 184);
            tbHeader.Name = "tbHeader";
            tbHeader.Size = new Size(486, 31);
            tbHeader.TabIndex = 7;
            tbHeader.Text = "Enter Doc Header Here";
            // 
            // btnSetHeader
            // 
            btnSetHeader.Location = new Point(58, 234);
            btnSetHeader.Name = "btnSetHeader";
            btnSetHeader.Size = new Size(236, 50);
            btnSetHeader.TabIndex = 8;
            btnSetHeader.Text = "Set Header";
            btnSetHeader.UseVisualStyleBackColor = true;
            btnSetHeader.Click += btnSetHeader_Click;
            // 
            // lbHeaderHistory
            // 
            lbHeaderHistory.FormattingEnabled = true;
            lbHeaderHistory.ItemHeight = 25;
            lbHeaderHistory.Location = new Point(577, 184);
            lbHeaderHistory.Name = "lbHeaderHistory";
            lbHeaderHistory.Size = new Size(557, 279);
            lbHeaderHistory.TabIndex = 9;
            lbHeaderHistory.SelectedIndexChanged += lbHeaderHistory_SelectedIndexChanged;
            lbHeaderHistory.DoubleClick += lbHeaderHistory_DoubleClick;
            // 
            // lblLocalPath
            // 
            lblLocalPath.Location = new Point(172, 68);
            lblLocalPath.Name = "lblLocalPath";
            lblLocalPath.Size = new Size(882, 51);
            lblLocalPath.TabIndex = 10;
            lblLocalPath.TabStop = true;
            lblLocalPath.Text = "Local Path";
            lblLocalPath.Click += lblLocalPath_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(61, 74);
            label1.Name = "label1";
            label1.Size = new Size(95, 25);
            label1.TabIndex = 11;
            label1.Text = "Local Path:";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(61, 125);
            label2.Name = "label2";
            label2.Size = new Size(94, 25);
            label2.TabIndex = 12;
            label2.Text = "File Name:";
            // 
            // tbFileName
            // 
            tbFileName.Location = new Point(172, 122);
            tbFileName.Name = "tbFileName";
            tbFileName.Size = new Size(613, 31);
            tbFileName.TabIndex = 13;
            tbFileName.Text = "Local File Name";
            tbFileName.TextChanged += tbFileName_TextChanged;
            // 
            // btnRename
            // 
            btnRename.Location = new Point(812, 120);
            btnRename.Name = "btnRename";
            btnRename.Size = new Size(112, 34);
            btnRename.TabIndex = 14;
            btnRename.Text = "Rename";
            btnRename.UseVisualStyleBackColor = true;
            btnRename.Click += btnRename_Click;
            // 
            // btnSaveAs
            // 
            btnSaveAs.Location = new Point(942, 119);
            btnSaveAs.Name = "btnSaveAs";
            btnSaveAs.Size = new Size(112, 34);
            btnSaveAs.TabIndex = 15;
            btnSaveAs.Text = "Save As";
            btnSaveAs.UseVisualStyleBackColor = true;
            btnSaveAs.Click += btnSaveAs_Click;
            // 
            // btnRefresh
            // 
            btnRefresh.Location = new Point(1064, 12);
            btnRefresh.Name = "btnRefresh";
            btnRefresh.Size = new Size(112, 34);
            btnRefresh.TabIndex = 16;
            btnRefresh.Text = "Refresh";
            btnRefresh.UseVisualStyleBackColor = true;
            btnRefresh.Click += btnRefresh_Click;
            // 
            // lbDocuments
            // 
            lbDocuments.FormattingEnabled = true;
            lbDocuments.ItemHeight = 25;
            lbDocuments.Location = new Point(577, 528);
            lbDocuments.Name = "lbDocuments";
            lbDocuments.Size = new Size(644, 279);
            lbDocuments.TabIndex = 17;
            lbDocuments.Click += lbDocuments_Click;
            lbDocuments.SelectedIndexChanged += lbDocuments_SelectedIndexChanged;
            lbDocuments.DoubleClick += lbDocuments_DoubleClick;
            // 
            // btnCloseAllWordApps
            // 
            btnCloseAllWordApps.Location = new Point(1203, 12);
            btnCloseAllWordApps.Name = "btnCloseAllWordApps";
            btnCloseAllWordApps.Size = new Size(211, 34);
            btnCloseAllWordApps.TabIndex = 18;
            btnCloseAllWordApps.Text = "Close All Word Apps";
            btnCloseAllWordApps.UseVisualStyleBackColor = true;
            btnCloseAllWordApps.Click += btnCloseAllWordApps_Click;
            // 
            // lblStatus
            // 
            lblStatus.Dock = DockStyle.Bottom;
            lblStatus.Location = new Point(0, 1043);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(1806, 38);
            lblStatus.TabIndex = 19;
            lblStatus.Text = "Status";
            // 
            // tbLog
            // 
            tbLog.Location = new Point(1271, 589);
            tbLog.Multiline = true;
            tbLog.Name = "tbLog";
            tbLog.ScrollBars = ScrollBars.Vertical;
            tbLog.Size = new Size(535, 454);
            tbLog.TabIndex = 20;
            // 
            // btn
            // 
            btn.AutoSize = true;
            btn.Location = new Point(30, 292);
            btn.Name = "btn";
            btn.Size = new Size(296, 35);
            btn.TabIndex = 0;
            btn.Text = "Find and Format V Choices";
            btn.UseVisualStyleBackColor = true;
            btn.Click += btnFormatVChoices_Click;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1806, 1081);
            Controls.Add(tbLog);
            Controls.Add(lblStatus);
            Controls.Add(btnCloseAllWordApps);
            Controls.Add(lbDocuments);
            Controls.Add(btnRefresh);
            Controls.Add(btnSaveAs);
            Controls.Add(btnRename);
            Controls.Add(tbFileName);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(lblLocalPath);
            Controls.Add(lbHeaderHistory);
            Controls.Add(btnSetHeader);
            Controls.Add(tbHeader);
            Controls.Add(lblCompatibility);
            Controls.Add(btnSaveAsPDF);
            Controls.Add(btnConvertToDocx);
            Controls.Add(lblWordDoc);
            Controls.Add(btnPageSetNarrorwMarginFooter);
            Controls.Add(groupBox1);
            Name = "MainForm";
            Text = "Word Doc Manager";
            Load += Form1_Load;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnMultiChoice_FindAndConvertNumberIndex;
        private GroupBox groupBox1;
        private Button btnPageSetNarrorwMarginFooter;
        private Button btnFindAndFormat4HChoices;
        private Button btnConvertToDocx;
        private Label lblWordDoc;
        private Button btnSaveAsPDF;
        private Label lblCompatibility;
        private TextBox tbHeader;
        private Button btnSetHeader;
        private ListBox lbHeaderHistory;
        private Label lblLocalPath;
        private Label label1;
        private Label label2;
        private TextBox tbFileName;
        private Button btnRename;
        private Button btnSaveAs;
        private Button btnRefresh;
        private ListBox lbDocuments;
        private Button btnCloseAllWordApps;
        private Label lblStatus;
        private TextBox tbLog;
        private Button btnFindAndFormat3HChoices;
        private Button btn;
    }
}
