namespace CSATExcelSorting {
    partial class CSATExcelSorting {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CSATExcelSorting));
            this.btnOpenCSAT = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.dgvLocalSites = new System.Windows.Forms.DataGridView();
            this.btnOpenFolder = new System.Windows.Forms.Button();
            this.btnBuildSelectedSheets = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.SiteColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cbSelectAllLocalSites = new System.Windows.Forms.CheckBox();
            this.dgvRemoteSites = new System.Windows.Forms.DataGridView();
            this.cbSelectAllRemoteSites = new System.Windows.Forms.CheckBox();
            this.btnBuildAllSites = new System.Windows.Forms.Button();
            this.btnBuildAllLocalSites = new System.Windows.Forms.Button();
            this.btnBuildAllRemoteSites = new System.Windows.Forms.Button();
            this.bwLoadingCSATs = new System.ComponentModel.BackgroundWorker();
            this.bwBuildingWorkbook = new System.ComponentModel.BackgroundWorker();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.frmLoadingProgress = new System.Windows.Forms.Form();
            this.tbLoadingProgress = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvLocalSites)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRemoteSites)).BeginInit();
            this.frmLoadingProgress.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOpenCSAT
            // 
            this.btnOpenCSAT.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenCSAT.Location = new System.Drawing.Point(217, 13);
            this.btnOpenCSAT.Name = "btnOpenCSAT";
            this.btnOpenCSAT.Size = new System.Drawing.Size(150, 30);
            this.btnOpenCSAT.TabIndex = 0;
            this.btnOpenCSAT.Text = "Open CSAT File";
            this.btnOpenCSAT.UseVisualStyleBackColor = true;
            this.btnOpenCSAT.Click += new System.EventHandler(this.btnOpenCSAT_Click);
            // 
            // btnClose
            // 
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClose.Location = new System.Drawing.Point(471, 409);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(100, 40);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // dgvLocalSites
            // 
            this.dgvLocalSites.AllowUserToAddRows = false;
            this.dgvLocalSites.AllowUserToDeleteRows = false;
            this.dgvLocalSites.AllowUserToResizeRows = false;
            this.dgvLocalSites.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvLocalSites.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dgvLocalSites.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvLocalSites.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dgvLocalSites.Enabled = false;
            this.dgvLocalSites.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvLocalSites.Location = new System.Drawing.Point(13, 73);
            this.dgvLocalSites.Name = "dgvLocalSites";
            this.dgvLocalSites.ReadOnly = true;
            this.dgvLocalSites.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToDisplayedHeaders;
            this.dgvLocalSites.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvLocalSites.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvLocalSites.ShowEditingIcon = false;
            this.dgvLocalSites.Size = new System.Drawing.Size(275, 205);
            this.dgvLocalSites.TabIndex = 3;
            this.dgvLocalSites.SelectionChanged += new System.EventHandler(this.dgvLocalSites_SelectionChanged);
            // 
            // btnOpenFolder
            // 
            this.btnOpenFolder.Enabled = false;
            this.btnOpenFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenFolder.Location = new System.Drawing.Point(217, 409);
            this.btnOpenFolder.Name = "btnOpenFolder";
            this.btnOpenFolder.Size = new System.Drawing.Size(150, 30);
            this.btnOpenFolder.TabIndex = 4;
            this.btnOpenFolder.Text = "Open Folder";
            this.btnOpenFolder.UseVisualStyleBackColor = true;
            this.btnOpenFolder.Click += new System.EventHandler(this.btnOpenFolder_Click);
            // 
            // btnBuildSelectedSheets
            // 
            this.btnBuildSelectedSheets.Enabled = false;
            this.btnBuildSelectedSheets.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBuildSelectedSheets.Location = new System.Drawing.Point(200, 284);
            this.btnBuildSelectedSheets.Name = "btnBuildSelectedSheets";
            this.btnBuildSelectedSheets.Size = new System.Drawing.Size(182, 30);
            this.btnBuildSelectedSheets.TabIndex = 5;
            this.btnBuildSelectedSheets.Text = "Build Selected Sites";
            this.btnBuildSelectedSheets.UseVisualStyleBackColor = true;
            this.btnBuildSelectedSheets.Click += new System.EventHandler(this.btnBuildSelectedSheets_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.DefaultExt = "xlsx";
            this.openFileDialog.Filter = "Excel Files | *.xlsx";
            // 
            // SiteColumn
            // 
            this.SiteColumn.Name = "SiteColumn";
            // 
            // cbSelectAllLocalSites
            // 
            this.cbSelectAllLocalSites.AutoSize = true;
            this.cbSelectAllLocalSites.Enabled = false;
            this.cbSelectAllLocalSites.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbSelectAllLocalSites.Location = new System.Drawing.Point(100, 47);
            this.cbSelectAllLocalSites.Name = "cbSelectAllLocalSites";
            this.cbSelectAllLocalSites.Size = new System.Drawing.Size(83, 20);
            this.cbSelectAllLocalSites.TabIndex = 7;
            this.cbSelectAllLocalSites.Text = "Select All";
            this.cbSelectAllLocalSites.UseVisualStyleBackColor = true;
            this.cbSelectAllLocalSites.CheckedChanged += new System.EventHandler(this.cbSelectAllLocalSites_CheckedChanged);
            // 
            // dgvRemoteSites
            // 
            this.dgvRemoteSites.AllowUserToAddRows = false;
            this.dgvRemoteSites.AllowUserToDeleteRows = false;
            this.dgvRemoteSites.AllowUserToResizeRows = false;
            this.dgvRemoteSites.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvRemoteSites.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dgvRemoteSites.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvRemoteSites.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dgvRemoteSites.Enabled = false;
            this.dgvRemoteSites.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvRemoteSites.Location = new System.Drawing.Point(296, 73);
            this.dgvRemoteSites.Name = "dgvRemoteSites";
            this.dgvRemoteSites.ReadOnly = true;
            this.dgvRemoteSites.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dgvRemoteSites.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dgvRemoteSites.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvRemoteSites.ShowEditingIcon = false;
            this.dgvRemoteSites.Size = new System.Drawing.Size(275, 205);
            this.dgvRemoteSites.TabIndex = 8;
            this.dgvRemoteSites.SelectionChanged += new System.EventHandler(this.dgvRemoteSites_SelectionChanged);
            // 
            // cbSelectAllRemoteSites
            // 
            this.cbSelectAllRemoteSites.AutoSize = true;
            this.cbSelectAllRemoteSites.Enabled = false;
            this.cbSelectAllRemoteSites.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbSelectAllRemoteSites.Location = new System.Drawing.Point(410, 47);
            this.cbSelectAllRemoteSites.Name = "cbSelectAllRemoteSites";
            this.cbSelectAllRemoteSites.Size = new System.Drawing.Size(83, 20);
            this.cbSelectAllRemoteSites.TabIndex = 9;
            this.cbSelectAllRemoteSites.Text = "Select All";
            this.cbSelectAllRemoteSites.UseVisualStyleBackColor = true;
            this.cbSelectAllRemoteSites.CheckedChanged += new System.EventHandler(this.cbSelectAllRemoteSites_CheckedChanged);
            // 
            // btnBuildAllSites
            // 
            this.btnBuildAllSites.Enabled = false;
            this.btnBuildAllSites.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBuildAllSites.Location = new System.Drawing.Point(217, 334);
            this.btnBuildAllSites.Name = "btnBuildAllSites";
            this.btnBuildAllSites.Size = new System.Drawing.Size(150, 60);
            this.btnBuildAllSites.TabIndex = 10;
            this.btnBuildAllSites.Text = "Build All Sites";
            this.btnBuildAllSites.UseVisualStyleBackColor = true;
            this.btnBuildAllSites.Click += new System.EventHandler(this.btnBuildAllSheets_Click);
            // 
            // btnBuildAllLocalSites
            // 
            this.btnBuildAllLocalSites.Enabled = false;
            this.btnBuildAllLocalSites.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBuildAllLocalSites.Location = new System.Drawing.Point(13, 284);
            this.btnBuildAllLocalSites.Name = "btnBuildAllLocalSites";
            this.btnBuildAllLocalSites.Padding = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.btnBuildAllLocalSites.Size = new System.Drawing.Size(150, 50);
            this.btnBuildAllLocalSites.TabIndex = 11;
            this.btnBuildAllLocalSites.Text = "Build All Local Sites";
            this.btnBuildAllLocalSites.UseVisualStyleBackColor = true;
            this.btnBuildAllLocalSites.Click += new System.EventHandler(this.btnBuildAllLocalSites_Click);
            // 
            // btnBuildAllRemoteSites
            // 
            this.btnBuildAllRemoteSites.Enabled = false;
            this.btnBuildAllRemoteSites.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBuildAllRemoteSites.Location = new System.Drawing.Point(421, 284);
            this.btnBuildAllRemoteSites.Name = "btnBuildAllRemoteSites";
            this.btnBuildAllRemoteSites.Padding = new System.Windows.Forms.Padding(6, 0, 0, 0);
            this.btnBuildAllRemoteSites.Size = new System.Drawing.Size(150, 50);
            this.btnBuildAllRemoteSites.TabIndex = 12;
            this.btnBuildAllRemoteSites.Text = "Build All Remote Sites";
            this.btnBuildAllRemoteSites.UseVisualStyleBackColor = true;
            this.btnBuildAllRemoteSites.Click += new System.EventHandler(this.btnBuildAllRemoteSites_Click);
            // 
            // bwLoadingCSATs
            // 
            this.bwLoadingCSATs.WorkerReportsProgress = true;
            this.bwLoadingCSATs.WorkerSupportsCancellation = true;
            this.bwLoadingCSATs.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bwLoadingCSATs_DoWork);
            this.bwLoadingCSATs.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bwLoadingCSATs_ProgressChanged);
            this.bwLoadingCSATs.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bwLoadingCSATs_Completed);
            // 
            // bwBuildingWorkbook
            // 
            this.bwBuildingWorkbook.WorkerReportsProgress = true;
            this.bwBuildingWorkbook.WorkerSupportsCancellation = true;
            this.bwBuildingWorkbook.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bwBuildingWorkbook_DoWork);
            this.bwBuildingWorkbook.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bwBuildingWorkbook_ProgressChanged);
            this.bwBuildingWorkbook.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bwBuildingWorkbook_Completed);
            // 
            // progressBar
            // 
            this.progressBar.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar.ForeColor = System.Drawing.Color.LimeGreen;
            this.progressBar.Location = new System.Drawing.Point(0, 30);
            this.progressBar.MarqueeAnimationSpeed = 10;
            this.progressBar.Maximum = 2000;
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(400, 50);
            this.progressBar.TabIndex = 13;
            // 
            // frmLoadingProgress
            // 
            this.frmLoadingProgress.ClientSize = new System.Drawing.Size(400, 80);
            this.frmLoadingProgress.ControlBox = false;
            this.frmLoadingProgress.Controls.Add(this.tbLoadingProgress);
            this.frmLoadingProgress.Controls.Add(this.progressBar);
            this.frmLoadingProgress.Location = new System.Drawing.Point(0, 0);
            this.frmLoadingProgress.Name = "frmLoadingProgress";
            this.frmLoadingProgress.ShowInTaskbar = false;
            this.frmLoadingProgress.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.frmLoadingProgress.Visible = false;
            // 
            // tbLoadingProgress
            // 
            this.tbLoadingProgress.Dock = System.Windows.Forms.DockStyle.Top;
            this.tbLoadingProgress.Enabled = false;
            this.tbLoadingProgress.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold);
            this.tbLoadingProgress.Location = new System.Drawing.Point(0, 0);
            this.tbLoadingProgress.Multiline = true;
            this.tbLoadingProgress.Name = "tbLoadingProgress";
            this.tbLoadingProgress.ReadOnly = true;
            this.tbLoadingProgress.Size = new System.Drawing.Size(400, 30);
            this.tbLoadingProgress.TabIndex = 0;
            this.tbLoadingProgress.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // CSATFixer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(584, 462);
            this.Controls.Add(this.btnBuildAllRemoteSites);
            this.Controls.Add(this.btnBuildAllLocalSites);
            this.Controls.Add(this.btnBuildAllSites);
            this.Controls.Add(this.cbSelectAllRemoteSites);
            this.Controls.Add(this.dgvRemoteSites);
            this.Controls.Add(this.cbSelectAllLocalSites);
            this.Controls.Add(this.btnBuildSelectedSheets);
            this.Controls.Add(this.btnOpenFolder);
            this.Controls.Add(this.dgvLocalSites);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnOpenCSAT);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CSATFixer";
            this.Padding = new System.Windows.Forms.Padding(10);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Filter CSATs";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.CSATFixer_Closing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.CSATFixer_Closing);
            this.Load += new System.EventHandler(this.CSATFixer_Load);
            this.Disposed += new System.EventHandler(this.CSATFixer_Closing);
            ((System.ComponentModel.ISupportInitialize)(this.dgvLocalSites)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRemoteSites)).EndInit();
            this.frmLoadingProgress.ResumeLayout(false);
            this.frmLoadingProgress.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpenCSAT;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnOpenFolder;
        private System.Windows.Forms.Button btnBuildSelectedSheets;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.DataGridViewTextBoxColumn SiteColumn;
        private System.Windows.Forms.DataGridView dgvLocalSites;
        private System.Windows.Forms.CheckBox cbSelectAllLocalSites;
        private System.Windows.Forms.DataGridView dgvRemoteSites;
        private System.Windows.Forms.CheckBox cbSelectAllRemoteSites;
        private System.Windows.Forms.Button btnBuildAllSites;
        private System.Windows.Forms.Button btnBuildAllLocalSites;
        private System.Windows.Forms.Button btnBuildAllRemoteSites;
        private System.ComponentModel.BackgroundWorker bwLoadingCSATs;
        private System.ComponentModel.BackgroundWorker bwBuildingWorkbook;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Form frmLoadingProgress;
        private System.Windows.Forms.TextBox tbLoadingProgress;
    }
}

