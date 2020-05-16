namespace OutlookGoogleCalendarSync.Forms {
    partial class ColourMap {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ColourMap));
            this.colourGridView = new System.Windows.Forms.DataGridView();
            this.OutlookColour = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.GoogleColour = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.btSave = new System.Windows.Forms.Button();
            this.btCancel = new System.Windows.Forms.Button();
            this.txtInfo = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.colourGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // colourGridView
            // 
            this.colourGridView.AllowUserToAddRows = false;
            this.colourGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.colourGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.colourGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.OutlookColour,
            this.GoogleColour});
            this.colourGridView.Location = new System.Drawing.Point(12, 95);
            this.colourGridView.Name = "colourGridView";
            this.colourGridView.Size = new System.Drawing.Size(468, 106);
            this.colourGridView.TabIndex = 0;
            this.colourGridView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.colourGridView_CellClick);
            this.colourGridView.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.colourGridView_DataError);
            this.colourGridView.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.colourGridView_EditingControlShowing);
            // 
            // OutlookColour
            // 
            this.OutlookColour.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.OutlookColour.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.OutlookColour.DividerWidth = 2;
            this.OutlookColour.FillWeight = 50F;
            this.OutlookColour.HeaderText = "Outlook Category";
            this.OutlookColour.Name = "OutlookColour";
            this.OutlookColour.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.OutlookColour.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // GoogleColour
            // 
            this.GoogleColour.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.GoogleColour.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.GoogleColour.FillWeight = 50F;
            this.GoogleColour.HeaderText = "Google Colour";
            this.GoogleColour.Name = "GoogleColour";
            // 
            // btSave
            // 
            this.btSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btSave.Location = new System.Drawing.Point(405, 211);
            this.btSave.Name = "btSave";
            this.btSave.Size = new System.Drawing.Size(75, 23);
            this.btSave.TabIndex = 1;
            this.btSave.Text = "Save";
            this.btSave.UseVisualStyleBackColor = false;
            this.btSave.Click += new System.EventHandler(this.btSave_Click);
            // 
            // btCancel
            // 
            this.btCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btCancel.Location = new System.Drawing.Point(324, 211);
            this.btCancel.Name = "btCancel";
            this.btCancel.Size = new System.Drawing.Size(75, 23);
            this.btCancel.TabIndex = 2;
            this.btCancel.Text = "Cancel";
            this.btCancel.UseVisualStyleBackColor = false;
            // 
            // txtInfo
            // 
            this.txtInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtInfo.BackColor = System.Drawing.SystemColors.Control;
            this.txtInfo.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtInfo.Location = new System.Drawing.Point(12, 12);
            this.txtInfo.Multiline = true;
            this.txtInfo.Name = "txtInfo";
            this.txtInfo.Size = new System.Drawing.Size(468, 77);
            this.txtInfo.TabIndex = 9;
            this.txtInfo.Text = resources.GetString("txtInfo.Text");
            // 
            // ColourMap
            // 
            this.AcceptButton = this.btSave;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btCancel;
            this.ClientSize = new System.Drawing.Size(492, 242);
            this.Controls.Add(this.txtInfo);
            this.Controls.Add(this.btCancel);
            this.Controls.Add(this.btSave);
            this.Controls.Add(this.colourGridView);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ColourMap";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Custom Colour Mapping";
            ((System.ComponentModel.ISupportInitialize)(this.colourGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView colourGridView;
        private System.Windows.Forms.Button btSave;
        private System.Windows.Forms.Button btCancel;
        private System.Windows.Forms.TextBox txtInfo;
        private System.Windows.Forms.DataGridViewComboBoxColumn OutlookColour;
        private System.Windows.Forms.DataGridViewComboBoxColumn GoogleColour;
    }
}