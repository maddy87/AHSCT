namespace maddytry1
{
    partial class Favourites
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dgFavourites = new System.Windows.Forms.DataGridView();
            this.grpFavourites = new System.Windows.Forms.GroupBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnCreateNew = new System.Windows.Forms.Button();
            this.btnCreateSimilar = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgFavourites)).BeginInit();
            this.grpFavourites.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgFavourites
            // 
            this.dgFavourites.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgFavourites.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgFavourites.BackgroundColor = System.Drawing.Color.White;
            this.dgFavourites.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dgFavourites.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dgFavourites.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgFavourites.Location = new System.Drawing.Point(7, 19);
            this.dgFavourites.Name = "dgFavourites";
            this.dgFavourites.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.dgFavourites.RowHeadersVisible = false;
            this.dgFavourites.Size = new System.Drawing.Size(696, 150);
            this.dgFavourites.TabIndex = 0;
            this.dgFavourites.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgFavourites_CellClick);
            // 
            // grpFavourites
            // 
            this.grpFavourites.AutoSize = true;
            this.grpFavourites.BackColor = System.Drawing.Color.White;
            this.grpFavourites.Controls.Add(this.dgFavourites);
            this.grpFavourites.ForeColor = System.Drawing.Color.Navy;
            this.grpFavourites.Location = new System.Drawing.Point(14, 12);
            this.grpFavourites.Name = "grpFavourites";
            this.grpFavourites.Size = new System.Drawing.Size(718, 189);
            this.grpFavourites.TabIndex = 1;
            this.grpFavourites.TabStop = false;
            this.grpFavourites.Text = "Select Favourites";
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.Color.White;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.ForeColor = System.Drawing.SystemColors.Highlight;
            this.btnCancel.Location = new System.Drawing.Point(482, 223);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(207, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "CANCEL";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnCreateNew
            // 
            this.btnCreateNew.BackColor = System.Drawing.Color.White;
            this.btnCreateNew.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCreateNew.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateNew.ForeColor = System.Drawing.Color.Navy;
            this.btnCreateNew.Location = new System.Drawing.Point(253, 223);
            this.btnCreateNew.Name = "btnCreateNew";
            this.btnCreateNew.Size = new System.Drawing.Size(207, 23);
            this.btnCreateNew.TabIndex = 5;
            this.btnCreateNew.Text = "CREATE NEW NOTIFICATION";
            this.btnCreateNew.UseVisualStyleBackColor = false;
            this.btnCreateNew.Click += new System.EventHandler(this.btnCreateNew_Click);
            // 
            // btnCreateSimilar
            // 
            this.btnCreateSimilar.BackColor = System.Drawing.Color.White;
            this.btnCreateSimilar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCreateSimilar.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateSimilar.ForeColor = System.Drawing.Color.Navy;
            this.btnCreateSimilar.Location = new System.Drawing.Point(24, 223);
            this.btnCreateSimilar.Name = "btnCreateSimilar";
            this.btnCreateSimilar.Size = new System.Drawing.Size(207, 23);
            this.btnCreateSimilar.TabIndex = 4;
            this.btnCreateSimilar.Text = "CREATE SIMILAR NOTIFICATION";
            this.btnCreateSimilar.UseVisualStyleBackColor = false;
            this.btnCreateSimilar.Click += new System.EventHandler(this.btnCreateSimilar_Click);
            // 
            // Favourites
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(741, 276);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnCreateNew);
            this.Controls.Add(this.btnCreateSimilar);
            this.Controls.Add(this.grpFavourites);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "Favourites";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Favourites";
            this.Load += new System.EventHandler(this.Favourites_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgFavourites)).EndInit();
            this.grpFavourites.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgFavourites;
        private System.Windows.Forms.GroupBox grpFavourites;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnCreateNew;
        private System.Windows.Forms.Button btnCreateSimilar;
    }
}