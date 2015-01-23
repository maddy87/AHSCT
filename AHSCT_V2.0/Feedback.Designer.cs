namespace maddytry1
{
    partial class Feedback
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Feedback));
            this.rdbFeedback = new System.Windows.Forms.RadioButton();
            this.rdbBug = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.txtMessage = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // rdbFeedback
            // 
            this.rdbFeedback.AutoSize = true;
            this.rdbFeedback.Location = new System.Drawing.Point(12, 25);
            this.rdbFeedback.Name = "rdbFeedback";
            this.rdbFeedback.Size = new System.Drawing.Size(87, 17);
            this.rdbFeedback.TabIndex = 0;
            this.rdbFeedback.TabStop = true;
            this.rdbFeedback.Text = "FEEDBACK";
            this.rdbFeedback.UseVisualStyleBackColor = true;
            this.rdbFeedback.CheckedChanged += new System.EventHandler(this.rdbFeedback_CheckedChanged);
            // 
            // rdbBug
            // 
            this.rdbBug.AutoSize = true;
            this.rdbBug.Location = new System.Drawing.Point(143, 25);
            this.rdbBug.Name = "rdbBug";
            this.rdbBug.Size = new System.Drawing.Size(112, 17);
            this.rdbBug.TabIndex = 1;
            this.rdbBug.TabStop = true;
            this.rdbBug.Text = "REPORT A BUG";
            this.rdbBug.UseVisualStyleBackColor = true;
            this.rdbBug.CheckedChanged += new System.EventHandler(this.rdbBug_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 60);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "MESSAGE";
            // 
            // txtMessage
            // 
            this.txtMessage.Location = new System.Drawing.Point(15, 77);
            this.txtMessage.Multiline = true;
            this.txtMessage.Name = "txtMessage";
            this.txtMessage.Size = new System.Drawing.Size(345, 184);
            this.txtMessage.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(143, 281);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "SEND";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Feedback
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(372, 316);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtMessage);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.rdbBug);
            this.Controls.Add(this.rdbFeedback);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Navy;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Feedback";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Feedback";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton rdbFeedback;
        private System.Windows.Forms.RadioButton rdbBug;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtMessage;
        private System.Windows.Forms.Button button1;
    }
}