namespace editdocuments
{
    partial class ProcessDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProcessDialog));
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.countProgressBar = new System.Windows.Forms.ProgressBar();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.counterTotalLabel = new System.Windows.Forms.Label();
            this.separatorLabel = new System.Windows.Forms.Label();
            this.counterValueLabel = new System.Windows.Forms.Label();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.closeButton = new System.Windows.Forms.Button();
            this.tableLayoutPanel5 = new System.Windows.Forms.TableLayoutPanel();
            this.currentFileLabel = new System.Windows.Forms.Label();
            this.fileNameLabel = new System.Windows.Forms.Label();
            this.logTextBox = new System.Windows.Forms.RichTextBox();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.tableLayoutPanel5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel2
            // 
            resources.ApplyResources(this.tableLayoutPanel2, "tableLayoutPanel2");
            this.tableLayoutPanel2.Controls.Add(this.countProgressBar, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel4, 1, 0);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            // 
            // countProgressBar
            // 
            resources.ApplyResources(this.countProgressBar, "countProgressBar");
            this.countProgressBar.Name = "countProgressBar";
            // 
            // tableLayoutPanel4
            // 
            resources.ApplyResources(this.tableLayoutPanel4, "tableLayoutPanel4");
            this.tableLayoutPanel4.Controls.Add(this.counterTotalLabel, 2, 0);
            this.tableLayoutPanel4.Controls.Add(this.separatorLabel, 1, 0);
            this.tableLayoutPanel4.Controls.Add(this.counterValueLabel, 0, 0);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            // 
            // counterTotalLabel
            // 
            resources.ApplyResources(this.counterTotalLabel, "counterTotalLabel");
            this.counterTotalLabel.Name = "counterTotalLabel";
            // 
            // separatorLabel
            // 
            resources.ApplyResources(this.separatorLabel, "separatorLabel");
            this.separatorLabel.Name = "separatorLabel";
            // 
            // counterValueLabel
            // 
            resources.ApplyResources(this.counterValueLabel, "counterValueLabel");
            this.counterValueLabel.Name = "counterValueLabel";
            // 
            // tableLayoutPanel1
            // 
            resources.ApplyResources(this.tableLayoutPanel1, "tableLayoutPanel1");
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel3, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel5, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.logTextBox, 0, 2);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            // 
            // tableLayoutPanel3
            // 
            resources.ApplyResources(this.tableLayoutPanel3, "tableLayoutPanel3");
            this.tableLayoutPanel3.Controls.Add(this.closeButton, 1, 0);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            // 
            // closeButton
            // 
            resources.ApplyResources(this.closeButton, "closeButton");
            this.closeButton.Name = "closeButton";
            this.closeButton.UseVisualStyleBackColor = true;
            this.closeButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // tableLayoutPanel5
            // 
            resources.ApplyResources(this.tableLayoutPanel5, "tableLayoutPanel5");
            this.tableLayoutPanel5.Controls.Add(this.currentFileLabel, 0, 0);
            this.tableLayoutPanel5.Controls.Add(this.fileNameLabel, 0, 1);
            this.tableLayoutPanel5.Name = "tableLayoutPanel5";
            // 
            // currentFileLabel
            // 
            resources.ApplyResources(this.currentFileLabel, "currentFileLabel");
            this.currentFileLabel.Name = "currentFileLabel";
            // 
            // fileNameLabel
            // 
            resources.ApplyResources(this.fileNameLabel, "fileNameLabel");
            this.fileNameLabel.Name = "fileNameLabel";
            // 
            // logTextBox
            // 
            resources.ApplyResources(this.logTextBox, "logTextBox");
            this.logTextBox.Name = "logTextBox";
            // 
            // ProcessDialog
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "ProcessDialog";
            this.ShowIcon = false;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ProcessDialog_FormClosing);
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel4.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel5.ResumeLayout(false);
            this.tableLayoutPanel5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.ProgressBar countProgressBar;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.Label counterTotalLabel;
        private System.Windows.Forms.Label separatorLabel;
        private System.Windows.Forms.Label counterValueLabel;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.Button closeButton;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel5;
        private System.Windows.Forms.Label currentFileLabel;
        private System.Windows.Forms.Label fileNameLabel;
        private System.Windows.Forms.RichTextBox logTextBox;
    }
}