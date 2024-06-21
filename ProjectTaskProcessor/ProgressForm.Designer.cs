using System;
using System.Windows.Forms;

namespace ProjectTaskProcessor
{
    partial class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label labelProgress;

        private int _totalTasks;
        private int _currentTask = 0;

        public ProgressForm(int totalTasks)
        {
            _totalTasks = totalTasks;
            InitializeComponent();
            progressBar.Maximum = _totalTasks;
        }

        public void UpdateProgress()
        {
            _currentTask++;
            if (_currentTask <= _totalTasks)
            {
                progressBar.Value = _currentTask;
                labelProgress.Text = $"Эмуляция высоконагруженной работы!" +
                    $"\n\n" +
                    $"Обрабатывается задача {_currentTask} из {_totalTasks}";
            }
        }

        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void ProgressForm_Load(object sender, EventArgs e)
        {
            CenterControls();
        }

        private void CenterControls()
        {
            progressBar.Location = new System.Drawing.Point(
                (this.ClientSize.Width - progressBar.Width) / 2,
                (this.ClientSize.Height - progressBar.Height) / 2 + 30
            );

            labelProgress.Location = new System.Drawing.Point(
                (this.ClientSize.Width - labelProgress.Width) / 2,
                progressBar.Top - labelProgress.Height - 40
            );
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.labelProgress = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBar
            // 
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(260, 23);
            this.progressBar.TabIndex = 0;
            // 
            // labelProgress
            // 
            this.labelProgress.AutoSize = true;
            this.labelProgress.ForeColor = System.Drawing.Color.Crimson;
            this.labelProgress.Name = "labelProgress";
            this.labelProgress.TabIndex = 1;
            this.labelProgress.Text = "Симуляция продолжительной работы";
            // 
            // ProgressForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 150);
            this.Controls.Add(this.labelProgress);
            this.Controls.Add(this.progressBar);
            this.Name = "ProgressForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Progress";
            this.Load += new System.EventHandler(this.ProgressForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
    }
}
