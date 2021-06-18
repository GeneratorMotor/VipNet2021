namespace RegKor
{
    partial class FormДокументооборот
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dt2 = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dt1 = new System.Windows.Forms.DateTimePicker();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.radButAll = new System.Windows.Forms.RadioButton();
            this.radButtonRead = new System.Windows.Forms.RadioButton();
            this.radButtonNoRead = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dt2);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.dt1);
            this.groupBox1.Location = new System.Drawing.Point(2, 1);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(394, 97);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "За какой период вести статистику?";
            // 
            // dt2
            // 
            this.dt2.Location = new System.Drawing.Point(180, 56);
            this.dt2.Name = "dt2";
            this.dt2.Size = new System.Drawing.Size(142, 20);
            this.dt2.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(177, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(136, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Конец отчётного периода";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(142, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Начало отчётного периода";
            // 
            // dt1
            // 
            this.dt1.Location = new System.Drawing.Point(18, 56);
            this.dt1.Name = "dt1";
            this.dt1.Size = new System.Drawing.Size(142, 20);
            this.dt1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(402, 6);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(114, 39);
            this.button1.TabIndex = 2;
            this.button1.Text = "Готово";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(402, 57);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(114, 39);
            this.button2.TabIndex = 2;
            this.button2.Text = "Отмена";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // radButAll
            // 
            this.radButAll.AutoSize = true;
            this.radButAll.Location = new System.Drawing.Point(13, 105);
            this.radButAll.Name = "radButAll";
            this.radButAll.Size = new System.Drawing.Size(44, 17);
            this.radButAll.TabIndex = 3;
            this.radButAll.TabStop = true;
            this.radButAll.Text = "Все";
            this.radButAll.UseVisualStyleBackColor = true;
            // 
            // radButtonRead
            // 
            this.radButtonRead.AutoSize = true;
            this.radButtonRead.Location = new System.Drawing.Point(63, 105);
            this.radButtonRead.Name = "radButtonRead";
            this.radButtonRead.Size = new System.Drawing.Size(93, 17);
            this.radButtonRead.TabIndex = 4;
            this.radButtonRead.TabStop = true;
            this.radButtonRead.Text = "Прочитанные";
            this.radButtonRead.UseVisualStyleBackColor = true;
            // 
            // radButtonNoRead
            // 
            this.radButtonNoRead.AutoSize = true;
            this.radButtonNoRead.Location = new System.Drawing.Point(162, 105);
            this.radButtonNoRead.Name = "radButtonNoRead";
            this.radButtonNoRead.Size = new System.Drawing.Size(108, 17);
            this.radButtonNoRead.TabIndex = 5;
            this.radButtonNoRead.TabStop = true;
            this.radButtonNoRead.Text = "Не прочитанные";
            this.radButtonNoRead.UseVisualStyleBackColor = true;
            // 
            // FormДокументооборот
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(534, 137);
            this.Controls.Add(this.radButtonNoRead);
            this.Controls.Add(this.radButtonRead);
            this.Controls.Add(this.radButAll);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "FormДокументооборот";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Отчет по документообороту";
            this.Load += new System.EventHandler(this.FormДокументооборот_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DateTimePicker dt2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dt1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.RadioButton radButAll;
        private System.Windows.Forms.RadioButton radButtonRead;
        private System.Windows.Forms.RadioButton radButtonNoRead;
    }
}