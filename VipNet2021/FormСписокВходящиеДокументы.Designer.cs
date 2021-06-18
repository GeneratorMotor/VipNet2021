namespace RegKor
{
    partial class FormСписокВходящиеДокументы
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose ( bool disposing )
        {
            if ( disposing && ( components != null ) )
            {
                components.Dispose( );
            }
            base.Dispose( disposing );
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent ( )
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormСписокВходящиеДокументы));
            this.textBoxПоиск = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.buttonОтмена = new System.Windows.Forms.Button();
            this.buttonСохранить = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.listBoxДокументы = new System.Windows.Forms.ListBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxПоиск
            // 
            this.textBoxПоиск.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.textBoxПоиск.Location = new System.Drawing.Point(0, 25);
            this.textBoxПоиск.Name = "textBoxПоиск";
            this.textBoxПоиск.Size = new System.Drawing.Size(207, 20);
            this.textBoxПоиск.TabIndex = 0;
            this.textBoxПоиск.TextChanged += new System.EventHandler(this.textBoxПоиск_TextChanged);
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(207, 25);
            this.label1.TabIndex = 2;
            this.label1.Text = "Введите текст для поиска:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.textBoxПоиск);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(211, 49);
            this.panel1.TabIndex = 3;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.buttonОтмена);
            this.panel2.Controls.Add(this.buttonСохранить);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 401);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(211, 41);
            this.panel2.TabIndex = 4;
            // 
            // buttonОтмена
            // 
            this.buttonОтмена.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonОтмена.Dock = System.Windows.Forms.DockStyle.Right;
            this.buttonОтмена.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonОтмена.Location = new System.Drawing.Point(108, 0);
            this.buttonОтмена.Name = "buttonОтмена";
            this.buttonОтмена.Size = new System.Drawing.Size(99, 37);
            this.buttonОтмена.TabIndex = 23;
            this.buttonОтмена.Text = "Отмена";
            this.buttonОтмена.Click += new System.EventHandler(this.buttonОтмена_Click);
            // 
            // buttonСохранить
            // 
            this.buttonСохранить.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonСохранить.Dock = System.Windows.Forms.DockStyle.Left;
            this.buttonСохранить.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonСохранить.Location = new System.Drawing.Point(0, 0);
            this.buttonСохранить.Name = "buttonСохранить";
            this.buttonСохранить.Size = new System.Drawing.Size(99, 37);
            this.buttonСохранить.TabIndex = 22;
            this.buttonСохранить.Text = "ОК";
            this.buttonСохранить.Click += new System.EventHandler(this.buttonСохранить_Click);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.panel4);
            this.panel3.Controls.Add(this.panel1);
            this.panel3.Controls.Add(this.panel2);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(211, 442);
            this.panel3.TabIndex = 5;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.listBoxДокументы);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(0, 49);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(211, 352);
            this.panel4.TabIndex = 5;
            // 
            // listBoxДокументы
            // 
            this.listBoxДокументы.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxДокументы.FormattingEnabled = true;
            this.listBoxДокументы.Location = new System.Drawing.Point(0, 0);
            this.listBoxДокументы.Name = "listBoxДокументы";
            this.listBoxДокументы.Size = new System.Drawing.Size(211, 342);
            this.listBoxДокументы.TabIndex = 1;
            this.listBoxДокументы.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listBoxДокументы_MouseDoubleClick);
            this.listBoxДокументы.Leave += new System.EventHandler(this.listBoxДокументы_Leave);
            this.listBoxДокументы.MouseDown += new System.Windows.Forms.MouseEventHandler(this.listBoxДокументы_MouseDown);
            // 
            // toolTip1
            // 
            this.toolTip1.UseAnimation = false;
            this.toolTip1.UseFading = false;
            // 
            // FormСписокВходящиеДокументы
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonОтмена;
            this.ClientSize = new System.Drawing.Size(211, 442);
            this.Controls.Add(this.panel3);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormСписокВходящиеДокументы";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Список входящих документов";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxПоиск;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button buttonОтмена;
        private System.Windows.Forms.Button buttonСохранить;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.ListBox listBoxДокументы;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}