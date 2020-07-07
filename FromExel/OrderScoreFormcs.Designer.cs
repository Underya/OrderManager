namespace FromExel
{
    partial class OrderScoreFormcs
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.просмотретьВсеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.Num = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.nameW = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.formatW = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.exem = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.countW = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.summW = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Number = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.NameAuth = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Execute = new System.Windows.Forms.DataGridViewButtonColumn();
            this.NumList = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Val = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.InfObj = new System.Windows.Forms.DataGridViewComboBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Number,
            this.Column1,
            this.NameAuth,
            this.Execute,
            this.NumList,
            this.Val,
            this.InfObj});
            this.dataGridView1.Location = new System.Drawing.Point(12, 80);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(431, 250);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_CellMouseClick);
            this.dataGridView1.SelectionChanged += new System.EventHandler(this.dataGridView1_SelectionChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.просмотретьВсеToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1057, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // просмотретьВсеToolStripMenuItem
            // 
            this.просмотретьВсеToolStripMenuItem.Name = "просмотретьВсеToolStripMenuItem";
            this.просмотретьВсеToolStripMenuItem.Size = new System.Drawing.Size(114, 20);
            this.просмотретьВсеToolStripMenuItem.Text = "Просмотреть все";
            this.просмотретьВсеToolStripMenuItem.Click += new System.EventHandler(this.просмотретьВсеToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(457, 57);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(107, 20);
            this.label1.TabIndex = 2;
            this.label1.Text = "Содержимое";
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.AllowUserToDeleteRows = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Num,
            this.nameW,
            this.formatW,
            this.exem,
            this.countW,
            this.summW});
            this.dataGridView2.Location = new System.Drawing.Point(461, 80);
            this.dataGridView2.MultiSelect = false;
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.ReadOnly = true;
            this.dataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView2.Size = new System.Drawing.Size(557, 250);
            this.dataGridView2.TabIndex = 3;
            // 
            // Num
            // 
            this.Num.HeaderText = "№";
            this.Num.Name = "Num";
            this.Num.ReadOnly = true;
            // 
            // nameW
            // 
            this.nameW.HeaderText = "Наименование";
            this.nameW.Name = "nameW";
            this.nameW.ReadOnly = true;
            // 
            // formatW
            // 
            this.formatW.HeaderText = "Формат";
            this.formatW.Name = "formatW";
            this.formatW.ReadOnly = true;
            // 
            // exem
            // 
            this.exem.HeaderText = "В еземпл";
            this.exem.Name = "exem";
            this.exem.ReadOnly = true;
            // 
            // countW
            // 
            this.countW.HeaderText = "Количество";
            this.countW.Name = "countW";
            this.countW.ReadOnly = true;
            // 
            // summW
            // 
            this.summW.HeaderText = "Сумма";
            this.summW.Name = "summW";
            this.summW.ReadOnly = true;
            // 
            // Number
            // 
            this.Number.HeaderText = "№";
            this.Number.Name = "Number";
            this.Number.ReadOnly = true;
            this.Number.Width = 107;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Дата";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            this.Column1.Width = 107;
            // 
            // NameAuth
            // 
            this.NameAuth.HeaderText = "Заказчик";
            this.NameAuth.Name = "NameAuth";
            this.NameAuth.ReadOnly = true;
            this.NameAuth.Width = 107;
            // 
            // Execute
            // 
            this.Execute.HeaderText = "Выполнить";
            this.Execute.Name = "Execute";
            this.Execute.ReadOnly = true;
            this.Execute.Width = 107;
            // 
            // NumList
            // 
            this.NumList.HeaderText = "NumList";
            this.NumList.Name = "NumList";
            this.NumList.ReadOnly = true;
            this.NumList.Visible = false;
            // 
            // Val
            // 
            this.Val.HeaderText = "Val";
            this.Val.Name = "Val";
            this.Val.ReadOnly = true;
            this.Val.Visible = false;
            // 
            // InfObj
            // 
            this.InfObj.HeaderText = "InfObj";
            this.InfObj.Name = "InfObj";
            this.InfObj.ReadOnly = true;
            this.InfObj.Visible = false;
            // 
            // OrderScoreFormcs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1057, 369);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OrderScoreFormcs";
            this.Text = "Просмотр заказов";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.OrderScoreFormcs_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem просмотретьВсеToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Num;
        private System.Windows.Forms.DataGridViewTextBoxColumn nameW;
        private System.Windows.Forms.DataGridViewTextBoxColumn formatW;
        private System.Windows.Forms.DataGridViewTextBoxColumn exem;
        private System.Windows.Forms.DataGridViewTextBoxColumn countW;
        private System.Windows.Forms.DataGridViewTextBoxColumn summW;
        private System.Windows.Forms.DataGridViewTextBoxColumn Number;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn NameAuth;
        private System.Windows.Forms.DataGridViewButtonColumn Execute;
        private System.Windows.Forms.DataGridViewTextBoxColumn NumList;
        private System.Windows.Forms.DataGridViewTextBoxColumn Val;
        private System.Windows.Forms.DataGridViewComboBoxColumn InfObj;
    }
}