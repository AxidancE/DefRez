namespace Wall_def
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            this.button1 = new System.Windows.Forms.Button();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.файлToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.открытьExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.закрытьExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.дополнительноToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.создатьСтолбецToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.проверкаТаблицыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.показатьОкноExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.остальноеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.оПрограммеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Location = new System.Drawing.Point(47, 94);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 40);
            this.button1.TabIndex = 0;
            this.button1.Text = "Сделать общую схему";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(52, 137);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(91, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Excel подключен";
            this.label1.Visible = false;
            // 
            // файлToolStripMenuItem
            // 
            this.файлToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.открытьExcelToolStripMenuItem,
            this.закрытьExcelToolStripMenuItem});
            this.файлToolStripMenuItem.Name = "файлToolStripMenuItem";
            this.файлToolStripMenuItem.Size = new System.Drawing.Size(48, 20);
            this.файлToolStripMenuItem.Text = "Файл";
            // 
            // открытьExcelToolStripMenuItem
            // 
            this.открытьExcelToolStripMenuItem.Name = "открытьExcelToolStripMenuItem";
            this.открытьExcelToolStripMenuItem.Size = new System.Drawing.Size(150, 22);
            this.открытьExcelToolStripMenuItem.Text = "Открыть Excel";
            this.открытьExcelToolStripMenuItem.Click += new System.EventHandler(this.открытьExcelToolStripMenuItem_Click);
            // 
            // закрытьExcelToolStripMenuItem
            // 
            this.закрытьExcelToolStripMenuItem.Enabled = false;
            this.закрытьExcelToolStripMenuItem.Name = "закрытьExcelToolStripMenuItem";
            this.закрытьExcelToolStripMenuItem.Size = new System.Drawing.Size(150, 22);
            this.закрытьExcelToolStripMenuItem.Text = "Закрыть Excel";
            this.закрытьExcelToolStripMenuItem.Click += new System.EventHandler(this.закрытьExcelToolStripMenuItem_Click);
            // 
            // дополнительноToolStripMenuItem
            // 
            this.дополнительноToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.создатьСтолбецToolStripMenuItem,
            this.проверкаТаблицыToolStripMenuItem,
            this.показатьОкноExcelToolStripMenuItem,
            this.остальноеToolStripMenuItem});
            this.дополнительноToolStripMenuItem.Name = "дополнительноToolStripMenuItem";
            this.дополнительноToolStripMenuItem.Size = new System.Drawing.Size(107, 20);
            this.дополнительноToolStripMenuItem.Text = "Дополнительно";
            // 
            // создатьСтолбецToolStripMenuItem
            // 
            this.создатьСтолбецToolStripMenuItem.Enabled = false;
            this.создатьСтолбецToolStripMenuItem.Name = "создатьСтолбецToolStripMenuItem";
            this.создатьСтолбецToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.создатьСтолбецToolStripMenuItem.Text = "Создать столбец";
            this.создатьСтолбецToolStripMenuItem.Click += new System.EventHandler(this.создатьСтолбецToolStripMenuItem_Click);
            // 
            // проверкаТаблицыToolStripMenuItem
            // 
            this.проверкаТаблицыToolStripMenuItem.Enabled = false;
            this.проверкаТаблицыToolStripMenuItem.Name = "проверкаТаблицыToolStripMenuItem";
            this.проверкаТаблицыToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.проверкаТаблицыToolStripMenuItem.Text = "Проверка таблицы";
            this.проверкаТаблицыToolStripMenuItem.Click += new System.EventHandler(this.проверкаТаблицыToolStripMenuItem_Click);
            // 
            // показатьОкноExcelToolStripMenuItem
            // 
            this.показатьОкноExcelToolStripMenuItem.Enabled = false;
            this.показатьОкноExcelToolStripMenuItem.Name = "показатьОкноExcelToolStripMenuItem";
            this.показатьОкноExcelToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.показатьОкноExcelToolStripMenuItem.Text = "Показать окно Excel";
            this.показатьОкноExcelToolStripMenuItem.Click += new System.EventHandler(this.показатьОкноExcelToolStripMenuItem_Click);
            // 
            // остальноеToolStripMenuItem
            // 
            this.остальноеToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.оПрограммеToolStripMenuItem});
            this.остальноеToolStripMenuItem.Name = "остальноеToolStripMenuItem";
            this.остальноеToolStripMenuItem.Size = new System.Drawing.Size(183, 22);
            this.остальноеToolStripMenuItem.Text = "Остальное";
            // 
            // оПрограммеToolStripMenuItem
            // 
            this.оПрограммеToolStripMenuItem.Name = "оПрограммеToolStripMenuItem";
            this.оПрограммеToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.оПрограммеToolStripMenuItem.Text = "О программе";
            this.оПрограммеToolStripMenuItem.Click += new System.EventHandler(this.оПрограммеToolStripMenuItem_Click_1);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.файлToolStripMenuItem,
            this.дополнительноToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(194, 24);
            this.menuStrip1.TabIndex = 4;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // comboBox1
            // 
            this.comboBox1.Enabled = false;
            this.comboBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "Стенка"});
            this.comboBox1.Location = new System.Drawing.Point(47, 39);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(100, 21);
            this.comboBox1.TabIndex = 6;
            // 
            // comboBox2
            // 
            this.comboBox2.Enabled = false;
            this.comboBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Strikeout, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "УЗК"});
            this.comboBox2.Location = new System.Drawing.Point(47, 67);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(100, 21);
            this.comboBox2.TabIndex = 7;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(194, 159);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.button1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "DefRez";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStripMenuItem файлToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem открытьExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem закрытьExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem дополнительноToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem создатьСтолбецToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem проверкаТаблицыToolStripMenuItem;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem показатьОкноExcelToolStripMenuItem;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ToolStripMenuItem остальноеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem оПрограммеToolStripMenuItem;
    }
}

