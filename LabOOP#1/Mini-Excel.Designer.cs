
namespace LabaOOP1
{
    partial class MiniExcel
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dgv = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.FormatTableToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AddRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AddColumnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DeleteRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DeleteColumnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.FormatCellsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ShowExpressionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ShowValuesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ProcessTableToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.OpenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SaveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.NewTableToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.довідкаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.HelpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ApplyButton = new System.Windows.Forms.Button();
            this.TextBox = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgv
            // 
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Location = new System.Drawing.Point(12, 116);
            this.dgv.Name = "dgv";
            this.dgv.RowHeadersWidth = 51;
            this.dgv.RowTemplate.Height = 29;
            this.dgv.Size = new System.Drawing.Size(1714, 509);
            this.dgv.TabIndex = 0;
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.FormatTableToolStripMenuItem,
            this.FormatCellsToolStripMenuItem,
            this.ProcessTableToolStripMenuItem,
            this.довідкаToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1738, 28);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // FormatTableToolStripMenuItem
            // 
            this.FormatTableToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AddRowToolStripMenuItem,
            this.AddColumnToolStripMenuItem,
            this.DeleteRowToolStripMenuItem,
            this.DeleteColumnToolStripMenuItem});
            this.FormatTableToolStripMenuItem.Name = "FormatTableToolStripMenuItem";
            this.FormatTableToolStripMenuItem.Size = new System.Drawing.Size(183, 24);
            this.FormatTableToolStripMenuItem.Text = "Форматування таблиці";
            // 
            // AddRowToolStripMenuItem
            // 
            this.AddRowToolStripMenuItem.Name = "AddRowToolStripMenuItem";
            this.AddRowToolStripMenuItem.Size = new System.Drawing.Size(226, 26);
            this.AddRowToolStripMenuItem.Text = "Додати рядок";
            this.AddRowToolStripMenuItem.Click += new System.EventHandler(this.AddRowToolStripMenuItem_Click);
            // 
            // AddColumnToolStripMenuItem
            // 
            this.AddColumnToolStripMenuItem.Name = "AddColumnToolStripMenuItem";
            this.AddColumnToolStripMenuItem.Size = new System.Drawing.Size(226, 26);
            this.AddColumnToolStripMenuItem.Text = "Додати стовбець";
            this.AddColumnToolStripMenuItem.Click += new System.EventHandler(this.AddColumnToolStripMenuItem_Click);
            // 
            // DeleteRowToolStripMenuItem
            // 
            this.DeleteRowToolStripMenuItem.Name = "DeleteRowToolStripMenuItem";
            this.DeleteRowToolStripMenuItem.Size = new System.Drawing.Size(226, 26);
            this.DeleteRowToolStripMenuItem.Text = "Видалити рядок";
            this.DeleteRowToolStripMenuItem.Click += new System.EventHandler(this.DeleteRowToolStripMenuItem_Click);
            // 
            // DeleteColumnToolStripMenuItem
            // 
            this.DeleteColumnToolStripMenuItem.Name = "DeleteColumnToolStripMenuItem";
            this.DeleteColumnToolStripMenuItem.Size = new System.Drawing.Size(226, 26);
            this.DeleteColumnToolStripMenuItem.Text = "Видалити стовбець";
            this.DeleteColumnToolStripMenuItem.Click += new System.EventHandler(this.DeleteColumnToolStripMenuItem_Click);
            // 
            // FormatCellsToolStripMenuItem
            // 
            this.FormatCellsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ShowExpressionsToolStripMenuItem,
            this.ShowValuesToolStripMenuItem});
            this.FormatCellsToolStripMenuItem.Name = "FormatCellsToolStripMenuItem";
            this.FormatCellsToolStripMenuItem.Size = new System.Drawing.Size(134, 24);
            this.FormatCellsToolStripMenuItem.Text = "Вигляд клітинок";
            // 
            // ShowExpressionsToolStripMenuItem
            // 
            this.ShowExpressionsToolStripMenuItem.Name = "ShowExpressionsToolStripMenuItem";
            this.ShowExpressionsToolStripMenuItem.Size = new System.Drawing.Size(227, 26);
            this.ShowExpressionsToolStripMenuItem.Text = "Показати вирази";
            this.ShowExpressionsToolStripMenuItem.Click += new System.EventHandler(this.ShowExpressionsToolStripMenuItem_Click);
            // 
            // ShowValuesToolStripMenuItem
            // 
            this.ShowValuesToolStripMenuItem.Name = "ShowValuesToolStripMenuItem";
            this.ShowValuesToolStripMenuItem.Size = new System.Drawing.Size(227, 26);
            this.ShowValuesToolStripMenuItem.Text = "Показати значення";
            this.ShowValuesToolStripMenuItem.Click += new System.EventHandler(this.ShowValuesToolStripMenuItem_Click);
            // 
            // ProcessTableToolStripMenuItem
            // 
            this.ProcessTableToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.OpenToolStripMenuItem,
            this.SaveToolStripMenuItem,
            this.NewTableToolStripMenuItem});
            this.ProcessTableToolStripMenuItem.Name = "ProcessTableToolStripMenuItem";
            this.ProcessTableToolStripMenuItem.Size = new System.Drawing.Size(146, 24);
            this.ProcessTableToolStripMenuItem.Text = "Обробка таблиць";
            // 
            // OpenToolStripMenuItem
            // 
            this.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem";
            this.OpenToolStripMenuItem.Size = new System.Drawing.Size(189, 26);
            this.OpenToolStripMenuItem.Text = "Відкрити";
            this.OpenToolStripMenuItem.Click += new System.EventHandler(this.OpenToolStripMenuItem_Click);
            // 
            // SaveToolStripMenuItem
            // 
            this.SaveToolStripMenuItem.Name = "SaveToolStripMenuItem";
            this.SaveToolStripMenuItem.Size = new System.Drawing.Size(189, 26);
            this.SaveToolStripMenuItem.Text = "Зберегти";
            this.SaveToolStripMenuItem.Click += new System.EventHandler(this.SaveToolStripMenuItem_Click);
            // 
            // NewTableToolStripMenuItem
            // 
            this.NewTableToolStripMenuItem.Name = "NewTableToolStripMenuItem";
            this.NewTableToolStripMenuItem.Size = new System.Drawing.Size(189, 26);
            this.NewTableToolStripMenuItem.Text = "Нова таблиця";
            this.NewTableToolStripMenuItem.Click += new System.EventHandler(this.NewTableToolStripMenuItem_Click);
            // 
            // довідкаToolStripMenuItem
            // 
            this.довідкаToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AboutToolStripMenuItem,
            this.HelpToolStripMenuItem});
            this.довідкаToolStripMenuItem.Name = "довідкаToolStripMenuItem";
            this.довідкаToolStripMenuItem.Size = new System.Drawing.Size(77, 24);
            this.довідкаToolStripMenuItem.Text = "Довідка";
            // 
            // AboutToolStripMenuItem
            // 
            this.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem";
            this.AboutToolStripMenuItem.Size = new System.Drawing.Size(252, 26);
            this.AboutToolStripMenuItem.Text = "Про програму";
            this.AboutToolStripMenuItem.Click += new System.EventHandler(this.AboutToolStripMenuItem_Click);
            // 
            // HelpToolStripMenuItem
            // 
            this.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem";
            this.HelpToolStripMenuItem.Size = new System.Drawing.Size(252, 26);
            this.HelpToolStripMenuItem.Text = "Допомога користувачу";
            this.HelpToolStripMenuItem.Click += new System.EventHandler(this.HelpToolStripMenuItem_Click);
            // 
            // ApplyButton
            // 
            this.ApplyButton.Location = new System.Drawing.Point(371, 68);
            this.ApplyButton.Name = "ApplyButton";
            this.ApplyButton.Size = new System.Drawing.Size(94, 29);
            this.ApplyButton.TabIndex = 2;
            this.ApplyButton.Text = "Обчислити";
            this.ApplyButton.UseVisualStyleBackColor = true;
            this.ApplyButton.Click += new System.EventHandler(this.ApplyButton_Click);
            // 
            // TextBox
            // 
            this.TextBox.Location = new System.Drawing.Point(209, 68);
            this.TextBox.Name = "TextBox";
            this.TextBox.Size = new System.Drawing.Size(156, 27);
            this.TextBox.TabIndex = 3;
            // 
            // MiniExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1738, 639);
            this.Controls.Add(this.TextBox);
            this.Controls.Add(this.ApplyButton);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MiniExcel";
            this.Text = "Міні-Ексель";
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem FormatTableToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem AddRowToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem AddColumnToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem FormatCellsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ShowExpressionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ShowValuesToolStripMenuItem;
        private System.Windows.Forms.Button ApplyButton;
        private System.Windows.Forms.TextBox TextBox;
        private System.Windows.Forms.ToolStripMenuItem ProcessTableToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem OpenToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem SaveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem довідкаToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem AboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem HelpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem DeleteRowToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem DeleteColumnToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem NewTableToolStripMenuItem;
    }
}

