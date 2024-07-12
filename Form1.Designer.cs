namespace ExcelVTEntegrasyon
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            richTextBox1 = new RichTextBox();
            richTextBox2 = new RichTextBox();
            BtnVTdenOku = new Button();
            BtnExceldenOku = new Button();
            SuspendLayout();
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(12, 72);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(467, 120);
            richTextBox1.TabIndex = 0;
            richTextBox1.Text = "";
            // 
            // richTextBox2
            // 
            richTextBox2.Location = new Point(12, 241);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.Size = new Size(467, 120);
            richTextBox2.TabIndex = 1;
            richTextBox2.Text = "";
            // 
            // BtnVTdenOku
            // 
            BtnVTdenOku.Location = new Point(497, 72);
            BtnVTdenOku.Name = "BtnVTdenOku";
            BtnVTdenOku.Size = new Size(137, 69);
            BtnVTdenOku.TabIndex = 2;
            BtnVTdenOku.Text = "Veri Tabanından Oku Excel'e Yaz";
            BtnVTdenOku.UseVisualStyleBackColor = true;
            BtnVTdenOku.Click += BtnVTdenOku_Click;
            // 
            // BtnExceldenOku
            // 
            BtnExceldenOku.Location = new Point(497, 241);
            BtnExceldenOku.Name = "BtnExceldenOku";
            BtnExceldenOku.Size = new Size(137, 69);
            BtnExceldenOku.TabIndex = 3;
            BtnExceldenOku.Text = "Excel'den Oku Veri Tabanına Yaz";
            BtnExceldenOku.UseVisualStyleBackColor = true;
            BtnExceldenOku.Click += BtnExceldenOku_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ActiveCaption;
            ClientSize = new Size(800, 450);
            Controls.Add(BtnExceldenOku);
            Controls.Add(BtnVTdenOku);
            Controls.Add(richTextBox2);
            Controls.Add(richTextBox1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Excel'den VT'ye, VT'den Excele Veri Aktarımı";
            ResumeLayout(false);
        }

        #endregion

        private RichTextBox richTextBox1;
        private RichTextBox richTextBox2;
        private Button BtnVTdenOku;
        private Button BtnExceldenOku;
    }
}
