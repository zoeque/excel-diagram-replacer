namespace ExcelDiagramReplacer
{
	partial class ExcelDiagramReplacer
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
            button1 = new Button();
            textBoxPath = new TextBox();
            label1 = new Label();
            button2 = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(302, 211);
            button1.Margin = new Padding(4, 5, 4, 5);
            button1.Name = "button1";
            button1.Size = new Size(156, 39);
            button1.TabIndex = 0;
            button1.Text = "テキスト生成";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // textBoxPath
            // 
            textBoxPath.AllowDrop = true;
            textBoxPath.Location = new Point(32, 131);
            textBoxPath.Margin = new Padding(4, 5, 4, 5);
            textBoxPath.Name = "textBoxPath";
            textBoxPath.Size = new Size(424, 31);
            textBoxPath.TabIndex = 1;
            textBoxPath.DragDrop += textBox1_DragDrop;
            textBoxPath.DragEnter += textBox1_DragEnter;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(31, 101);
            label1.Margin = new Padding(4, 0, 4, 0);
            label1.Name = "label1";
            label1.Size = new Size(453, 25);
            label1.TabIndex = 2;
            label1.Text = "変換対象のエクセルを以下にドラッグアンドドロップしてください。";
            // 
            // button2
            // 
            button2.Location = new Point(302, 282);
            button2.Margin = new Padding(4, 4, 4, 4);
            button2.Name = "button2";
            button2.Size = new Size(155, 39);
            button2.TabIndex = 3;
            button2.Text = "変換開始";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // ExcelDiagramReplacer
            // 
            AllowDrop = true;
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(548, 370);
            Controls.Add(button2);
            Controls.Add(label1);
            Controls.Add(textBoxPath);
            Controls.Add(button1);
            Margin = new Padding(4, 5, 4, 5);
            Name = "ExcelDiagramReplacer";
            Text = "エクセル図形内文字変換ツール";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
		private TextBox textBoxPath;
		private Label label1;
        private Button button2;
    }
}