namespace XmlToXls_3
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBox_SourceAddress = new System.Windows.Forms.TextBox();
            this.textBox_TargetAddress = new System.Windows.Forms.TextBox();
            this.button_SourceAddress = new System.Windows.Forms.Button();
            this.button_TargetAddress = new System.Windows.Forms.Button();
            this.button_Start = new System.Windows.Forms.Button();
            this.button_UnPack = new System.Windows.Forms.Button();
            this.button_Processing = new System.Windows.Forms.Button();
            this.button_ReNameZip = new System.Windows.Forms.Button();
            this.label_NumericInfo = new System.Windows.Forms.Label();
            this.button_UnPack_Rename = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.label_Progress = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // textBox_SourceAddress
            // 
            this.textBox_SourceAddress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_SourceAddress.Location = new System.Drawing.Point(12, 24);
            this.textBox_SourceAddress.Name = "textBox_SourceAddress";
            this.textBox_SourceAddress.Size = new System.Drawing.Size(358, 20);
            this.textBox_SourceAddress.TabIndex = 0;
            // 
            // textBox_TargetAddress
            // 
            this.textBox_TargetAddress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_TargetAddress.Location = new System.Drawing.Point(13, 66);
            this.textBox_TargetAddress.Name = "textBox_TargetAddress";
            this.textBox_TargetAddress.Size = new System.Drawing.Size(358, 20);
            this.textBox_TargetAddress.TabIndex = 1;
            // 
            // button_SourceAddress
            // 
            this.button_SourceAddress.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_SourceAddress.Location = new System.Drawing.Point(377, 22);
            this.button_SourceAddress.Name = "button_SourceAddress";
            this.button_SourceAddress.Size = new System.Drawing.Size(75, 23);
            this.button_SourceAddress.TabIndex = 2;
            this.button_SourceAddress.Text = "Обзор";
            this.button_SourceAddress.UseVisualStyleBackColor = true;
            this.button_SourceAddress.Click += new System.EventHandler(this.button_SourceAddress_Click);
            // 
            // button_TargetAddress
            // 
            this.button_TargetAddress.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_TargetAddress.Location = new System.Drawing.Point(377, 66);
            this.button_TargetAddress.Name = "button_TargetAddress";
            this.button_TargetAddress.Size = new System.Drawing.Size(75, 23);
            this.button_TargetAddress.TabIndex = 3;
            this.button_TargetAddress.Text = "Обзор";
            this.button_TargetAddress.UseVisualStyleBackColor = true;
            this.button_TargetAddress.Click += new System.EventHandler(this.button_TargetAddress_Click);
            // 
            // button_Start
            // 
            this.button_Start.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.button_Start.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button_Start.Location = new System.Drawing.Point(165, 95);
            this.button_Start.Name = "button_Start";
            this.button_Start.Size = new System.Drawing.Size(149, 51);
            this.button_Start.TabIndex = 4;
            this.button_Start.Text = "Искать файлы";
            this.button_Start.UseVisualStyleBackColor = true;
            this.button_Start.Click += new System.EventHandler(this.button_Start_Click);
            // 
            // button_UnPack
            // 
            this.button_UnPack.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button_UnPack.Enabled = false;
            this.button_UnPack.Location = new System.Drawing.Point(34, 273);
            this.button_UnPack.Name = "button_UnPack";
            this.button_UnPack.Size = new System.Drawing.Size(100, 36);
            this.button_UnPack.TabIndex = 5;
            this.button_UnPack.Text = "Распаковать архивы";
            this.button_UnPack.UseVisualStyleBackColor = true;
            this.button_UnPack.Click += new System.EventHandler(this.button_UnPack_Click);
            // 
            // button_Processing
            // 
            this.button_Processing.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.button_Processing.Enabled = false;
            this.button_Processing.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button_Processing.Location = new System.Drawing.Point(163, 165);
            this.button_Processing.Name = "button_Processing";
            this.button_Processing.Size = new System.Drawing.Size(149, 55);
            this.button_Processing.TabIndex = 6;
            this.button_Processing.Text = "Выгрузить реестры";
            this.button_Processing.UseVisualStyleBackColor = true;
            this.button_Processing.Click += new System.EventHandler(this.button_Processing_Click);
            // 
            // button_ReNameZip
            // 
            this.button_ReNameZip.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button_ReNameZip.Enabled = false;
            this.button_ReNameZip.Location = new System.Drawing.Point(330, 273);
            this.button_ReNameZip.Name = "button_ReNameZip";
            this.button_ReNameZip.Size = new System.Drawing.Size(100, 36);
            this.button_ReNameZip.TabIndex = 7;
            this.button_ReNameZip.Text = "Переименовать архивы";
            this.button_ReNameZip.UseVisualStyleBackColor = true;
            this.button_ReNameZip.Click += new System.EventHandler(this.button_ReNameZip_Click);
            // 
            // label_NumericInfo
            // 
            this.label_NumericInfo.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label_NumericInfo.AutoSize = true;
            this.label_NumericInfo.Location = new System.Drawing.Point(162, 149);
            this.label_NumericInfo.Name = "label_NumericInfo";
            this.label_NumericInfo.Size = new System.Drawing.Size(150, 13);
            this.label_NumericInfo.TabIndex = 8;
            this.label_NumericInfo.Text = "Найдено 0 XML и 0 архивов.";
            // 
            // button_UnPack_Rename
            // 
            this.button_UnPack_Rename.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button_UnPack_Rename.Enabled = false;
            this.button_UnPack_Rename.Location = new System.Drawing.Point(165, 273);
            this.button_UnPack_Rename.Name = "button_UnPack_Rename";
            this.button_UnPack_Rename.Size = new System.Drawing.Size(149, 36);
            this.button_UnPack_Rename.TabIndex = 9;
            this.button_UnPack_Rename.Text = "Распаковать и переименовать";
            this.button_UnPack_Rename.UseVisualStyleBackColor = true;
            this.button_UnPack_Rename.Click += new System.EventHandler(this.button_UnPack_Rename_Click);
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(34, 239);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(396, 23);
            this.progressBar.TabIndex = 10;
            // 
            // label_Progress
            // 
            this.label_Progress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label_Progress.AutoSize = true;
            this.label_Progress.Location = new System.Drawing.Point(31, 223);
            this.label_Progress.Name = "label_Progress";
            this.label_Progress.Size = new System.Drawing.Size(59, 13);
            this.label_Progress.TabIndex = 11;
            this.label_Progress.Text = "Прогресс:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(104, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "Папка назначения:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Исходная папка:";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(464, 321);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label_Progress);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.button_UnPack_Rename);
            this.Controls.Add(this.label_NumericInfo);
            this.Controls.Add(this.button_ReNameZip);
            this.Controls.Add(this.button_Processing);
            this.Controls.Add(this.button_UnPack);
            this.Controls.Add(this.button_Start);
            this.Controls.Add(this.button_TargetAddress);
            this.Controls.Add(this.button_SourceAddress);
            this.Controls.Add(this.textBox_TargetAddress);
            this.Controls.Add(this.textBox_SourceAddress);
            this.MinimumSize = new System.Drawing.Size(480, 360);
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox_SourceAddress;
        private System.Windows.Forms.TextBox textBox_TargetAddress;
        private System.Windows.Forms.Button button_SourceAddress;
        private System.Windows.Forms.Button button_TargetAddress;
        private System.Windows.Forms.Button button_Start;
        private System.Windows.Forms.Button button_UnPack;
        private System.Windows.Forms.Button button_Processing;
        private System.Windows.Forms.Button button_ReNameZip;
        private System.Windows.Forms.Label label_NumericInfo;
        private System.Windows.Forms.Button button_UnPack_Rename;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label label_Progress;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}

