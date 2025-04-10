namespace WinFormsApp1
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            dataGridView1 = new System.Windows.Forms.DataGridView();
            button1 = new System.Windows.Forms.Button();
            button2 = new System.Windows.Forms.Button();
            Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();

            ((System.ComponentModel.ISupportInitialize)(dataGridView1)).BeginInit();
            SuspendLayout();

            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            Column1,
            Column2,
            Column3});
            dataGridView1.Location = new System.Drawing.Point(50, 20);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.ReadOnly = true;
            dataGridView1.Size = new System.Drawing.Size(444, 183);
            dataGridView1.TabIndex = 0;

            // 
            // Column1
            // 
            Column1.HeaderText = "Name";
            Column1.Name = "Column1";
            Column1.ReadOnly = true;

            // 
            // Column2
            // 
            Column2.HeaderText = "Reg no";
            Column2.Name = "Column2";
            Column2.ReadOnly = true;

            // 
            // Column3
            // 
            Column3.HeaderText = "Department";
            Column3.Name = "Column3";
            Column3.ReadOnly = true;

            // 
            // button1
            // 
            button1.Location = new System.Drawing.Point(50, 220);
            button1.Name = "button1";
            button1.Size = new System.Drawing.Size(200, 35);
            button1.TabIndex = 1;
            button1.Text = "Import Excel to DataGrid";
            button1.UseVisualStyleBackColor = true;
            button1.Click += new System.EventHandler(this.Button1_Click);

            // 
            // button2
            // 
            button2.Location = new System.Drawing.Point(280, 220);
            button2.Name = "button2";
            button2.Size = new System.Drawing.Size(200, 35);
            button2.TabIndex = 2;
            button2.Text = "Save to Database";
            button2.UseVisualStyleBackColor = true;
            button2.Click += new System.EventHandler(this.Button2_Click);

            // 
            // Form1
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(550, 300);
            Controls.Add(dataGridView1);
            Controls.Add(button1);
            Controls.Add(button2);
            Name = "Form1";
            Text = "Excel to SQL Importer";
            ((System.ComponentModel.ISupportInitialize)(dataGridView1)).EndInit();
            ResumeLayout(false);
        }

        #endregion
    }
}
