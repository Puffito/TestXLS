namespace TestXLS
{
    partial class Form2
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
            nombre_Box = new ComboBox();
            label1 = new Label();
            nombrePLC = new TextBox();
            b_cargar = new Button();
            SuspendLayout();
            // 
            // nombre_Box
            // 
            nombre_Box.Cursor = Cursors.SizeAll;
            nombre_Box.FormattingEnabled = true;
            nombre_Box.Items.AddRange(new object[] { "durkopp", "tgw" });
            nombre_Box.Location = new Point(177, 85);
            nombre_Box.Name = "nombre_Box";
            nombre_Box.Size = new Size(463, 28);
            nombre_Box.TabIndex = 0;
            nombre_Box.Text = "durkopp";
            nombre_Box.SelectedIndexChanged += nombre_Box_SelectedIndexChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label1.Location = new Point(259, 173);
            label1.Name = "label1";
            label1.Size = new Size(141, 31);
            label1.TabIndex = 1;
            label1.Text = "Nombre PLC";
            // 
            // nombrePLC
            // 
            nombrePLC.Location = new Point(419, 177);
            nombrePLC.Name = "nombrePLC";
            nombrePLC.Size = new Size(132, 27);
            nombrePLC.TabIndex = 2;
            // 
            // b_cargar
            // 
            b_cargar.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            b_cargar.Location = new Point(341, 283);
            b_cargar.Name = "b_cargar";
            b_cargar.Size = new Size(130, 47);
            b_cargar.TabIndex = 3;
            b_cargar.Text = "Cargar";
            b_cargar.UseVisualStyleBackColor = true;
            b_cargar.Click += b_cargar_Click;
            // 
            // Form2
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(b_cargar);
            Controls.Add(nombrePLC);
            Controls.Add(label1);
            Controls.Add(nombre_Box);
            Name = "Form2";
            Text = "Form2";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private ComboBox nombre_Box;
        private Label label1;
        private TextBox nombrePLC;
        private Button b_cargar;
    }
}