namespace TestXLS
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
            b_cargar = new Button();
            Titulo = new Label();
            b_anadir = new Button();
            b_crear = new Button();
            textBox1 = new TextBox();
            SuspendLayout();
            // 
            // b_cargar
            // 
            b_cargar.Location = new Point(124, 190);
            b_cargar.Name = "b_cargar";
            b_cargar.Size = new Size(94, 29);
            b_cargar.TabIndex = 0;
            b_cargar.Text = "Cargar";
            b_cargar.UseVisualStyleBackColor = true;
            b_cargar.Click += b_cargar_Click;
            // 
            // Titulo
            // 
            Titulo.AutoSize = true;
            Titulo.Font = new Font("Segoe UI", 18F, FontStyle.Regular, GraphicsUnit.Point, 0);
            Titulo.Location = new Point(252, 41);
            Titulo.Name = "Titulo";
            Titulo.Size = new Size(293, 41);
            Titulo.TabIndex = 3;
            Titulo.Text = "Convertir XLS to CSV";
            // 
            // b_anadir
            // 
            b_anadir.Location = new Point(350, 190);
            b_anadir.Name = "b_anadir";
            b_anadir.Size = new Size(94, 29);
            b_anadir.TabIndex = 4;
            b_anadir.Text = "Añadir";
            b_anadir.UseVisualStyleBackColor = true;
            b_anadir.Click += b_anadir_Click;
            // 
            // b_crear
            // 
            b_crear.Location = new Point(553, 190);
            b_crear.Name = "b_crear";
            b_crear.Size = new Size(94, 29);
            b_crear.TabIndex = 5;
            b_crear.Text = "Crear";
            b_crear.UseVisualStyleBackColor = true;
            b_crear.Click += b_crear_Click;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(192, 124);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(419, 27);
            textBox1.TabIndex = 6;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(textBox1);
            Controls.Add(b_crear);
            Controls.Add(b_anadir);
            Controls.Add(Titulo);
            Controls.Add(b_cargar);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button b_cargar;
        private Label Titulo;
        private Button b_anadir;
        private Button b_crear;
        private TextBox textBox1;
    }
}
