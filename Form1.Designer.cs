namespace COTACAO_INSUMO
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;

            ClientSize = new Size(1400, 850);

            Name = "Form1";

            Text = "Sistema de cotação de insumos";

            Load += Form1_Load;

            ResumeLayout(false);
        }
    }
}