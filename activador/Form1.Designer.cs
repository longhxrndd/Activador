using System.Drawing;
using System.Windows.Forms;

namespace activador
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.terminalRichTextBox = new System.Windows.Forms.RichTextBox();
            this.materialButton1 = new MaterialSkin.Controls.MaterialButton();
            this.progressBar = new MaterialSkin.Controls.MaterialProgressBar();
            this.materialProgressBar1 = new MaterialSkin.Controls.MaterialProgressBar();
            this.materialButton2 = new MaterialSkin.Controls.MaterialButton();
            this.SuspendLayout();
            // 
            // terminalRichTextBox
            // 
            this.terminalRichTextBox.BackColor = System.Drawing.Color.Black;
            this.terminalRichTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.terminalRichTextBox.Font = new System.Drawing.Font("Consolas", 10F);
            this.terminalRichTextBox.ForeColor = System.Drawing.Color.White;
            this.terminalRichTextBox.Location = new System.Drawing.Point(178, 192);
            this.terminalRichTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.terminalRichTextBox.Name = "terminalRichTextBox";
            this.terminalRichTextBox.ReadOnly = true;
            this.terminalRichTextBox.Size = new System.Drawing.Size(595, 322);
            this.terminalRichTextBox.TabIndex = 1;
            this.terminalRichTextBox.Text = "";
            // 
            // materialButton1
            // 
            this.materialButton1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.materialButton1.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.materialButton1.Depth = 0;
            this.materialButton1.HighEmphasis = true;
            this.materialButton1.Icon = null;
            this.materialButton1.Location = new System.Drawing.Point(398, 15);
            this.materialButton1.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.materialButton1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialButton1.Name = "materialButton1";
            this.materialButton1.NoAccentTextColor = System.Drawing.Color.Empty;
            this.materialButton1.Size = new System.Drawing.Size(135, 36);
            this.materialButton1.TabIndex = 3;
            this.materialButton1.Text = "Activar office";
            this.materialButton1.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.materialButton1.UseAccentColor = false;
            this.materialButton1.UseVisualStyleBackColor = true;
            this.materialButton1.Click += new System.EventHandler(this.materialButton1_Click);
            // 
            // progressBar
            // 
            this.progressBar.Depth = 0;
            this.progressBar.Location = new System.Drawing.Point(0, 0);
            this.progressBar.MouseState = MaterialSkin.MouseState.HOVER;
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(100, 5);
            this.progressBar.TabIndex = 0;
            // 
            // materialProgressBar1
            // 
            this.materialProgressBar1.Depth = 0;
            this.materialProgressBar1.Location = new System.Drawing.Point(178, 125);
            this.materialProgressBar1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialProgressBar1.Name = "materialProgressBar1";
            this.materialProgressBar1.Size = new System.Drawing.Size(595, 5);
            this.materialProgressBar1.TabIndex = 4;
            // 
            // materialButton2
            // 
            this.materialButton2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.materialButton2.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.materialButton2.Depth = 0;
            this.materialButton2.HighEmphasis = true;
            this.materialButton2.Icon = null;
            this.materialButton2.Location = new System.Drawing.Point(387, 63);
            this.materialButton2.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.materialButton2.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialButton2.Name = "materialButton2";
            this.materialButton2.NoAccentTextColor = System.Drawing.Color.Empty;
            this.materialButton2.Size = new System.Drawing.Size(158, 36);
            this.materialButton2.TabIndex = 5;
            this.materialButton2.Text = "Activar windows";
            this.materialButton2.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.materialButton2.UseAccentColor = false;
            this.materialButton2.UseVisualStyleBackColor = true;
            this.materialButton2.Click += new System.EventHandler(this.materialButton2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 554);
            this.Controls.Add(this.materialButton2);
            this.Controls.Add(this.materialProgressBar1);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.materialButton1);
            this.Controls.Add(this.terminalRichTextBox);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "Office Activation";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.RichTextBox terminalRichTextBox;
        private MaterialSkin.Controls.MaterialButton materialButton1;
        private MaterialSkin.Controls.MaterialProgressBar progressBar;
        private MaterialSkin.Controls.MaterialProgressBar materialProgressBar1;
        private MaterialSkin.Controls.MaterialButton materialButton2;
    }
}

