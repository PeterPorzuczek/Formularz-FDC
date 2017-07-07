namespace Formularz_FDC
{
	partial class Form1
	{
		/// <summary>
		/// Wymagana zmienna projektanta.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Wyczyść wszystkie używane zasoby.
		/// </summary>
		/// <param name="disposing">prawda, jeżeli zarządzane zasoby powinny zostać zlikwidowane; Fałsz w przeciwnym wypadku.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Kod generowany przez Projektanta formularzy systemu Windows

		/// <summary>
		/// Wymagana metoda obsługi projektanta — nie należy modyfikować 
		/// zawartość tej metody z edytorem kodu.
		/// </summary>
		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			this.relativesBox = new System.Windows.Forms.GroupBox();
			this.relativesBtn = new System.Windows.Forms.Button();
			this.relativesTb = new System.Windows.Forms.TextBox();
			this.relativesLbl = new System.Windows.Forms.Label();
			this.descriptionBox = new System.Windows.Forms.GroupBox();
			this.descriptionLabel = new System.Windows.Forms.Label();
			this.relativesBox.SuspendLayout();
			this.descriptionBox.SuspendLayout();
			this.SuspendLayout();
			// 
			// relativesBox
			// 
			this.relativesBox.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.relativesBox.Controls.Add(this.relativesBtn);
			this.relativesBox.Controls.Add(this.relativesTb);
			this.relativesBox.Controls.Add(this.relativesLbl);
			this.relativesBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F);
			this.relativesBox.Location = new System.Drawing.Point(30, 179);
			this.relativesBox.Name = "relativesBox";
			this.relativesBox.Size = new System.Drawing.Size(400, 200);
			this.relativesBox.TabIndex = 0;
			this.relativesBox.TabStop = false;
			this.relativesBox.Text = "Formularz FDC";
			// 
			// relativesBtn
			// 
			this.relativesBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
			this.relativesBtn.Location = new System.Drawing.Point(28, 143);
			this.relativesBtn.Name = "relativesBtn";
			this.relativesBtn.Size = new System.Drawing.Size(337, 33);
			this.relativesBtn.TabIndex = 2;
			this.relativesBtn.Text = "Dalej";
			this.relativesBtn.UseVisualStyleBackColor = true;
			this.relativesBtn.Click += new System.EventHandler(this.relativesBtn_Click);
			// 
			// relativesTb
			// 
			this.relativesTb.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
			this.relativesTb.HideSelection = false;
			this.relativesTb.Location = new System.Drawing.Point(28, 94);
			this.relativesTb.MaxLength = 2;
			this.relativesTb.Name = "relativesTb";
			this.relativesTb.Size = new System.Drawing.Size(337, 26);
			this.relativesTb.TabIndex = 1;
			this.relativesTb.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.relativesTb_KeyPress);
			// 
			// relativesLbl
			// 
			this.relativesLbl.AutoSize = true;
			this.relativesLbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
			this.relativesLbl.Location = new System.Drawing.Point(25, 34);
			this.relativesLbl.Name = "relativesLbl";
			this.relativesLbl.Size = new System.Drawing.Size(271, 40);
			this.relativesLbl.TabIndex = 0;
			this.relativesLbl.Text = "Liczba członków rodziny zgłaszanych\r\ndo ubezpieczenia zdrowotnego:";
			// 
			// descriptionBox
			// 
			this.descriptionBox.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.descriptionBox.Controls.Add(this.descriptionLabel);
			this.descriptionBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F);
			this.descriptionBox.Location = new System.Drawing.Point(30, 31);
			this.descriptionBox.Name = "descriptionBox";
			this.descriptionBox.Size = new System.Drawing.Size(400, 116);
			this.descriptionBox.TabIndex = 1;
			this.descriptionBox.TabStop = false;
			this.descriptionBox.Text = "Opis";
			// 
			// descriptionLabel
			// 
			this.descriptionLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
			this.descriptionLabel.Location = new System.Drawing.Point(25, 20);
			this.descriptionLabel.Name = "descriptionLabel";
			this.descriptionLabel.Size = new System.Drawing.Size(351, 85);
			this.descriptionLabel.TabIndex = 0;
			this.descriptionLabel.Text = "Celem programu jest wygenerowanie skoroszytu programu Excel, będącego formularzem" +
    " zgłoszenia członków rodziny do ubezpieczenia zdrowotnego.";
			this.descriptionLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(464, 411);
			this.Controls.Add(this.descriptionBox);
			this.Controls.Add(this.relativesBox);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MinimumSize = new System.Drawing.Size(480, 400);
			this.Name = "Form1";
			this.Text = "Formularz FDC";
			this.relativesBox.ResumeLayout(false);
			this.relativesBox.PerformLayout();
			this.descriptionBox.ResumeLayout(false);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.GroupBox relativesBox;
		private System.Windows.Forms.GroupBox descriptionBox;
		private System.Windows.Forms.TextBox relativesTb;
		private System.Windows.Forms.Label relativesLbl;
		private System.Windows.Forms.Label descriptionLabel;
		private System.Windows.Forms.Button relativesBtn;
	}
}

