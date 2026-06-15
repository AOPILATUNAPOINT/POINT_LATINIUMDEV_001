using System;
using System.Drawing;
using System.Windows.Forms;

namespace Latinium
{
	public class frmFamilyFriendsConfirm : Form
	{
		public bool Confirmado = false;

		private Label lblTitulo;
		private Label lblMensaje;
		private Button btnSi;
		private Button btnNo;
		private Panel panelTop;

		public frmFamilyFriendsConfirm(string mensaje)
		{
			InicializarComponentes();
			lblMensaje.Text = mensaje;
		}

		private void InicializarComponentes()
		{
			// --- PROPIEDADES DEL FORM ---
			this.Text = "Confirmación";
			this.FormBorderStyle = FormBorderStyle.FixedDialog;
			this.StartPosition = FormStartPosition.CenterScreen;
			this.BackColor = Color.White;
			this.ClientSize = new Size(420, 230);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.ControlBox = false; // .NET 1.1 soportado

			// === PANEL SUPERIOR ===
			panelTop = new Panel();
			panelTop.BackColor = Color.FromArgb(30, 80, 160);
			panelTop.Location = new Point(0, 0);
			panelTop.Size = new Size(420, 40);
			this.Controls.Add(panelTop);

			// === TÍTULO ===
			lblTitulo = new Label();
			lblTitulo.Text = "Cliente Family & Friends";
			lblTitulo.ForeColor = Color.White;
			lblTitulo.Font = new Font("Microsoft Sans Serif", 12, FontStyle.Bold);
			lblTitulo.TextAlign = ContentAlignment.MiddleCenter;
			lblTitulo.Location = new Point(0, 10);
			lblTitulo.Size = new Size(420, 20);
			panelTop.Controls.Add(lblTitulo);

			// === MENSAJE ===
			lblMensaje = new Label();
			lblMensaje.Font = new Font("Microsoft Sans Serif", 10);
			lblMensaje.TextAlign = ContentAlignment.MiddleCenter;
			lblMensaje.Location = new Point(20, 60);
			lblMensaje.Size = new Size(380, 80);
			this.Controls.Add(lblMensaje);

			// === BOTÓN SÍ ===
			btnSi = new Button();
			btnSi.Text = "Sí, continuar";
			btnSi.Location = new Point(60, 160);
			btnSi.Size = new Size(130, 35);
			btnSi.Click += new EventHandler(BtnSi_Click);
			this.Controls.Add(btnSi);

			// === BOTÓN NO ===
			btnNo = new Button();
			btnNo.Text = "No";
			btnNo.Location = new Point(230, 160);
			btnNo.Size = new Size(130, 35);
			btnNo.Click += new EventHandler(BtnNo_Click);
			this.Controls.Add(btnNo);
		}

		private void BtnSi_Click(object sender, EventArgs e)
		{
			Confirmado = true;
			this.DialogResult = DialogResult.OK;
			this.Close();
		}

		private void BtnNo_Click(object sender, EventArgs e)
		{
			Confirmado = false;
			this.DialogResult = DialogResult.Cancel;
			this.Close();
		}
	}
}
