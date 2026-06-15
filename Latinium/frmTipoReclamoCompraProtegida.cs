using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTipoReclamoCompraProtegida.
	/// </summary>
	public class frmTipoReclamoCompraProtegida : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnRobo;
		private System.Windows.Forms.Button btnDaño;
		private System.Windows.Forms.Button btnCerrar;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label lblBodega;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtSiniestro;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtNotificacion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmTipoReclamoCompraProtegida()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		/// <summary>
		/// Limpiar los recursos que se estén utilizando.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTipoReclamoCompraProtegida));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.btnRobo = new System.Windows.Forms.Button();
			this.btnDaño = new System.Windows.Forms.Button();
			this.btnCerrar = new System.Windows.Forms.Button();
			this.label53 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtSiniestro = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtNotificacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.dtSiniestro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtNotificacion)).BeginInit();
			this.SuspendLayout();
			// 
			// btnRobo
			// 
			this.btnRobo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRobo.Location = new System.Drawing.Point(8, 72);
			this.btnRobo.Name = "btnRobo";
			this.btnRobo.Size = new System.Drawing.Size(88, 24);
			this.btnRobo.TabIndex = 2;
			this.btnRobo.Text = "ROBO";
			this.btnRobo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRobo.Click += new System.EventHandler(this.btnRobo_Click);
			// 
			// btnDaño
			// 
			this.btnDaño.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDaño.Location = new System.Drawing.Point(104, 72);
			this.btnDaño.Name = "btnDaño";
			this.btnDaño.Size = new System.Drawing.Size(88, 24);
			this.btnDaño.TabIndex = 3;
			this.btnDaño.Text = "DAÑO";
			this.btnDaño.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDaño.Click += new System.EventHandler(this.btnDaño_Click);
			// 
			// btnCerrar
			// 
			this.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCerrar.Location = new System.Drawing.Point(200, 72);
			this.btnCerrar.Name = "btnCerrar";
			this.btnCerrar.Size = new System.Drawing.Size(88, 24);
			this.btnCerrar.TabIndex = 4;
			this.btnCerrar.Text = "CERRAR";
			this.btnCerrar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(168, 8);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(97, 16);
			this.label53.TabIndex = 980;
			this.label53.Text = "Fecha Notificacion";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(83, 16);
			this.lblBodega.TabIndex = 979;
			this.lblBodega.Text = "Fecha Siniestro";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtSiniestro
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtSiniestro.Appearance = appearance1;
			this.dtSiniestro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtSiniestro.DateButtons.Add(dateButton1);
			this.dtSiniestro.Format = "dd/MM/yyyy";
			this.dtSiniestro.Location = new System.Drawing.Point(8, 40);
			this.dtSiniestro.Name = "dtSiniestro";
			this.dtSiniestro.NonAutoSizeHeight = 23;
			this.dtSiniestro.Size = new System.Drawing.Size(112, 21);
			this.dtSiniestro.SpinButtonsVisible = true;
			this.dtSiniestro.TabIndex = 0;
			this.dtSiniestro.Value = ((object)(resources.GetObject("dtSiniestro.Value")));
			// 
			// dtNotificacion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtNotificacion.Appearance = appearance2;
			this.dtNotificacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtNotificacion.DateButtons.Add(dateButton2);
			this.dtNotificacion.Format = "dd/MM/yyyy";
			this.dtNotificacion.Location = new System.Drawing.Point(168, 40);
			this.dtNotificacion.Name = "dtNotificacion";
			this.dtNotificacion.NonAutoSizeHeight = 23;
			this.dtNotificacion.Size = new System.Drawing.Size(112, 21);
			this.dtNotificacion.SpinButtonsVisible = true;
			this.dtNotificacion.TabIndex = 1;
			this.dtNotificacion.Value = ((object)(resources.GetObject("dtNotificacion.Value")));
			// 
			// frmTipoReclamoCompraProtegida
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(298, 104);
			this.ControlBox = false;
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.dtSiniestro);
			this.Controls.Add(this.dtNotificacion);
			this.Controls.Add(this.btnCerrar);
			this.Controls.Add(this.btnDaño);
			this.Controls.Add(this.btnRobo);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmTipoReclamoCompraProtegida";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Seleccione Una Opción:";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTipoReclamoCompraProtegida_KeyDown);
			this.Load += new System.EventHandler(this.frmTipoReclamoCompraProtegida_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtSiniestro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtNotificacion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		public int iTipo = 0;

		private void frmTipoReclamoCompraProtegida_Load(object sender, System.EventArgs e)
		{
			this.dtSiniestro.Value = DateTime.Today;
			this.dtSiniestro.CalendarInfo.MaxDate = DateTime.Today;
			this.dtSiniestro.CalendarInfo.MinDate = DateTime.Parse("05/07/2019");//DateTime.Today.AddDays(-30);

			this.dtNotificacion.Value = DateTime.Today;
			this.dtNotificacion.CalendarInfo.MaxDate = DateTime.Today;
			this.dtNotificacion.CalendarInfo.MinDate = DateTime.Parse("05/07/2019");//DateTime.Today.AddDays(-30);
		}

		private void frmTipoReclamoCompraProtegida_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnCerrar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnRobo_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Esta seguro de generar la notificacion por ROBO", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;

			iTipo = 1;

			this.DialogResult = DialogResult.OK;
		}

		private void btnDaño_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Esta seguro de generar la notificacion por DAÑO", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;

			iTipo = 2;

			this.DialogResult = DialogResult.OK;
		}
	}
}
