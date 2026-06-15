using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFreelanceSalida.
	/// </summary>
	public class frmFreelanceSalida : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

    string sNombres;
		DateTime dDesde;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		DateTime dHasta;

		public frmFreelanceSalida( string SNombres , DateTime DDesde, DateTime DHasta )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
        sNombres = SNombres;
				dDesde = DDesde;
				dHasta = DHasta;
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}


		public frmFreelanceSalida()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFreelanceSalida));
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(200, 8);
			this.groupBox1.TabIndex = 7;
			this.groupBox1.TabStop = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.Location = new System.Drawing.Point(104, 88);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 6;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Location = new System.Drawing.Point(16, 88);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 5;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 16);
			this.label1.TabIndex = 759;
			this.label1.Text = "Fecha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance1;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 22);
			this.dtDesde.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(120, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 766;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// frmFreelanceSalida
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(192, 118);
			this.ControlBox = false;
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.KeyPreview = true;
			this.Name = "frmFreelanceSalida";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Salida:";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFreelanceSalida_KeyDown);
			this.Load += new System.EventHandler(this.frmFreelanceSalida_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
						this.DialogResult = DialogResult.No;
		}

		DateTime dtFecha = DateTime.Today;
		private void frmFreelanceSalida_Load(object sender, System.EventArgs e)
		{
			this.Text = "Salida:" + sNombres;
			this.dtDesde.Value = DateTime.Today;
		}

		private void frmFreelanceSalida_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if( Convert.ToDateTime( this.dtDesde.Value)  < dtFecha)
			{
				MessageBox.Show("La Fecha no puede ser menor a la Fecha Actual.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if(Convert.ToDateTime( this.dtDesde.Value) < dDesde)
			{
				MessageBox.Show("La Fecha no puede ser menor al Inicio del Periodo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if(Convert.ToDateTime(this.dtDesde.Value) > dHasta)
			{
				MessageBox.Show("La Fecha no puede ser mayor a la del Final de Periodo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
					this.DialogResult = DialogResult.OK;
		}
	}
}
