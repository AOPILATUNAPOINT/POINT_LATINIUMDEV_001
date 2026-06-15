using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ReservaReporte.
	/// </summary>
	public class ReservaReporte : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFecha;
		private Infragistics.Win.Misc.UltraButton btUtilidad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkReserva;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReserva;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtUsuario;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkUsuario;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ReservaReporte()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.optFecha = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btUtilidad = new Infragistics.Win.Misc.UltraButton();
			this.chkReserva = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtReserva = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.chkUsuario = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtUsuario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDesde
			// 
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(128, 32);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 8;
			this.cmbDesde.Size = new System.Drawing.Size(104, 21);
			this.cmbDesde.TabIndex = 0;
			this.cmbDesde.Value = new System.DateTime(2006, 1, 2, 0, 0, 0, 0);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(56, 32);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 14);
			this.ultraLabel1.TabIndex = 1;
			this.ultraLabel1.Text = "Desde";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(56, 64);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(34, 14);
			this.ultraLabel2.TabIndex = 3;
			this.ultraLabel2.Text = "Hasta";
			// 
			// cmbHasta
			// 
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(128, 64);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 8;
			this.cmbHasta.Size = new System.Drawing.Size(104, 21);
			this.cmbHasta.TabIndex = 2;
			this.cmbHasta.Value = new System.DateTime(2006, 1, 2, 0, 0, 0, 0);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(280, 32);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(280, 104);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// optFecha
			// 
			this.optFecha.CheckedIndex = 0;
			this.optFecha.ItemAppearance = appearance1;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Fecha";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Fecha Salida";
			this.optFecha.Items.Add(valueListItem1);
			this.optFecha.Items.Add(valueListItem2);
			this.optFecha.Location = new System.Drawing.Point(128, 96);
			this.optFecha.Name = "optFecha";
			this.optFecha.Size = new System.Drawing.Size(96, 40);
			this.optFecha.TabIndex = 6;
			this.optFecha.Text = "Fecha";
			// 
			// btUtilidad
			// 
			this.btUtilidad.Location = new System.Drawing.Point(280, 152);
			this.btUtilidad.Name = "btUtilidad";
			this.btUtilidad.TabIndex = 7;
			this.btUtilidad.Text = "Utilidad";
			this.btUtilidad.Click += new System.EventHandler(this.btUtilidad_Click);
			// 
			// chkReserva
			// 
			this.chkReserva.Checked = true;
			this.chkReserva.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkReserva.Location = new System.Drawing.Point(24, 152);
			this.chkReserva.Name = "chkReserva";
			this.chkReserva.Size = new System.Drawing.Size(96, 20);
			this.chkReserva.TabIndex = 8;
			this.chkReserva.Text = "Toda Reserva";
			this.chkReserva.CheckedChanged += new System.EventHandler(this.chkReserva_CheckedChanged);
			// 
			// txtReserva
			// 
			this.txtReserva.Enabled = false;
			this.txtReserva.Location = new System.Drawing.Point(128, 152);
			this.txtReserva.Name = "txtReserva";
			this.txtReserva.Size = new System.Drawing.Size(112, 21);
			this.txtReserva.TabIndex = 9;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// chkUsuario
			// 
			this.chkUsuario.Checked = true;
			this.chkUsuario.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkUsuario.Location = new System.Drawing.Point(24, 184);
			this.chkUsuario.Name = "chkUsuario";
			this.chkUsuario.Size = new System.Drawing.Size(96, 20);
			this.chkUsuario.TabIndex = 10;
			this.chkUsuario.Text = "Todo Usuario";
			this.chkUsuario.CheckedChanged += new System.EventHandler(this.chkUsuario_CheckedChanged);
			// 
			// txtUsuario
			// 
			this.txtUsuario.Enabled = false;
			this.txtUsuario.Location = new System.Drawing.Point(128, 184);
			this.txtUsuario.Name = "txtUsuario";
			this.txtUsuario.Size = new System.Drawing.Size(112, 21);
			this.txtUsuario.TabIndex = 11;
			// 
			// ReservaReporte
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(392, 246);
			this.Controls.Add(this.txtUsuario);
			this.Controls.Add(this.chkUsuario);
			this.Controls.Add(this.txtReserva);
			this.Controls.Add(this.chkReserva);
			this.Controls.Add(this.btUtilidad);
			this.Controls.Add(this.optFecha);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbDesde);
			this.MaximizeBox = false;
			this.Name = "ReservaReporte";
			this.Text = "Reserva Reporte";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ReservaReporte_Closing);
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stTitulo = "Desde " + dtDesde.ToString("dd/MMM/yyyy") + 
				" Hasta " + dtHasta.ToString("dd/MMM/yyyy");
			dtHasta = dtHasta.AddDays(1);
			string stFiltro = "{Reservacion.Fecha} >= #" + dtDesde.ToString("MM/dd/yyyy")
				+ "# And {Reservacion.Fecha} < #" + dtHasta.ToString("MM/dd/yyyy") + "#";
			if ((int) optFecha.Value == 1)
				stFiltro = "{Reservacion.FechaSalida} >= #" + dtDesde.ToString("MM/dd/yyyy")
					+ "# And {Reservacion.FechaSalida} < #" + dtHasta.ToString("MM/dd/yyyy") + "#";
      
				stFiltro += " And {Reservacion.Borrar} = False ";
			Reporte miReporte = new Reporte("AgenciaViaje.Rpt", stFiltro);
			miReporte.Titulo(stTitulo);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btUtilidad_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(txtReserva, "");
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stTitulo = "Desde " + dtDesde.ToString("dd/MMM/yyyy") + 
				" Hasta " + dtHasta.ToString("dd/MMM/yyyy");
			dtHasta = dtHasta.AddDays(1);
			string stFiltro = "{Compra.Fecha} >= #" + dtDesde.ToString("MM/dd/yyyy")
				+ "# And {Compra.Fecha} < #" + dtHasta.ToString("MM/dd/yyyy") + "#";
			stFiltro += " And Length(Trim({Compra.Otro})) > 0 ";
				stFiltro += " And {Reservacion.Borrar} = False ";
			stTitulo = "Reporte de Utilidad desde " + stTitulo;

			if (!chkReserva.Checked)
			{
				if (txtReserva.Text.Trim().Length == 0)
				{
					errorProvider.SetError(txtReserva, "Ingrese Reserva");
					return;
				}
				stFiltro = "Trim({Compra.Otro}) = '" + txtReserva.Text +"'";
				stTitulo = "Reporte de Utilidad de Reserva " + txtReserva.Text;
					stFiltro += " And {Reservacion.Borrar} = False ";
			}

			if (!chkUsuario.Checked)
			{
				if (txtUsuario.Text.Trim().Length == 0)
				{
					errorProvider.SetError(txtReserva, "Ingrese Usuario");
					return;
				}
				stFiltro += " And Trim({Reservacion.Usuario}) = '" + txtUsuario.Text +"'";
				stFiltro += " And {Reservacion.Borrar} = False ";
				stTitulo += " Usuario " + txtUsuario.Text;
			}
			
			Reporte miRep = new Reporte("UtilidadAgencia.rpt", stFiltro);
			miRep.Titulo(stTitulo);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
		}

		private void chkReserva_CheckedChanged(object sender, System.EventArgs e)
		{
			txtReserva.Enabled = !chkReserva.Checked;
			if (!chkReserva.Checked) txtReserva.Select();
		}

		private void chkUsuario_CheckedChanged(object sender, System.EventArgs e)
		{
			txtUsuario.Enabled = !chkUsuario.Checked;
			if (!chkUsuario.Checked) txtUsuario.Select();		
		}

		private void ReservaReporte_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
		}
	}
}
