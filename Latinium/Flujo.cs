using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Flujo.
	/// </summary>
	public class Flujo : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkResumen;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.Misc.UltraButton btManual;
		private Infragistics.Win.Misc.UltraButton btCuentas;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Flujo()
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
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.chkResumen = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btManual = new Infragistics.Win.Misc.UltraButton();
			this.btCuentas = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDesde
			// 
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(64, 24);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(104, 21);
			this.cmbDesde.TabIndex = 0;
			this.cmbDesde.Value = new System.DateTime(2009, 8, 13, 0, 0, 0, 0);
			// 
			// cmbHasta
			// 
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(64, 64);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(104, 21);
			this.cmbHasta.TabIndex = 1;
			this.cmbHasta.Value = new System.DateTime(2009, 8, 13, 0, 0, 0, 0);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(16, 24);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 14);
			this.ultraLabel1.TabIndex = 2;
			this.ultraLabel1.Text = "Desde";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(16, 64);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(34, 14);
			this.ultraLabel2.TabIndex = 3;
			this.ultraLabel2.Text = "Hasta";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(64, 160);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(64, 24);
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Reporte";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(216, 160);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(64, 24);
			this.btCancelar.TabIndex = 5;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// chkResumen
			// 
			this.chkResumen.Location = new System.Drawing.Point(64, 112);
			this.chkResumen.Name = "chkResumen";
			this.chkResumen.TabIndex = 6;
			this.chkResumen.Text = "Resumen";
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// btManual
			// 
			this.btManual.Location = new System.Drawing.Point(248, 24);
			this.btManual.Name = "btManual";
			this.btManual.Size = new System.Drawing.Size(64, 24);
			this.btManual.TabIndex = 7;
			this.btManual.Text = "&Manual";
			this.btManual.Click += new System.EventHandler(this.btManual_Click);
			// 
			// btCuentas
			// 
			this.btCuentas.Location = new System.Drawing.Point(248, 64);
			this.btCuentas.Name = "btCuentas";
			this.btCuentas.Size = new System.Drawing.Size(64, 24);
			this.btCuentas.TabIndex = 8;
			this.btCuentas.Text = "&Cuentas";
			this.btCuentas.Click += new System.EventHandler(this.btCuentas_Click);
			// 
			// Flujo
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(352, 205);
			this.Controls.Add(this.btCuentas);
			this.Controls.Add(this.btManual);
			this.Controls.Add(this.chkResumen);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Name = "Flujo";
			this.Text = "Flujo";
			this.Load += new System.EventHandler(this.Flujo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stTitulo = "Flujo de Caja desde " + dtDesde.ToString("dd/MMM/yyyy")
				+ " hasta " + dtHasta.ToString("dd/MMM/yyyy");

			#region parametros
			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FechaDesde";
			discreteVal1.Value = dtDesde.ToString("yyyy-MM-dd HH:mm:ss");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);
		
			ParameterField paramField2 = new ParameterField ();
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "@FechaHasta";
			discreteVal2.Value = dtHasta.ToString("yyyy-MM-dd HH:mm:ss");
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);
			#endregion parametros

			Reporte miRep = new Reporte("FlujoCaja.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.crVista.ParameterFieldInfo = paramFields;
			if (chkResumen.Checked) miRep.Resumen();
			miRep.Titulo(stTitulo);
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void btManual_Click(object sender, System.EventArgs e)
		{
			FlujoManual miFlujo = new FlujoManual();
			miFlujo.MdiParent = this.MdiParent;
			miFlujo.Show();
		}

		private void btCuentas_Click(object sender, System.EventArgs e)
		{
			FlujoCuenta miFlujo = new FlujoCuenta();
			miFlujo.MdiParent = this.MdiParent;
			miFlujo.Show();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Flujo_Load(object sender, System.EventArgs e)
		{
			cmbDesde.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
		}
	}
}
