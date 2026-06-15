using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolUtilidad.
	/// </summary>
	public class RolUtilidad : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUtilidad;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btReporte;
		private C1.Data.C1DataSet cdsSeteoR;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		CultureInfo us = new CultureInfo("en-US");

		public RolUtilidad()
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
			this.txtUtilidad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btReporte = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsSeteoR = new C1.Data.C1DataSet();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoR)).BeginInit();
			this.SuspendLayout();
			// 
			// txtUtilidad
			// 
			this.txtUtilidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUtilidad.FormatString = "#,##0.00";
			this.txtUtilidad.Location = new System.Drawing.Point(112, 52);
			this.txtUtilidad.Name = "txtUtilidad";
			this.txtUtilidad.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtUtilidad.PromptChar = ' ';
			this.txtUtilidad.Size = new System.Drawing.Size(100, 22);
			this.txtUtilidad.TabIndex = 0;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(48, 52);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(42, 15);
			this.ultraLabel1.TabIndex = 1;
			this.ultraLabel1.Text = "Utilidad";
			// 
			// cmbFecha
			// 
			this.cmbFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(112, 86);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(96, 21);
			this.cmbFecha.TabIndex = 2;
			this.cmbFecha.Value = new System.DateTime(2008, 3, 23, 0, 0, 0, 0);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(48, 86);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(34, 15);
			this.ultraLabel2.TabIndex = 3;
			this.ultraLabel2.Text = "Fecha";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(40, 146);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 4;
			this.btAceptar.Text = "&Generar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btReporte
			// 
			this.btReporte.Location = new System.Drawing.Point(120, 146);
			this.btReporte.Name = "btReporte";
			this.btReporte.Size = new System.Drawing.Size(75, 25);
			this.btReporte.TabIndex = 5;
			this.btReporte.Text = "&Reporte";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdsSeteoR
			// 
			this.cdsSeteoR.BindingContextCtrl = this;
			this.cdsSeteoR.DataLibrary = "LibPersonal";
			this.cdsSeteoR.DataLibraryUrl = "";
			this.cdsSeteoR.DataSetDef = "dsSeteoRol";
			this.cdsSeteoR.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoR.Name = "cdsSeteoR";
			this.cdsSeteoR.SchemaClassName = "LibPersonal.DataClass";
			this.cdsSeteoR.SchemaDef = null;
			this.cdsSeteoR.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(200, 146);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(75, 25);
			this.ultraButton1.TabIndex = 6;
			this.ultraButton1.Text = "&Resumen";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// RolUtilidad
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(312, 213);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btReporte);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtUtilidad);
			this.MaximizeBox = false;
			this.Name = "RolUtilidad";
			this.Text = "Reparticion de Utilidades";
			this.Load += new System.EventHandler(this.RolUtilidad_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtUtilidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoR)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoR.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stFiltro = "{RolPago.Utilidad} > 0 And {RolPago.Anio} = " + dtFecha.ToString("yyyy") ;
			Reporte miRep = new Reporte("RolUtilidad.rpt", stFiltro);
			miRep.MdiParent = MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			double dUtilidad = double.Parse(txtUtilidad.Value.ToString());
			errorProvider.SetError(txtUtilidad, "");
			if (dUtilidad == 0)
			{
				errorProvider.SetError(txtUtilidad, "Ingrese Utilidad");
				return;
			}
			Cursor = Cursors.WaitCursor;
			string stExec = "Exec RolUtilidad " + dtFecha.ToString("yyyy") 
				+ ", " + dUtilidad.ToString("0.00", us);
			Funcion.EjecutaSQL(cdsSeteoR, stExec, true);
			Cursor = Cursors.Default;
			MessageBox.Show("Reparticion de Utilidades Terminada", "Información");
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stFiltro = "{RolPago.Utilidad} > 0 And {RolPago.Anio} = " + dtFecha.ToString("yyyy") ;
			Reporte miRep = new Reporte("RolUtilidadResumen.rpt", stFiltro);
			miRep.MdiParent = MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;

		}

		private void RolUtilidad_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
