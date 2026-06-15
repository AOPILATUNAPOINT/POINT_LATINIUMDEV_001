using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ProduccionGeneraConsumo.
	/// </summary>
	public class ProduccionGeneraConsumo : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbDesde;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private Infragistics.Win.Misc.UltraButton btReporte;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkResumen;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ProduccionGeneraConsumo()
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.cmbDesde = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.btReporte = new Infragistics.Win.Misc.UltraButton();
			this.chkResumen = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDesde
			// 
			this.cmbDesde.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(104, 24);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbDesde.TabIndex = 0;
			// 
			// cmbHasta
			// 
			this.cmbHasta.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(104, 56);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbHasta.TabIndex = 1;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(40, 24);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 14);
			this.ultraLabel1.TabIndex = 2;
			this.ultraLabel1.Text = "Desde";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(40, 56);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(34, 14);
			this.ultraLabel2.TabIndex = 3;
			this.ultraLabel2.Text = "Hasta";
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(176, 224);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.TabIndex = 4;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// btReporte
			// 
			this.btReporte.Location = new System.Drawing.Point(40, 224);
			this.btReporte.Name = "btReporte";
			this.btReporte.TabIndex = 5;
			this.btReporte.Text = "Reporte";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// chkResumen
			// 
			this.chkResumen.Location = new System.Drawing.Point(40, 96);
			this.chkResumen.Name = "chkResumen";
			this.chkResumen.TabIndex = 6;
			this.chkResumen.Text = "Resumen";
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibFacturacion";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteoF";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteo.SchemaDef = null;
			// 
			// optTipo
			// 
			this.optTipo.CheckedIndex = 0;
			this.optTipo.ItemAppearance = appearance1;
			this.optTipo.ItemOrigin = new System.Drawing.Point(6, 6);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Agrupado por Materia Prima";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Agrupado por Producto Terminado";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.ItemSpacingVertical = 6;
			this.optTipo.Location = new System.Drawing.Point(40, 128);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(208, 72);
			this.optTipo.TabIndex = 7;
			this.optTipo.Text = "Agrupado por Materia Prima";
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(40, 256);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(208, 23);
			this.ultraButton1.TabIndex = 8;
			this.ultraButton1.Text = "Actualiza costo de producto terminado";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// ProduccionGeneraConsumo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(280, 294);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.chkResumen);
			this.Controls.Add(this.btReporte);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Name = "ProduccionGeneraConsumo";
			this.Text = "Producción Genera Consumo";
			this.Load += new System.EventHandler(this.ProduccionGeneraConsumo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ProduccionGeneraConsumo_Load(object sender, System.EventArgs e)
		{
			cmbDesde.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stExec = "Exec ConsumoFacturasFecha '" + dtDesde.ToString("yyyyMMdd")
				+ "', '" + dtHasta.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsSeteo, stExec, true);
			MessageBox.Show("Consumo real del periodo Generado", "Información");
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stFiltro = "{Compra.Fecha} >= #" + dtDesde.ToString("MM-dd-yyyy")
				+ "# And {Compra.Fecha} < #" + dtHasta.AddDays(1).ToString("MM-dd-yyyy") 
				+ "# And {Compra.idTipoFactura} = 1 And Not {Compra.Borrar} "
				+ "And {Articulo.idTipoGrupo} = 6";
			Cursor = Cursors.WaitCursor;
			string stReporte = "ConsumoFactMateriaP.Rpt";
			if ((int) optTipo.Value == 0) stReporte = "ConsumoFactFecha.Rpt";
			if ((int) optTipo.Value == 2) stReporte = "ConsumoFactProd.Rpt";
			Reporte miRep = new Reporte(stReporte, stFiltro);
			miRep.MdiParent = this.MdiParent;
			if (chkResumen.Checked) miRep.Resumen();
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stExec = "Exec ProductoActualizaCosto '" + dtDesde.ToString("yyyyMMdd") 
				+ "', '" + dtHasta.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsSeteo, stExec, true);
			MessageBox.Show("Generada Actualizacion de costo de producto terminado", "Información");
		}
	}
}
