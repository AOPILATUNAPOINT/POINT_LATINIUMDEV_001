using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmDocsRefinanciacion.
	/// </summary>
	public class frmDocsRefinanciacion : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGastosDeCobranza;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAmortizaciones;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPagare;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCarta;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDocsRefinanciacion()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idCompra = 0;
		decimal dMonto = 0.00m;
		int idCliente = 0;
		DateTime dtFechaGenera = DateTime.Today;
		DateTime dtFecha = DateTime.Today;

		public frmDocsRefinanciacion(int IdCompra, decimal DMonto, int IdCliente, DateTime DTFechaGenera, DateTime DTFecha)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idCompra = IdCompra;
			dMonto = DMonto;
			idCliente = IdCliente;
			dtFechaGenera = DTFechaGenera;
			dtFecha = DTFecha;

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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDocsRefinanciacion));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.chkGastosDeCobranza = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAmortizaciones = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPagare = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkCarta = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbCobrador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			this.SuspendLayout();
			// 
			// chkGastosDeCobranza
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGastosDeCobranza.Appearance = appearance1;
			this.chkGastosDeCobranza.BackColor = System.Drawing.Color.Transparent;
			this.chkGastosDeCobranza.Checked = true;
			this.chkGastosDeCobranza.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkGastosDeCobranza.Enabled = false;
			this.chkGastosDeCobranza.Location = new System.Drawing.Point(16, 56);
			this.chkGastosDeCobranza.Name = "chkGastosDeCobranza";
			this.chkGastosDeCobranza.Size = new System.Drawing.Size(160, 22);
			this.chkGastosDeCobranza.TabIndex = 149;
			this.chkGastosDeCobranza.Text = "Gastos de cobranza";
			// 
			// chkAmortizaciones
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAmortizaciones.Appearance = appearance2;
			this.chkAmortizaciones.BackColor = System.Drawing.Color.Transparent;
			this.chkAmortizaciones.Checked = true;
			this.chkAmortizaciones.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkAmortizaciones.Enabled = false;
			this.chkAmortizaciones.Location = new System.Drawing.Point(16, 32);
			this.chkAmortizaciones.Name = "chkAmortizaciones";
			this.chkAmortizaciones.Size = new System.Drawing.Size(160, 22);
			this.chkAmortizaciones.TabIndex = 148;
			this.chkAmortizaciones.Text = "Amortizaciones";
			// 
			// chkPagare
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPagare.Appearance = appearance3;
			this.chkPagare.BackColor = System.Drawing.Color.Transparent;
			this.chkPagare.Checked = true;
			this.chkPagare.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPagare.Enabled = false;
			this.chkPagare.Location = new System.Drawing.Point(16, 8);
			this.chkPagare.Name = "chkPagare";
			this.chkPagare.Size = new System.Drawing.Size(160, 22);
			this.chkPagare.TabIndex = 147;
			this.chkPagare.Text = "Pagaré";
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(8, 266);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 419;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(96, 266);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(77, 23);
			this.btnCancelar.TabIndex = 420;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 250);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(278, 8);
			this.groupBox1.TabIndex = 421;
			this.groupBox1.TabStop = false;
			// 
			// chkCarta
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCarta.Appearance = appearance4;
			this.chkCarta.BackColor = System.Drawing.Color.Transparent;
			this.chkCarta.Checked = true;
			this.chkCarta.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCarta.Enabled = false;
			this.chkCarta.Location = new System.Drawing.Point(16, 80);
			this.chkCarta.Name = "chkCarta";
			this.chkCarta.Size = new System.Drawing.Size(160, 22);
			this.chkCarta.TabIndex = 422;
			this.chkCarta.Text = "Carta";
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
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 118);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(147, 17);
			this.label1.TabIndex = 641;
			this.label1.Text = "Responsable Documentación";
			// 
			// cmbCobrador
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador.Appearance = appearance5;
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobrador.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 335;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCobrador.DisplayMember = "Nombre";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCobrador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(8, 144);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(256, 21);
			this.cmbCobrador.TabIndex = 642;
			this.cmbCobrador.ValueMember = "idPersonal";
			this.cmbCobrador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCobrador_KeyDown);
			// 
			// txtObservaciones
			// 
			appearance6.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance6.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance6;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtObservaciones.Location = new System.Drawing.Point(8, 176);
			this.txtObservaciones.MaxLength = 1000;
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(256, 66);
			this.txtObservaciones.TabIndex = 643;
			this.txtObservaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtObservaciones_KeyDown);
			// 
			// frmDocsRefinanciacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(274, 296);
			this.ControlBox = false;
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.cmbCobrador);
			this.Controls.Add(this.chkCarta);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.chkGastosDeCobranza);
			this.Controls.Add(this.chkAmortizaciones);
			this.Controls.Add(this.chkPagare);
			this.Controls.Add(this.label1);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "frmDocsRefinanciacion";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Impresión";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDocsRefinanciacion_KeyDown);
			this.Load += new System.EventHandler(this.frmDocsRefinanciacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if (this.cmbCobrador.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Responsable de la documentación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCobrador.Focus();
				return;
			}

			if (this.chkPagare.Checked)
			{
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(idCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				ParameterField paramFieldFecha = new ParameterField ();
				ParameterDiscreteValue ValIdFecha = new ParameterDiscreteValue ();
				paramFieldFecha.ParameterFieldName = "@Monto";
				ValIdFecha.Value = dMonto;
				paramFieldFecha.CurrentValues.Add (ValIdFecha);
				paramFields.Add (paramFieldFecha);
				
				Reporte miRep = new Reporte("PagareRefinanciacion.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Págare");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.ShowDialog();
			}

			if (this.chkAmortizaciones.Checked)
			{
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfCapital = new ParameterField ();
				ParameterDiscreteValue ValCapital = new ParameterDiscreteValue ();
				pfCapital.ParameterFieldName = "@Capital";
				ValCapital.Value = dMonto;
				pfCapital.CurrentValues.Add (ValCapital);
				paramFields.Add (pfCapital);
				
				ParameterField pfCuotas = new ParameterField ();
				ParameterDiscreteValue ValCuotas = new ParameterDiscreteValue ();
				pfCuotas.ParameterFieldName = "@Cuotas";
				ValCuotas.Value = 3;
				pfCuotas.CurrentValues.Add (ValCuotas);
				paramFields.Add (pfCuotas);
				
				ParameterField pfIdCliente = new ParameterField ();
				ParameterDiscreteValue ValIdCliente = new ParameterDiscreteValue ();
				pfIdCliente.ParameterFieldName = "@IdCliente";
				ValIdCliente.Value = idCliente;
				pfIdCliente.CurrentValues.Add (ValIdCliente);
				paramFields.Add (pfIdCliente);
				
				ParameterField pfFecha = new ParameterField ();
				ParameterDiscreteValue ValFecha = new ParameterDiscreteValue ();
				pfFecha.ParameterFieldName = "@Fecha";
				ValFecha.Value = dtFechaGenera;
				pfFecha.CurrentValues.Add (ValFecha);
				paramFields.Add (pfFecha);
				
				ParameterField pfFechaGenera = new ParameterField ();
				ParameterDiscreteValue ValFechaGenera = new ParameterDiscreteValue ();
				pfFechaGenera.ParameterFieldName = "@FechaGenera";
				ValFechaGenera.Value = dtFechaGenera;
				pfFechaGenera.CurrentValues.Add (ValFechaGenera);
				paramFields.Add (pfFechaGenera);

				ParameterField pfEF = new ParameterField ();
				ParameterDiscreteValue ValEF = new ParameterDiscreteValue ();
				pfEF.ParameterFieldName = "@idEntidadFinanciera";
				ValEF.Value = 1;
				pfEF.CurrentValues.Add (ValEF);
				paramFields.Add (pfEF);
				
				Reporte miRepor = new Reporte("Informe12.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Págare");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				miRepor.ShowDialog();
			}

			if (this.chkGastosDeCobranza.Checked)
			{				
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(idCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				int bRef = 1;
				ParameterField paramFieldRef = new ParameterField ();
				ParameterDiscreteValue ValIdRef = new ParameterDiscreteValue ();
				paramFieldRef.ParameterFieldName = "@Refinanciacion";
				ValIdRef.Value = int.Parse(bRef.ToString());
				paramFieldRef.CurrentValues.Add (ValIdRef);
				paramFields.Add (paramFieldRef);

				Reporte miReporte = new Reporte("GestionDeCobranza.rpt", "");
				miReporte.MdiParent = MdiParent;
				miReporte.Titulo("Gastos de Cobranza");
				miReporte.crVista.ParameterFieldInfo = paramFields;
				miReporte.ShowDialog();
			}	

			if (this.chkCarta.Checked)
			{				
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(idCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				Reporte miReporte = new Reporte("CartaRefinanciacion.rpt", "");
				miReporte.MdiParent = MdiParent;
				miReporte.Titulo("Gastos de Cobranza");
				miReporte.crVista.ParameterFieldInfo = paramFields;
				miReporte.ShowDialog();
			}

			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CobranzaActualizaEstadoRefinanciacion {0}, {1}, {2}, 1, '{3}', '{4}'", 
				idCompra, dMonto, (int)this.cmbCobrador.Value, this.dtFecha.ToString("yyyyMMdd"), this.txtObservaciones.Text.ToString()));

			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void frmDocsRefinanciacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.Close();
		}

		private void frmDocsRefinanciacion_Load(object sender, System.EventArgs e)
		{
			this.cmbCobrador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCobranzasPersonalRefinanciacion");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void cmbCobrador_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtObservaciones.Focus();
		}

		private void txtObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnImprimir.Focus();
		}
	}
}
