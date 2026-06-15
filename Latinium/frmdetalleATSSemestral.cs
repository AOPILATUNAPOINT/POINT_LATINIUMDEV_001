using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Xml.Schema;
using System.Globalization;
using C1.Data;
using System.IO;
using System.Text;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmdetalleATSSemestral.
	/// </summary>
	public class frmdetalleATSSemestral : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTransacciones;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbComprobantes;
		private System.Windows.Forms.Button btnGenerar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdDebitoRenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCierreImpuestos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidAsientoRenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidAsientoIVA;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidDebitoIVA;
		private System.Windows.Forms.Button btnCerrarImpuestos;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnTalonResumen;
		private System.Windows.Forms.Button btnGeneraXML;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridTransacciones;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.DateTimePicker dtPeriodoDesde;
		private System.Windows.Forms.DateTimePicker dtPeriodoHasta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmdetalleATSSemestral()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComprobante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobante");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fiscal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base0");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseIVA");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetenidoIVA");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetenidoRenta");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Base0", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Base0", 8, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BaseIVA", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BaseIVA", 9, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Iva", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Iva", 10, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 11, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetenidoIVA", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetenidoIVA", 12, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetenidoRenta", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetenidoRenta", 13, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fiscal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BaseIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Iva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetenidoIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetenidoRenta");
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.dtPeriodoDesde = new System.Windows.Forms.DateTimePicker();
			this.label3 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbTransacciones = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbComprobantes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdDebitoRenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidCierreImpuestos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidAsientoRenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidAsientoIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidDebitoIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnCerrarImpuestos = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnTalonResumen = new System.Windows.Forms.Button();
			this.btnGeneraXML = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridTransacciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.dtPeriodoHasta = new System.Windows.Forms.DateTimePicker();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.saveFileDialog2 = new System.Windows.Forms.SaveFileDialog();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransacciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobantes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDebitoRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCierreImpuestos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsientoRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsientoIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDebitoIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTransacciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.SuspendLayout();
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
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(424, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 662;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(37, 16);
			this.label4.TabIndex = 661;
			this.label4.Text = "Desde";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtPeriodoDesde
			// 
			this.dtPeriodoDesde.Checked = false;
			this.dtPeriodoDesde.CustomFormat = "dd/MM/yyyy";
			this.dtPeriodoDesde.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtPeriodoDesde.Location = new System.Drawing.Point(96, 8);
			this.dtPeriodoDesde.Name = "dtPeriodoDesde";
			this.dtPeriodoDesde.ShowUpDown = true;
			this.dtPeriodoDesde.Size = new System.Drawing.Size(88, 20);
			this.dtPeriodoDesde.TabIndex = 660;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(66, 16);
			this.label3.TabIndex = 659;
			this.label3.Text = "Transaccion";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(8, 72);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(73, 16);
			this.label9.TabIndex = 657;
			this.label9.Text = "Comprobante";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTransacciones
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTransacciones.Appearance = appearance1;
			this.cmbTransacciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTransacciones.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 304;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTransacciones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTransacciones.DisplayMember = "Nombre";
			this.cmbTransacciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTransacciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTransacciones.Location = new System.Drawing.Point(96, 40);
			this.cmbTransacciones.Name = "cmbTransacciones";
			this.cmbTransacciones.Size = new System.Drawing.Size(304, 21);
			this.cmbTransacciones.TabIndex = 658;
			this.cmbTransacciones.ValueMember = "idTipoFactura";
			// 
			// cmbComprobantes
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbComprobantes.Appearance = appearance2;
			this.cmbComprobantes.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbComprobantes.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 304;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbComprobantes.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbComprobantes.DisplayMember = "Comprobante";
			this.cmbComprobantes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbComprobantes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbComprobantes.Location = new System.Drawing.Point(96, 72);
			this.cmbComprobantes.Name = "cmbComprobantes";
			this.cmbComprobantes.Size = new System.Drawing.Size(304, 21);
			this.cmbComprobantes.TabIndex = 656;
			this.cmbComprobantes.ValueMember = "idComprobante";
			// 
			// btnGenerar
			// 
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(424, 72);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.TabIndex = 655;
			this.btnGenerar.Text = "&Ver";
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(648, 8);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 674;
			this.txtEstado.Visible = false;
			// 
			// txtIdDebitoRenta
			// 
			this.txtIdDebitoRenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdDebitoRenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdDebitoRenta.Enabled = false;
			this.txtIdDebitoRenta.Location = new System.Drawing.Point(632, 8);
			this.txtIdDebitoRenta.Name = "txtIdDebitoRenta";
			this.txtIdDebitoRenta.PromptChar = ' ';
			this.txtIdDebitoRenta.Size = new System.Drawing.Size(16, 21);
			this.txtIdDebitoRenta.TabIndex = 673;
			this.txtIdDebitoRenta.Visible = false;
			// 
			// txtidCierreImpuestos
			// 
			this.txtidCierreImpuestos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCierreImpuestos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCierreImpuestos.Enabled = false;
			this.txtidCierreImpuestos.Location = new System.Drawing.Point(568, 8);
			this.txtidCierreImpuestos.Name = "txtidCierreImpuestos";
			this.txtidCierreImpuestos.PromptChar = ' ';
			this.txtidCierreImpuestos.Size = new System.Drawing.Size(16, 21);
			this.txtidCierreImpuestos.TabIndex = 672;
			this.txtidCierreImpuestos.Visible = false;
			// 
			// txtidAsientoRenta
			// 
			this.txtidAsientoRenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidAsientoRenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidAsientoRenta.Enabled = false;
			this.txtidAsientoRenta.Location = new System.Drawing.Point(600, 8);
			this.txtidAsientoRenta.Name = "txtidAsientoRenta";
			this.txtidAsientoRenta.PromptChar = ' ';
			this.txtidAsientoRenta.Size = new System.Drawing.Size(16, 21);
			this.txtidAsientoRenta.TabIndex = 671;
			this.txtidAsientoRenta.Visible = false;
			// 
			// txtidAsientoIVA
			// 
			this.txtidAsientoIVA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidAsientoIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidAsientoIVA.Enabled = false;
			this.txtidAsientoIVA.Location = new System.Drawing.Point(584, 8);
			this.txtidAsientoIVA.Name = "txtidAsientoIVA";
			this.txtidAsientoIVA.PromptChar = ' ';
			this.txtidAsientoIVA.Size = new System.Drawing.Size(16, 21);
			this.txtidAsientoIVA.TabIndex = 670;
			this.txtidAsientoIVA.Visible = false;
			// 
			// txtidDebitoIVA
			// 
			this.txtidDebitoIVA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidDebitoIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidDebitoIVA.Enabled = false;
			this.txtidDebitoIVA.Location = new System.Drawing.Point(616, 8);
			this.txtidDebitoIVA.Name = "txtidDebitoIVA";
			this.txtidDebitoIVA.PromptChar = ' ';
			this.txtidDebitoIVA.Size = new System.Drawing.Size(16, 21);
			this.txtidDebitoIVA.TabIndex = 669;
			this.txtidDebitoIVA.Visible = false;
			// 
			// btnCerrarImpuestos
			// 
			this.btnCerrarImpuestos.CausesValidation = false;
			this.btnCerrarImpuestos.Enabled = false;
			this.btnCerrarImpuestos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCerrarImpuestos.Location = new System.Drawing.Point(728, 8);
			this.btnCerrarImpuestos.Name = "btnCerrarImpuestos";
			this.btnCerrarImpuestos.Size = new System.Drawing.Size(120, 23);
			this.btnCerrarImpuestos.TabIndex = 668;
			this.btnCerrarImpuestos.Text = "&Cerrar Impuestos";
			this.btnCerrarImpuestos.Click += new System.EventHandler(this.btnCerrarImpuestos_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(544, 72);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 667;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnExportar
			// 
			this.btnExportar.CausesValidation = false;
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(864, 72);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(120, 23);
			this.btnExportar.TabIndex = 666;
			this.btnExportar.Text = "&Exportar Excel";
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnTalonResumen
			// 
			this.btnTalonResumen.CausesValidation = false;
			this.btnTalonResumen.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTalonResumen.Location = new System.Drawing.Point(864, 8);
			this.btnTalonResumen.Name = "btnTalonResumen";
			this.btnTalonResumen.Size = new System.Drawing.Size(120, 23);
			this.btnTalonResumen.TabIndex = 665;
			this.btnTalonResumen.Text = "&Ver Talón Resumen";
			this.btnTalonResumen.Click += new System.EventHandler(this.btnTalonResumen_Click);
			// 
			// btnGeneraXML
			// 
			this.btnGeneraXML.CausesValidation = false;
			this.btnGeneraXML.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGeneraXML.Location = new System.Drawing.Point(864, 40);
			this.btnGeneraXML.Name = "btnGeneraXML";
			this.btnGeneraXML.Size = new System.Drawing.Size(120, 23);
			this.btnGeneraXML.TabIndex = 664;
			this.btnGeneraXML.Text = "&Generar XML";
			this.btnGeneraXML.Click += new System.EventHandler(this.btnGeneraXML_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(8, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(992, 8);
			this.groupBox1.TabIndex = 675;
			this.groupBox1.TabStop = false;
			// 
			// uGridTransacciones
			// 
			this.uGridTransacciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridTransacciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTransacciones.DataSource = this.ultraDataSource4;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridTransacciones.DisplayLayout.Appearance = appearance3;
			this.uGridTransacciones.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "Serie";
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 96;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 112;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 110;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 110;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 90;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 235;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Width = 65;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Width = 85;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance4;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Base 0";
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Width = 80;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance5;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Base I.V.A.";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance6;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "I.V.A.";
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn15.Width = 80;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance7;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance8;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Ret IVA";
			ultraGridColumn17.Header.VisiblePosition = 12;
			ultraGridColumn17.Width = 80;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance9;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Ret Renta";
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.Width = 80;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance10;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance11;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance12;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance13;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance14;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance15;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridTransacciones.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTransacciones.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.uGridTransacciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.uGridTransacciones.DisplayLayout.Override.CardAreaAppearance = appearance17;
			this.uGridTransacciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8.5F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTransacciones.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.uGridTransacciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransacciones.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTransacciones.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.uGridTransacciones.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridTransacciones.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTransacciones.Location = new System.Drawing.Point(8, 120);
			this.uGridTransacciones.Name = "uGridTransacciones";
			this.uGridTransacciones.Size = new System.Drawing.Size(976, 288);
			this.uGridTransacciones.TabIndex = 676;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(bool);
			ultraDataColumn2.DefaultValue = false;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn7.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn9.DataType = typeof(System.Double);
			ultraDataColumn10.DataType = typeof(System.Double);
			ultraDataColumn11.DataType = typeof(System.Double);
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn13.DataType = typeof(System.Double);
			ultraDataColumn14.DataType = typeof(System.Double);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(200, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 678;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtPeriodoHasta
			// 
			this.dtPeriodoHasta.Checked = false;
			this.dtPeriodoHasta.CustomFormat = "dd/MM/yyyy";
			this.dtPeriodoHasta.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtPeriodoHasta.Location = new System.Drawing.Point(240, 8);
			this.dtPeriodoHasta.Name = "dtPeriodoHasta";
			this.dtPeriodoHasta.ShowUpDown = true;
			this.dtPeriodoHasta.Size = new System.Drawing.Size(88, 20);
			this.dtPeriodoHasta.TabIndex = 677;
			this.dtPeriodoHasta.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
			// 
			// frmdetalleATSSemestral
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(992, 414);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtPeriodoHasta);
			this.Controls.Add(this.uGridTransacciones);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdDebitoRenta);
			this.Controls.Add(this.txtidCierreImpuestos);
			this.Controls.Add(this.txtidAsientoRenta);
			this.Controls.Add(this.txtidAsientoIVA);
			this.Controls.Add(this.txtidDebitoIVA);
			this.Controls.Add(this.btnCerrarImpuestos);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnTalonResumen);
			this.Controls.Add(this.btnGeneraXML);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtPeriodoDesde);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.cmbTransacciones);
			this.Controls.Add(this.cmbComprobantes);
			this.Controls.Add(this.btnGenerar);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmdetalleATSSemestral";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Anexo transaccional Simplificado Semestral";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmdetalleATSSemestral_KeyDown);
			this.Load += new System.EventHandler(this.frmdetalleATSSemestral_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransacciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobantes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDebitoRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCierreImpuestos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsientoRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsientoIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDebitoIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTransacciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		CultureInfo us = new CultureInfo("en-US");
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private string NombreGenerado(string stArchivo)
		{
//			DateTime Fecha = (DateTime)this.dtPeriodo.Value;
//			int UDia = DateTime.DaysInMonth(Fecha.Year, Fecha.Month);
//			int iMes = Fecha.Month;
//			int iAño = Fecha.Year;			
//			DateTime dtFecha = Convert.ToDateTime(UDia.ToString() + "/" + iMes.ToString() + "/" + iAño.ToString());

			DateTime dtFechaD = (DateTime)this.dtPeriodoDesde.Value;
			int UDia = DateTime.DaysInMonth(dtFechaD.Year, dtFechaD.Month);
			int iMes = dtFechaD.Month;
			int iAño = dtFechaD.Year;

			DateTime dtFechaH = (DateTime)this.dtPeriodoHasta.Value;
			int VDia = DateTime.DaysInMonth(dtFechaH.Year, dtFechaH.Month);
			int jMes = dtFechaH.Month;
			int jAño = dtFechaH.Year;

			DateTime dtFechaDesde = Convert.ToDateTime("01/" + iMes.ToString() + "/" + iAño.ToString());
			DateTime dtFechaHasta = Convert.ToDateTime(VDia.ToString() + "/" + jMes.ToString() + "/" + jAño.ToString());
			
			stArchivo += dtFechaD.ToString("MMyyyy") + dtFechaH.ToString("MMyyyy") + ".XML";
			if (Directory.Exists(@"C:\XML_ANEXO_TRANSACCIONAL\")) stArchivo = @"C:\XML_ANEXO_TRANSACCIONAL\" + stArchivo;
			return stArchivo;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmdetalleATSSemestral_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "040704");

			if (!Funcion.Permiso("1519", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Anexos Transaccionales Semestrales", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			
			this.dtPeriodoDesde.Value = DateTime.Today; //MenuLatinium.dtFechaSistema;
			this.dtPeriodoHasta.Value = DateTime.Today; //MenuLatinium.dtFechaSistema;
			
			this.cmbTransacciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoFactura, Nombre From CompraNumero Where idTipoFactura In (1, 4, 5, 6) Order By Nombre");
			this.cmbComprobantes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idComprobante, Comprobante From CompraComprobante Where idComprobante In (1, 2, 3) Order By Codigo");
		}

		private void dateTimePicker1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			if (this.cmbTransacciones.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Transaccion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{
//				DateTime dtFechaDesde = (DateTime)this.dtPeriodoDesde.Value;
//			  int UDia = DateTime.DaysInMonth(dtFechaDesde.Year, dtFechaDesde.Month);
//				int iMes = dtFechaDesde.Month;
//				int iAño = dtFechaDesde.Year;
//
//				DateTime dtFechaHasta = (DateTime)this.dtPeriodoHasta.Value;
//				int VDia = DateTime.DaysInMonth(dtFechaHasta.Year, dtFechaHasta.Month);
//				int jMes = dtFechaHasta.Month;
//				int jAño = dtFechaHasta.Year;

//				DateTime Desde = Convert.ToDateTime("01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
//				DateTime Hasta = Convert.ToDateTime(VDia.ToString() + "/" + jMes.ToString() + "/" + jAño.ToString());

				int idComprobante = 0;

				if (this.cmbComprobantes.ActiveRow != null) idComprobante = (int)this.cmbComprobantes.Value;

				string sSQL = string.Format("Exec ConsultaTransacciones '{0}', '{1}', {2}, {3}", Convert.ToDateTime(this.dtPeriodoDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtPeriodoHasta.Value).ToString("yyyyMMdd"), (int)this.cmbTransacciones.Value, idComprobante);
				this.uGridTransacciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				this.lblContador.Text = this.uGridTransacciones.Rows.Count + " REGISTROS";
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridTransacciones);
		}

		private void btnCerrarImpuestos_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnTalonResumen_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaD = (DateTime)this.dtPeriodoDesde.Value;
			int UDia = DateTime.DaysInMonth(dtFechaD.Year, dtFechaD.Month);
			int iMes = dtFechaD.Month;
			int iAño = dtFechaD.Year;

			DateTime dtFechaH = (DateTime)this.dtPeriodoHasta.Value;
			int VDia = DateTime.DaysInMonth(dtFechaH.Year, dtFechaH.Month);
			int jMes = dtFechaH.Month;
			int jAño = dtFechaH.Year;

			DateTime dtFechaDesde = Convert.ToDateTime("01/" + iMes.ToString() + "/" + iAño.ToString());
			DateTime dtFechaHasta = Convert.ToDateTime(VDia.ToString() + "/" + jMes.ToString() + "/" + jAño.ToString());

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@Desde";
			discreteVal.Value = dtFechaDesde;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField0 = new ParameterField ();
			ParameterDiscreteValue discreteVal0 = new ParameterDiscreteValue ();
			paramField0.ParameterFieldName = "@Hasta";
			discreteVal0.Value = dtFechaHasta;
			paramField0.CurrentValues.Add (discreteVal0);
			paramFields.Add (paramField0);

			Reporte miReporte = new Reporte("TalonResumenSemestral.rpt");
			miReporte.MdiParent = MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Titulo("Talón Resumen Anexo Transaccional Simplificado");
			miReporte.Show();
		}

		private void btnGeneraXML_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;

			if (DialogResult.Yes == MessageBox.Show("¿Desea Generar el Anexo Transaccional Simplificado?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
			{
				Funcion.TiempoInicia();

				#region Variables
				DateTime dtFechaD = (DateTime)this.dtPeriodoDesde.Value;
				int UDia = DateTime.DaysInMonth(dtFechaD.Year, dtFechaD.Month);
				int iMes = dtFechaD.Month;
				int iAño = dtFechaD.Year;

				DateTime dtFechaH = (DateTime)this.dtPeriodoHasta.Value;
				int VDia = DateTime.DaysInMonth(dtFechaH.Year, dtFechaH.Month);
				int jMes = dtFechaH.Month;
				int jAño = dtFechaH.Year;

				DateTime dtFecha = Convert.ToDateTime("01/" + iMes.ToString() + "/" + iAño.ToString());
				DateTime dtHasta = Convert.ToDateTime(VDia.ToString() + "/" + jMes.ToString() + "/" + jAño.ToString());

				string sTipoId = "";
				string sRUCEmpresa = "";
				string sEmpresa = "";
				string regEmpresa = "";
				char Pad = '0';
				int idCompra = 0;
				int idOrigen = 0;
				int idTipoFactura = 0;
				DateTime dtFechaRegistro = DateTime.Today;
				DateTime dtFechaEmision = DateTime.Today;
				decimal dBaseNGrvIva = 0.00m;
				decimal dBase0 = 0.00m;
				decimal dBase = 0.00m;
				decimal dICE = 0.00m;
				decimal dIVA = 0.00m;
				decimal dIVABienes10 = 0.00m;
				decimal dIVAServicios20 = 0.00m;
				decimal dIVABienes = 0.00m;
				decimal dIVAServicios = 0.00m;
				decimal dIVABienesServicios = 0.00m;
				decimal dvalorRetencionNc = 0.00m;

				decimal dVentaTotal = 0.00m;
				decimal dVentaEstablecimiento = 0.00m;
				decimal dVentaBaseNGIVA = 0.00m;
				decimal dVentaBase0 = 0.00m;
				decimal dVentaBase = 0.00m;
				decimal dVentaIVA = 0.00m;
				decimal dVentaRetIVA = 0.00m;
				decimal dVentaRetRenta = 0.00m;

				DateTime dtFechaRet = DateTime.Today;

				DataSet oDSCompras = new DataSet();
				DataSet oDSVentas = new DataSet();
				DataSet oDSVentasRB = new DataSet();
				DataSet oDSVentasRB1 = new DataSet();
				cdsSeteoF.Schema.Connections[0].Open();
				C1.Data.SchemaObjects.Connection miConeccion = cdsSeteoF.Schema.Connections[0];
				miConeccion = cdsSeteoF.Schema.Connections[0];
				#endregion Variables

				this.saveFileDialog1.Filter = "XML (*.xml)|*.xml";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{	
					//					try
					//					{
					#region IVA
					string stArchivo = saveFileDialog1.FileName; // stArchivo = @"C:\XML_ANEXO_TRANSACCIONAL\" + stArchivo; NombreGenerado("AT");
					
					XmlTextWriter miXML = new XmlTextWriter(stArchivo, Encoding.GetEncoding("UTF-8"));					
					//miXML.Formatting = Formatting.Indented;			//		
					miXML.WriteStartDocument(true);
          					
					miXML.WriteStartElement("iva");
			
					#region Encabezado
					SqlDataReader drEmpresa = Funcion.rEscalarSQL(cdsSeteoF, "Select TipoId, Ruc, Empresa From Seteo", true);
					drEmpresa.Read();
					if (drEmpresa.HasRows)
					{
						sTipoId = drEmpresa.GetString(0);
						sRUCEmpresa = drEmpresa.GetString(1);
						sEmpresa = drEmpresa.GetString(2);
					}
					drEmpresa.Close();

					miXML.WriteElementString("TipoIDInformante", sTipoId);
					miXML.WriteElementString("IdInformante", sRUCEmpresa);
					miXML.WriteElementString("razonSocial", sEmpresa);
					miXML.WriteElementString("Anio", dtFecha.ToString("yyyy"));
					miXML.WriteElementString("Mes", dtHasta.ToString("MM"));
					
					miXML.WriteElementString("regimenMicroempresa", regEmpresa = "SI");
					miXML.WriteElementString("numEstabRuc", Funcion.iEscalarSQL(cdsSeteoF, 
						string.Format("Select COUNT(*) From Bodega Where Activo = 1 And Factura = 1 And Principal = 1 And '{0}' Between Convert(Date, FechaApertura) And Convert(Date, ISNULL(FechaCierre, '20301231'))", 
						dtHasta.ToString("yyyyMMdd"))).ToString().PadLeft(3, Pad)); //+ Iva - DescPorcIva  - DescPorcIva0 
					string sSQLVentaTotal = string.Format("Exec dbo.ATSVentasTotales '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					dVentaTotal = Funcion.decEscalarSQL(cdsSeteoF, sSQLVentaTotal);

					miXML.WriteElementString("totalVentas", dVentaTotal.ToString("0.00", us));//Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select SUM((SubtotalIva - Descuento) + (SubtotalExcento - Descuento0)) From Compra Where idTipoFactura = 1 And Borrar = 0 And Bodega Not In (50) And Estado In (10, 11) And CONVERT(Date, Fecha) Between '{0}' And '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"))).ToString("0.00", us));
					miXML.WriteElementString("codigoOperativo", "IVA");
					#endregion Encabezado

					#region Compras
					string sSQLCompras = string.Format("Exec ATSCompras '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					SqlDataAdapter oDACompras = new SqlDataAdapter(sSQLCompras, (SqlConnection) miConeccion.DbConnection);
					oDACompras.Fill(oDSCompras);

					if (oDSCompras.Tables[0].Rows.Count > 1)
					{
						miXML.WriteStartElement("compras");

						foreach(DataRow oDRCompras in oDSCompras.Tables[0].Rows)
						{
							#region Inicia Variables
							idCompra = (int)oDRCompras["idCompra"];
							idOrigen = (int)oDRCompras["idOrigen"];
							idTipoFactura = (int)oDRCompras["idTipoFactura"];
							dBaseNGrvIva = 0.00m;
							dBase0 = 0.00m;
							dBase = 0.00m;
							dICE = 0.00m;
							dIVA = 0.00m;
							dIVABienes10 = 0.00m;
							dIVAServicios20 = 0.00m;
							dIVABienes = 0.00m;
							dIVAServicios = 0.00m;
							dIVABienesServicios = 0.00m;
							#endregion Inicia Variables

							miXML.WriteStartElement("detalleCompras");
					
							#region Generales
							miXML.WriteElementString("codSustento", oDRCompras["Sustento"].ToString().Trim());
							miXML.WriteElementString("tpIdProv", oDRCompras["TipoRuc"].ToString().Trim());
							miXML.WriteElementString("idProv", oDRCompras["Ruc"].ToString().Trim());
							miXML.WriteElementString("tipoComprobante", oDRCompras["Comprobante"].ToString().Trim());
					
							miXML.WriteElementString("parteRel", "NO");

							if (oDRCompras["TipoRuc"].ToString().Trim() == "03")
							{
								miXML.WriteElementString("tipoProv", oDRCompras["TipoProveedor"].ToString().Trim());
								miXML.WriteElementString("denopr", oDRCompras["Proveedor"].ToString().Trim());
							}

							dtFechaRegistro = (DateTime)oDRCompras["FechaRegistro"];
					
							miXML.WriteElementString("fechaRegistro", dtFechaRegistro.ToString("dd/MM/yyyy"));

							//							if (oDRCompras["TipoRuc"].ToString().Trim() != "03")
							//							{
							miXML.WriteElementString("establecimiento", oDRCompras["SerieFactura"].ToString().Trim().Substring(0, 3));
							miXML.WriteElementString("puntoEmision", oDRCompras["SerieFactura"].ToString().Trim().Substring(3, 3));
							//							}
							//							else
							//							{
							//								miXML.WriteElementString("establecimiento", "999");
							//								miXML.WriteElementString("puntoEmision", "999");
							//							}

							miXML.WriteElementString("secuencial", oDRCompras["Numero"].ToString().Trim());

							dtFechaEmision = (DateTime)oDRCompras["Fecha"];
					
							miXML.WriteElementString("fechaEmision", dtFechaEmision.ToString("dd/MM/yyyy"));
							
							if (oDRCompras["TipoRuc"].ToString().Trim() != "03") miXML.WriteElementString("autorizacion", oDRCompras["AutFactura"].ToString());
							else miXML.WriteElementString("autorizacion", "9999999999");
							#endregion Generales

							#region Base IVA
							dBaseNGrvIva = (decimal)oDRCompras["BaseNoImp"];
							dBase0 = (decimal)oDRCompras["Base0"];
							dBase = (decimal)oDRCompras["Base"];
							dICE = Convert.ToDecimal(oDRCompras["Ice"]);
							dIVA = (decimal)oDRCompras["Iva"];

							miXML.WriteElementString("baseNoGraIva", dBaseNGrvIva.ToString("0.00", us));
							miXML.WriteElementString("baseImponible", dBase0.ToString("0.00", us));
							miXML.WriteElementString("baseImpGrav", dBase.ToString("0.00", us));
							miXML.WriteElementString("baseImpExe", dBaseNGrvIva.ToString("0.00", us));

							miXML.WriteElementString("montoIce", dICE.ToString("0.00", us));
							miXML.WriteElementString("montoIva", dIVA.ToString("0.00", us));
							#endregion Base IVA

							#region Retenciones de IVA
							dIVABienes10 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ATSRetencionIVA {0}, 10", idCompra));
							dIVAServicios20 = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ATSRetencionIVA {0}, 20", idCompra));
							dIVABienes = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ATSRetencionIVA {0}, 30", idCompra));
							dIVAServicios = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ATSRetencionIVA {0}, 70", idCompra));;
							dIVABienesServicios = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ATSRetencionIVA {0}, 100", idCompra));;
							decimal dBaseReembolsos = 0.00m;
							decimal dIVA50 = 0.00m;
							miXML.WriteElementString("valRetBien10", dIVABienes10.ToString("0.00", us));
							miXML.WriteElementString("valRetServ20", dIVAServicios20.ToString("0.00", us));
							miXML.WriteElementString("valorRetBienes", dIVABienes.ToString("0.00", us));
							miXML.WriteElementString("valRetServ50", dIVA50.ToString("0.00", us));
							miXML.WriteElementString("valorRetServicios", dIVAServicios.ToString("0.00", us));
							miXML.WriteElementString("valRetServ100", dIVABienesServicios.ToString("0.00", us));
							miXML.WriteElementString("totbasesImpReemb", dBaseReembolsos.ToString("0.00", us));
							//miXML.WriteElementString("valorRetencionNc", dvalorRetencionNc.ToString("0.00", us));
							#endregion Retenciones de IVA

							#region Pago Exterior o Local
							miXML.WriteStartElement("pagoExterior");
							miXML.WriteElementString("pagoLocExt", oDRCompras["TipoDePago"].ToString());

							if (oDRCompras["TipoDePago"].ToString() == "01")
							{
								miXML.WriteElementString("paisEfecPago", "NA");
								miXML.WriteElementString("aplicConvDobTrib", "NA");
								miXML.WriteElementString("pagExtSujRetNorLeg", "NA");
								miXML.WriteElementString("pagoRegFis", "NA");
							}
							else
							{
								miXML.WriteElementString("paisEfecPago", oDRCompras["CodigoPais"].ToString());
								string ACDT = "NO";
								miXML.WriteElementString("aplicConvDobTrib", ACDT);
								if (ACDT == "NO") miXML.WriteElementString("pagExtSujRetNorLeg", "SI");
							}				
							miXML.WriteEndElement();
							#endregion Pago Exterior o Local

							#region Forma de pago
							if ((decimal)oDRCompras["Total"] >= 800.00m)
							{
								miXML.WriteStartElement("formasDePago");
								miXML.WriteElementString("formaPago", oDRCompras["TipoFormaPago"].ToString());
								miXML.WriteEndElement();
							}
							#endregion Forma de pago					

							#region Retencion de Renta
							miXML.WriteStartElement("air");

							string sSQLAIR = string.Format("Exec ATSRetencionRenta {0}", idCompra);
							SqlDataReader drAIR = Funcion.rEscalarSQL(cdsSeteoF, sSQLAIR, true);

							while(drAIR.Read())
							{
								miXML.WriteStartElement("detalleAir");
					
								miXML.WriteElementString("codRetAir", drAIR.GetString(0).Trim());
								miXML.WriteElementString("baseImpAir", drAIR.GetDecimal(1).ToString("0.00", us));
								miXML.WriteElementString("porcentajeAir", drAIR.GetDecimal(2).ToString("0.00", us));
								miXML.WriteElementString("valRetAir", drAIR.GetDecimal(3).ToString("0.00", us));
					
								miXML.WriteEndElement();
							}
							drAIR.Close();

							miXML.WriteEndElement();
							#endregion Retencion de Renta

							#region Comprobante de Retencion
							if (Convert.ToDecimal(oDRCompras["Retenido"]) > 0.00m)
							{
								miXML.WriteElementString("estabRetencion1", oDRCompras["Serie"].ToString().Trim().Substring(0,3));
								miXML.WriteElementString("ptoEmiRetencion1", oDRCompras["Serie"].ToString().Trim().Substring(3,3));
								miXML.WriteElementString("secRetencion1", oDRCompras["NumRet"].ToString().Trim());
								miXML.WriteElementString("autRetencion1", oDRCompras["Autorizacion"].ToString().Trim());
								dtFechaRet = (DateTime)oDRCompras["FecRet"];
								miXML.WriteElementString("fechaEmiRet1", dtFechaRet.ToString("dd/MM/yyyy"));
							}
							#endregion Comprobante de Retencion

							#region Documento modificado
							if (idTipoFactura != 4)
							{
								if (idOrigen > 0)
								{
									string sSQLMod = string.Format("Select cc.Codigo, c.SerieFactura, Numero, AutFactura From Compra c Inner Join CompraComprobante cc On cc.idComprobante = c.idComprobante Where idTipoFactura = 4 And idCompra = {0}", idOrigen);
									SqlDataReader drModificado = Funcion.rEscalarSQL(cdsSeteoF, sSQLMod, true);
									drModificado.Read();
									if (drModificado.HasRows)
									{
										miXML.WriteElementString("docModificado", drModificado.GetString(0).ToString().Trim());// int.Parse(stDevCompTipo[k]).ToString());
										miXML.WriteElementString("estabModificado", drModificado.GetString(1).ToString().Trim().Substring(0,3));
										miXML.WriteElementString("ptoEmiModificado", drModificado.GetString(1).ToString().Trim().Substring(3,3));
										miXML.WriteElementString("secModificado", drModificado.GetString(2).ToString().Trim());
										miXML.WriteElementString("autModificado", drModificado.GetString(3).ToString().Trim());
									}
									drModificado.Close();
								}
								else
								{
									string sSQLMod = string.Format("Select cc.Codigo, c.SerieFactura, Numero, AutFactura From Compra c Inner Join CompraComprobante cc On cc.idComprobante = c.idComprobante Where idCompra = (Select idOrigen From AutorizacionDePagos Where idTipo = 4 And idAutorizacionPago = (Select idAutorizacionPago From AutPagoNC where idCompra = {0}))", idCompra);
									SqlDataReader drModificado = Funcion.rEscalarSQL(cdsSeteoF, sSQLMod, true);
									drModificado.Read();
									if (drModificado.HasRows)
									{
										miXML.WriteElementString("docModificado", drModificado.GetString(0).ToString().Trim());// int.Parse(stDevCompTipo[k]).ToString());
										miXML.WriteElementString("estabModificado", drModificado.GetString(1).ToString().Trim().Substring(0,3));
										miXML.WriteElementString("ptoEmiModificado", drModificado.GetString(1).ToString().Trim().Substring(3,3));
										miXML.WriteElementString("secModificado", drModificado.GetString(2).ToString().Trim());
										miXML.WriteElementString("autModificado", drModificado.GetString(3).ToString().Trim());
									}
									drModificado.Close();
								}
							}
							#endregion Documento modificado

							miXML.WriteEndElement();
						}			
						miXML.WriteEndElement();
					}					
					#endregion Compras

					#region Ventas
					miXML.WriteStartElement("ventas");
					string sSQLVentas = string.Format("Exec ATSVentas '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					SqlDataAdapter oDAVentas = new SqlDataAdapter(sSQLVentas, (SqlConnection) miConeccion.DbConnection);
					oDAVentas.Fill(oDSVentas);

					if (oDSVentas.Tables[0].Rows.Count > 0)
					{
						foreach(DataRow oDRVentas in oDSVentas.Tables[0].Rows)
						{
							dVentaBaseNGIVA = 0;
							dVentaBase0 = 0;
							dVentaBase = 0;
							dIVA = 0;
							dVentaRetIVA = 0;
							dVentaRetRenta = 0;
							dICE = 0.00m;

							miXML.WriteStartElement("detalleVentas");
							miXML.WriteElementString("tpIdCliente", oDRVentas["CodTipoRuc"].ToString().Trim());
							miXML.WriteElementString("idCliente", oDRVentas["Ruc"].ToString().Trim());
							if (oDRVentas["CodTipoRuc"].ToString().Trim() != "07") miXML.WriteElementString("parteRelVtas", "NO");
								
							if (oDRVentas["CodTipoRuc"].ToString().Trim() == "06") 
							{
								miXML.WriteElementString("tipoCliente", "02");
								miXML.WriteElementString("denoCli", oDRVentas["Nombre"].ToString().Trim());//Empresa_Pas
							}
								
							miXML.WriteElementString("tipoComprobante", oDRVentas["CodComprobante"].ToString().Trim());
							miXML.WriteElementString("tipoEmision", oDRVentas["TipoEmision"].ToString().Trim());
							miXML.WriteElementString("numeroComprobantes", oDRVentas["Registros"].ToString().Trim());

							miXML.WriteElementString("baseNoGraIva", oDRVentas["BaseNGrIVA"].ToString());//"0.00", us));
							miXML.WriteElementString("baseImponible", oDRVentas["Base0"].ToString());//"0.00", us));
							miXML.WriteElementString("baseImpGrav", oDRVentas["Base"].ToString());//"0.00", us));
							miXML.WriteElementString("montoIva", oDRVentas["IVA"].ToString());//"0.00", us));
							//Tipo de Compensaciones
							//Monto de Compensaciones
							miXML.WriteElementString("montoIce", dICE.ToString("0.00", us));
							miXML.WriteElementString("valorRetIva", oDRVentas["IVARetenido"].ToString());//"0.00", us));
							miXML.WriteElementString("valorRetRenta", oDRVentas["RentaRetenido"].ToString());//"0.00", us));

							if (oDRVentas["CodComprobante"].ToString().Trim() == "18")
							{
								string sSQLVFP = string.Format("Exec ATSVentasFormasDePago '{0}', '{1}', '{2}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), oDRVentas["Ruc"].ToString().Trim());
								SqlDataReader drVFP = Funcion.rEscalarSQL(cdsSeteoF, sSQLVFP, true);

								//								drVFP.Read();
								//								if (drVFP.HasRows)
								//								{
								//									if (drVFP.GetValue(0) == System.DBNull.Value) 
								//									{
								//										MessageBox.Show(oDRVentas["Ruc"].ToString().Trim());
								//										drVFP.Close();
								//										return;
								//									}									
								//								}
								//								else 
								//								{
								//									MessageBox.Show(oDRVentas["Ruc"].ToString().Trim());
								//									drVFP.Close();
								//									return;
								//								}							
								//								if (oDRVentas["Ruc"].ToString().Trim() == "0131986288000")
								//									MessageBox.Show(oDRVentas["Ruc"].ToString().Trim());

								miXML.WriteStartElement("formasDePago");
								while(drVFP.Read())
								{
									miXML.WriteElementString("formaPago", drVFP.GetString(0).Trim());
								}
								drVFP.Close();
								
								miXML.WriteEndElement();
							}

							miXML.WriteEndElement();
						}
					}
					//miXML.WriteEndElement();
					#endregion Ventas

					#region Ventas - Retenciones Bancarias
					string sSQLVentasRB = string.Format("Exec ATSVentasRetencionesBancarias '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					SqlDataAdapter oDAVentasRB = new SqlDataAdapter(sSQLVentasRB, (SqlConnection) miConeccion.DbConnection);
					oDAVentasRB.Fill(oDSVentasRB);

					if (oDSVentasRB.Tables[0].Rows.Count > 0)
					{
						foreach(DataRow oDRVentasRB in oDSVentasRB.Tables[0].Rows)
						{				
							dICE = 0.00m;

							miXML.WriteStartElement("detalleVentas");
							miXML.WriteElementString("tpIdCliente", oDRVentasRB["CodTipoRuc"].ToString().Trim());
							miXML.WriteElementString("idCliente", oDRVentasRB["Ruc"].ToString().Trim());
							miXML.WriteElementString("parteRelVtas", "NO");
																								
							miXML.WriteElementString("tipoComprobante", oDRVentasRB["CodComprobante"].ToString().Trim());
							miXML.WriteElementString("tipoEmision", oDRVentasRB["TipoEmision"].ToString().Trim());
							miXML.WriteElementString("numeroComprobantes", oDRVentasRB["Registros"].ToString().Trim());

							miXML.WriteElementString("baseNoGraIva", oDRVentasRB["BaseNGrIVA"].ToString());//"0.00", us));
							miXML.WriteElementString("baseImponible", oDRVentasRB["Base0"].ToString());//"0.00", us));
							miXML.WriteElementString("baseImpGrav", oDRVentasRB["Base"].ToString());//"0.00", us));
							miXML.WriteElementString("montoIva", oDRVentasRB["IVA"].ToString());//"0.00", us));
							//Tipo de Compensaciones
							//Monto de Compensaciones
							miXML.WriteElementString("montoIce", dICE.ToString("0.00", us));
							miXML.WriteElementString("valorRetIva", oDRVentasRB["IVARetenido"].ToString());//"0.00", us));
							miXML.WriteElementString("valorRetRenta", oDRVentasRB["RentaRetenido"].ToString());//"0.00", us));

							miXML.WriteStartElement("formasDePago");								
							miXML.WriteElementString("formaPago", "01");								
							miXML.WriteEndElement();
								
							miXML.WriteEndElement();
						}
					}
					//miXML.WriteEndElement();
					#endregion Ventas - Retenciones Bancarias
					
					#region Ventas - Retenciones Bancarias - Temporal Factura Banco del Pacifico
					if (MenuLatinium.stDirFacturacion == "POINT")
					{
						string sSQLVentasRB1 = string.Format("Exec ATSVentasRB '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
						SqlDataAdapter oDAVentasRB1 = new SqlDataAdapter(sSQLVentasRB1, (SqlConnection) miConeccion.DbConnection);
						oDAVentasRB1.Fill(oDSVentasRB1);

						if (oDSVentasRB1.Tables[0].Rows.Count > 0)
						{
							foreach(DataRow oDRVentasRB1 in oDSVentasRB1.Tables[0].Rows)
							{				
								dICE = 0.00m;

								miXML.WriteStartElement("detalleVentas");
								miXML.WriteElementString("tpIdCliente", oDRVentasRB1["CodTipoRuc"].ToString().Trim());
								miXML.WriteElementString("idCliente", oDRVentasRB1["Ruc"].ToString().Trim());
								miXML.WriteElementString("parteRelVtas", "NO");
																								
								miXML.WriteElementString("tipoComprobante", oDRVentasRB1["CodComprobante"].ToString().Trim());
								miXML.WriteElementString("tipoEmision", oDRVentasRB1["TipoEmision"].ToString().Trim());
								miXML.WriteElementString("numeroComprobantes", oDRVentasRB1["Registros"].ToString().Trim());

								miXML.WriteElementString("baseNoGraIva", oDRVentasRB1["BaseNGrIVA"].ToString());//"0.00", us));
								miXML.WriteElementString("baseImponible", oDRVentasRB1["Base0"].ToString());//"0.00", us));
								miXML.WriteElementString("baseImpGrav", oDRVentasRB1["Base"].ToString());//"0.00", us));
								miXML.WriteElementString("montoIva", oDRVentasRB1["IVA"].ToString());//"0.00", us));
								//Tipo de Compensaciones
								//Monto de Compensaciones
								miXML.WriteElementString("montoIce", dICE.ToString("0.00", us));
								miXML.WriteElementString("valorRetIva", oDRVentasRB1["IVARetenido"].ToString());//"0.00", us));
								miXML.WriteElementString("valorRetRenta", oDRVentasRB1["RentaRetenido"].ToString());//"0.00", us));

								miXML.WriteStartElement("formasDePago");								
								miXML.WriteElementString("formaPago", "01");								
								miXML.WriteEndElement();
								
								miXML.WriteEndElement();
							}
						}						
					}					
					#endregion Ventas - Retenciones Bancarias - Temporal Factura Banco del Pacifico
					miXML.WriteEndElement();

					#region Ventas Establecimiento
					miXML.WriteStartElement("ventasEstablecimiento");
			
					int iContBodegas = (int)Funcion.iEscalarSQL(cdsSeteoF, "Select Count(*) From Bodega");
								
					string sSerie = "";
					int iLocales = 0;

					for (int l = 1; l <= iContBodegas; l++)
					{	
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Activo Bit = 0 If Exists(Select Serie From Bodega Where Activo = 1 And Factura = 1 And Principal = 1 And Bodega = {0} And '{1}' Between Convert(Date, FechaApertura) And Convert(Date, ISNULL(FechaCierre, '20301231'))) Set @Activo = 1 Select @Activo", l, dtHasta.ToString("yyyyMMdd")), true))
						{
							sSerie = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Serie From Bodega Where Activo = 1 And Factura = 1 And Principal = 1 And Bodega = {0}", l), true);
							
							if (sSerie.Length > 0)
							{
								miXML.WriteStartElement("ventaEst");

								miXML.WriteElementString("codEstab", Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select SUBSTRING(Serie, 1, 3) From Bodega Where Bodega = {0} And Principal = 1 And Activo = 1 And Factura = 1", l)));

								string sSQLVentaEstab = string.Format("Exec dbo.ATSVentasEstablecimiento '{0}', '{1}', '{2}'", 
									dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), sSerie);
								dVentaEstablecimiento = Funcion.decEscalarSQL(cdsSeteoF, sSQLVentaEstab);

								iLocales++;
 
								miXML.WriteElementString("ventasEstab", dVentaEstablecimiento.ToString("0.00", us));
					
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Insert Into TempVentaEstab (Serie, Valor) Values ('{0}', {1})", sSerie, dVentaEstablecimiento));

								miXML.WriteEndElement();
							}
						}
					}	
					miXML.WriteEndElement();
					
					//MessageBox.Show(iLocales.ToString());
					#endregion Ventas Establecimiento

					#region Comprobantes Anulados
					#region Facturas - Notas de crédito
					string sSQLFacAnuladas = string.Format("Exec ATSComprobantesAnulados '{0}', '{1}', 0", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					SqlDataReader drFacAnuladas = Funcion.rEscalarSQL(cdsSeteoF, sSQLFacAnuladas, true);

					miXML.WriteStartElement("anulados");

					while (drFacAnuladas.Read())
					{
						miXML.WriteStartElement("detalleAnulados");
				
						miXML.WriteElementString("tipoComprobante", drFacAnuladas.GetString(0).ToString().Trim());
						miXML.WriteElementString("establecimiento", drFacAnuladas.GetString(1).ToString().Trim().Substring(0, 3));
						miXML.WriteElementString("puntoEmision", drFacAnuladas.GetString(1).ToString().Trim().Substring(3, 3));
						miXML.WriteElementString("secuencialInicio", drFacAnuladas.GetString(2).ToString().Trim());
						miXML.WriteElementString("secuencialFin", drFacAnuladas.GetString(2).ToString().Trim());
						miXML.WriteElementString("autorizacion", drFacAnuladas.GetString(3).ToString().Trim());

						miXML.WriteEndElement();
					}
					drFacAnuladas.Close();
					#endregion Facturas - Notas de crédito
			
					#region Retenciones
					string sSQLRetAnuladas = string.Format("Exec ATSComprobantesAnulados '{0}', '{1}', 1", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
					SqlDataReader drRetAnuladas = Funcion.rEscalarSQL(cdsSeteoF, sSQLRetAnuladas, true);

					while (drRetAnuladas.Read())
					{
						miXML.WriteStartElement("detalleAnulados");
				
						miXML.WriteElementString("tipoComprobante", drRetAnuladas.GetString(0).ToString().Trim());
						miXML.WriteElementString("establecimiento", drRetAnuladas.GetString(1).ToString().Trim().Substring(0, 3));
						miXML.WriteElementString("puntoEmision", drRetAnuladas.GetString(1).ToString().Trim().Substring(3, 3));
						miXML.WriteElementString("secuencialInicio", drRetAnuladas.GetString(2).ToString().Trim());
						miXML.WriteElementString("secuencialFin", drRetAnuladas.GetString(2).ToString().Trim());
						miXML.WriteElementString("autorizacion", drRetAnuladas.GetString(3).ToString().Trim());

						miXML.WriteEndElement();
					}
					drRetAnuladas.Close();
					#endregion Retenciones
					#endregion Comprobantes Anulados

					miXML.WriteEndElement();

					miXML.WriteEndElement(); 
						
					miXML.WriteEndDocument();

					miXML.Flush();
					miXML.Close();
					#endregion IVA
			
					MessageBox.Show("Anexo Transaccional Simplificado " + stArchivo + " Generado " + Funcion.Tiempo("En"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				
					//					}
					//					catch (Exception ex)
					//					{						
					//						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
					//						this.Cursor = Cursors.Default;	
					//						return;											
					//					}
				}
			}

			Cursor = Cursors.Default;
		}

		private void frmdetalleATSSemestral_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
