using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFAC_DescuentosValidacion.
	/// </summary>
	public class frmFAC_DescuentosValidacion : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblVendedor;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProducto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidBanco;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFormaDePago;
		private System.Windows.Forms.Label lblPlanes;
		private System.Windows.Forms.Label lblTarjeta;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbPlanes;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblPlazo;
		private System.Windows.Forms.Label lblBanco;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbPlazo;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTarjeta;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.LinkLabel linkLabel1;
		public System.Windows.Forms.Label lblIce;
		public System.Windows.Forms.Label label8;
		public System.Windows.Forms.Label label9;
		public System.Windows.Forms.Label label10;
		public System.Windows.Forms.Label label11;
		public System.Windows.Forms.Label label12;
		public System.Windows.Forms.Label label13;
		public System.Windows.Forms.Label label14;
		public System.Windows.Forms.Label label15;
		public System.Windows.Forms.Label label16;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbPlanesC;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbPlazosC;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTarjetaC;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBancoC;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFormaPagoC;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCompetencia;
		private System.Windows.Forms.Button btnCalcular;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioCompetencia;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCosto;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecio;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenMinimoPorcentaje;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtComisionTarjeta;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProteccion;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenMinimoValor;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioMinimoVenta;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescuento;
		public System.Windows.Forms.Label label17;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCostoConIVA;
		public System.Windows.Forms.Label label18;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCostoFinal;
		public System.Windows.Forms.Label label19;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCompetenciaPrecioMinimo;
		public System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidArticulo;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenIgualarCompetencia;
		public System.Windows.Forms.Label label21;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFAC_DescuentosValidacion()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFAC_DescuentosValidacion));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			this.lblNumero = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.txtProducto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidBanco = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.optFormaDePago = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblPlanes = new System.Windows.Forms.Label();
			this.lblTarjeta = new System.Windows.Forms.Label();
			this.cmbPlanes = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblPlazo = new System.Windows.Forms.Label();
			this.lblBanco = new System.Windows.Forms.Label();
			this.cmbPlazo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTarjeta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbPlanesC = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbPlazosC = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTarjetaC = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBancoC = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.optFormaPagoC = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtCompetencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			this.lblIce = new System.Windows.Forms.Label();
			this.txtPrecioCompetencia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtCosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtPrecio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.txtMargenMinimoPorcentaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtComisionTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtProteccion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.txtMargenMinimoValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtPrecioMinimoVenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.txtDescuento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnCalcular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.label17 = new System.Windows.Forms.Label();
			this.txtCostoConIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.txtCostoFinal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.txtCompetenciaPrecioMinimo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.txtMargenIgualarCompetencia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label21 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProducto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFormaDePago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanesC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazosC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjetaC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBancoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFormaPagoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCompetencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioCompetencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenMinimoPorcentaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComisionTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProteccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenMinimoValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioMinimoVenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCostoConIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCostoFinal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCompetenciaPrecioMinimo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenIgualarCompetencia)).BeginInit();
			this.SuspendLayout();
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(528, 8);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 17);
			this.lblNumero.TabIndex = 198;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 197;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(8, 56);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 17);
			this.lblVendedor.TabIndex = 196;
			this.lblVendedor.Text = "Vendedor";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 32);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 17);
			this.lblBodega.TabIndex = 195;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 191;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// cmbBodega
			// 
			appearance2.FontData.Name = "Tahoma";
			this.cmbBodega.Appearance = appearance2;
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(72, 32);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(232, 22);
			this.cmbBodega.TabIndex = 192;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// txtNumero
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance3;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(584, 8);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 194;
			// 
			// cmbVendedor
			// 
			appearance4.FontData.Name = "Tahoma";
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance4;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 335;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(72, 56);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(232, 22);
			this.cmbVendedor.TabIndex = 193;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 80);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(42, 16);
			this.label7.TabIndex = 200;
			this.label7.Text = "Artículo";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtProducto
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProducto.Appearance = appearance5;
			this.txtProducto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProducto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProducto.Location = new System.Drawing.Point(72, 80);
			this.txtProducto.Name = "txtProducto";
			this.txtProducto.Size = new System.Drawing.Size(264, 22);
			this.txtProducto.TabIndex = 199;
			this.txtProducto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtProducto_KeyDown);
			// 
			// txtidBanco
			// 
			this.txtidBanco.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidBanco.Enabled = false;
			this.txtidBanco.Location = new System.Drawing.Point(344, 8);
			this.txtidBanco.Name = "txtidBanco";
			this.txtidBanco.PromptChar = ' ';
			this.txtidBanco.Size = new System.Drawing.Size(16, 22);
			this.txtidBanco.TabIndex = 274;
			this.txtidBanco.Visible = false;
			// 
			// txtidArticulo
			// 
			this.txtidArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidArticulo.Enabled = false;
			this.txtidArticulo.Location = new System.Drawing.Point(360, 8);
			this.txtidArticulo.Name = "txtidArticulo";
			this.txtidArticulo.PromptChar = ' ';
			this.txtidArticulo.Size = new System.Drawing.Size(16, 22);
			this.txtidArticulo.TabIndex = 273;
			this.txtidArticulo.Visible = false;
			// 
			// optFormaDePago
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFormaDePago.Appearance = appearance6;
			this.optFormaDePago.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFormaDePago.CheckedIndex = 0;
			this.optFormaDePago.ItemAppearance = appearance7;
			this.optFormaDePago.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Efectivo";
			valueListItem2.DataValue = 3;
			valueListItem2.DisplayText = "Tarjeta";
			valueListItem3.DataValue = 92;
			valueListItem3.DisplayText = "Tarjeta Sin Interés";
			this.optFormaDePago.Items.Add(valueListItem1);
			this.optFormaDePago.Items.Add(valueListItem2);
			this.optFormaDePago.Items.Add(valueListItem3);
			this.optFormaDePago.ItemSpacingVertical = 10;
			this.optFormaDePago.Location = new System.Drawing.Point(8, 112);
			this.optFormaDePago.Name = "optFormaDePago";
			this.optFormaDePago.Size = new System.Drawing.Size(264, 24);
			this.optFormaDePago.TabIndex = 363;
			this.optFormaDePago.Text = "Efectivo";
			this.optFormaDePago.ValueChanged += new System.EventHandler(this.optFormaDePago_ValueChanged);
			// 
			// lblPlanes
			// 
			this.lblPlanes.AutoSize = true;
			this.lblPlanes.BackColor = System.Drawing.Color.Transparent;
			this.lblPlanes.Location = new System.Drawing.Point(8, 200);
			this.lblPlanes.Name = "lblPlanes";
			this.lblPlanes.Size = new System.Drawing.Size(36, 17);
			this.lblPlanes.TabIndex = 371;
			this.lblPlanes.Text = "Planes";
			this.lblPlanes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblTarjeta
			// 
			this.lblTarjeta.AutoSize = true;
			this.lblTarjeta.BackColor = System.Drawing.Color.Transparent;
			this.lblTarjeta.Location = new System.Drawing.Point(8, 176);
			this.lblTarjeta.Name = "lblTarjeta";
			this.lblTarjeta.Size = new System.Drawing.Size(40, 17);
			this.lblTarjeta.TabIndex = 370;
			this.lblTarjeta.Text = "Tarjeta";
			this.lblTarjeta.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbPlanes
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPlanes.Appearance = appearance8;
			this.cmbPlanes.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPlanes.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlanes.DisplayMember = "Descripcion";
			this.cmbPlanes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlanes.Enabled = false;
			this.cmbPlanes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlanes.Location = new System.Drawing.Point(64, 192);
			this.cmbPlanes.Name = "cmbPlanes";
			this.cmbPlanes.Size = new System.Drawing.Size(240, 22);
			this.cmbPlanes.TabIndex = 366;
			this.cmbPlanes.ValueMember = "idPlan";
			this.cmbPlanes.ValueChanged += new System.EventHandler(this.cmbPlanes_ValueChanged);
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
			// lblPlazo
			// 
			this.lblPlazo.AutoSize = true;
			this.lblPlazo.BackColor = System.Drawing.Color.Transparent;
			this.lblPlazo.Location = new System.Drawing.Point(8, 224);
			this.lblPlazo.Name = "lblPlazo";
			this.lblPlazo.Size = new System.Drawing.Size(30, 17);
			this.lblPlazo.TabIndex = 369;
			this.lblPlazo.Text = "Plazo";
			this.lblPlazo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBanco
			// 
			this.lblBanco.AutoSize = true;
			this.lblBanco.BackColor = System.Drawing.Color.Transparent;
			this.lblBanco.Location = new System.Drawing.Point(8, 152);
			this.lblBanco.Name = "lblBanco";
			this.lblBanco.Size = new System.Drawing.Size(34, 17);
			this.lblBanco.TabIndex = 368;
			this.lblBanco.Text = "Banco";
			this.lblBanco.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbPlazo
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPlazo.Appearance = appearance9;
			this.cmbPlazo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPlazo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlazo.DisplayMember = "Descripcion";
			this.cmbPlazo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlazo.Enabled = false;
			this.cmbPlazo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlazo.Location = new System.Drawing.Point(64, 216);
			this.cmbPlazo.Name = "cmbPlazo";
			this.cmbPlazo.Size = new System.Drawing.Size(240, 22);
			this.cmbPlazo.TabIndex = 367;
			this.cmbPlazo.ValueMember = "idPlazo";
			// 
			// cmbTarjeta
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTarjeta.Appearance = appearance10;
			this.cmbTarjeta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTarjeta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTarjeta.DisplayMember = "Nombre";
			this.cmbTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTarjeta.Enabled = false;
			this.cmbTarjeta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTarjeta.Location = new System.Drawing.Point(64, 168);
			this.cmbTarjeta.Name = "cmbTarjeta";
			this.cmbTarjeta.Size = new System.Drawing.Size(240, 22);
			this.cmbTarjeta.TabIndex = 365;
			this.cmbTarjeta.ValueMember = "idTarjeta";
			this.cmbTarjeta.ValueChanged += new System.EventHandler(this.cmbTarjeta_ValueChanged);
			// 
			// cmbBanco
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance11;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataSource = this.ultraDataSource1;
			ultraGridBand3.AddButtonCaption = "Compra";
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Width = 224;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.Enabled = false;
			this.cmbBanco.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(64, 144);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(240, 22);
			this.cmbBanco.TabIndex = 364;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.ValueChanged += new System.EventHandler(this.cmbBanco_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbPlanesC
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPlanesC.Appearance = appearance12;
			this.cmbPlanesC.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPlanesC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlanesC.DisplayMember = "Descripcion";
			this.cmbPlanesC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlanesC.Enabled = false;
			this.cmbPlanesC.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlanesC.Location = new System.Drawing.Point(440, 192);
			this.cmbPlanesC.Name = "cmbPlanesC";
			this.cmbPlanesC.Size = new System.Drawing.Size(240, 22);
			this.cmbPlanesC.TabIndex = 375;
			this.cmbPlanesC.ValueMember = "idPlan";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(360, 224);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(30, 17);
			this.label1.TabIndex = 378;
			this.label1.Text = "Plazo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(360, 152);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 17);
			this.label2.TabIndex = 377;
			this.label2.Text = "Banco";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbPlazosC
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbPlazosC.Appearance = appearance13;
			this.cmbPlazosC.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPlazosC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlazosC.DisplayMember = "Descripcion";
			this.cmbPlazosC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPlazosC.Enabled = false;
			this.cmbPlazosC.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPlazosC.Location = new System.Drawing.Point(440, 216);
			this.cmbPlazosC.Name = "cmbPlazosC";
			this.cmbPlazosC.Size = new System.Drawing.Size(240, 22);
			this.cmbPlazosC.TabIndex = 376;
			this.cmbPlazosC.ValueMember = "idPlazo";
			// 
			// cmbTarjetaC
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTarjetaC.Appearance = appearance14;
			this.cmbTarjetaC.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTarjetaC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTarjetaC.DisplayMember = "Nombre";
			this.cmbTarjetaC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTarjetaC.Enabled = false;
			this.cmbTarjetaC.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTarjetaC.Location = new System.Drawing.Point(440, 168);
			this.cmbTarjetaC.Name = "cmbTarjetaC";
			this.cmbTarjetaC.Size = new System.Drawing.Size(240, 22);
			this.cmbTarjetaC.TabIndex = 374;
			this.cmbTarjetaC.ValueMember = "idTarjeta";
			// 
			// cmbBancoC
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBancoC.Appearance = appearance15;
			this.cmbBancoC.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBancoC.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBancoC.DataSource = this.ultraDataSource1;
			ultraGridBand4.AddButtonCaption = "Compra";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Width = 224;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbBancoC.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBancoC.DisplayMember = "Nombre";
			this.cmbBancoC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBancoC.Enabled = false;
			this.cmbBancoC.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBancoC.Location = new System.Drawing.Point(440, 144);
			this.cmbBancoC.Name = "cmbBancoC";
			this.cmbBancoC.Size = new System.Drawing.Size(240, 22);
			this.cmbBancoC.TabIndex = 373;
			this.cmbBancoC.ValueMember = "idBanco";
			// 
			// optFormaPagoC
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFormaPagoC.Appearance = appearance16;
			this.optFormaPagoC.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFormaPagoC.CheckedIndex = 0;
			this.optFormaPagoC.ItemAppearance = appearance17;
			this.optFormaPagoC.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "Efectivo";
			valueListItem5.DataValue = 3;
			valueListItem5.DisplayText = "Tarjeta";
			valueListItem6.DataValue = 92;
			valueListItem6.DisplayText = "Tarjeta Sin Interés";
			this.optFormaPagoC.Items.Add(valueListItem4);
			this.optFormaPagoC.Items.Add(valueListItem5);
			this.optFormaPagoC.Items.Add(valueListItem6);
			this.optFormaPagoC.ItemSpacingVertical = 10;
			this.optFormaPagoC.Location = new System.Drawing.Point(360, 112);
			this.optFormaPagoC.Name = "optFormaPagoC";
			this.optFormaPagoC.Size = new System.Drawing.Size(264, 24);
			this.optFormaPagoC.TabIndex = 372;
			this.optFormaPagoC.Text = "Efectivo";
			this.optFormaPagoC.ValueChanged += new System.EventHandler(this.optFormaPagoC_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(360, 200);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(36, 17);
			this.label3.TabIndex = 380;
			this.label3.Text = "Planes";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(360, 176);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(40, 17);
			this.label4.TabIndex = 379;
			this.label4.Text = "Tarjeta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(360, 80);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(71, 16);
			this.label5.TabIndex = 382;
			this.label5.Text = "Competencia";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCompetencia
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCompetencia.Appearance = appearance18;
			this.txtCompetencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCompetencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCompetencia.Location = new System.Drawing.Point(448, 80);
			this.txtCompetencia.Name = "txtCompetencia";
			this.txtCompetencia.Size = new System.Drawing.Size(232, 22);
			this.txtCompetencia.TabIndex = 381;
			// 
			// linkLabel1
			// 
			this.linkLabel1.AutoSize = true;
			this.linkLabel1.Location = new System.Drawing.Point(440, 248);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(173, 17);
			this.linkLabel1.TabIndex = 393;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "Subir Imagen Precio Competencia";
			this.linkLabel1.Visible = false;
			// 
			// lblIce
			// 
			this.lblIce.AutoSize = true;
			this.lblIce.Location = new System.Drawing.Point(8, 275);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(123, 17);
			this.lblIce.TabIndex = 395;
			this.lblIce.Text = "(+) Precio Competencia";
			this.lblIce.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPrecioCompetencia
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioCompetencia.Appearance = appearance19;
			this.txtPrecioCompetencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioCompetencia.FormatString = "#,##0.00";
			this.txtPrecioCompetencia.Location = new System.Drawing.Point(144, 272);
			this.txtPrecioCompetencia.Name = "txtPrecioCompetencia";
			this.txtPrecioCompetencia.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioCompetencia.PromptChar = ' ';
			this.txtPrecioCompetencia.Size = new System.Drawing.Size(112, 22);
			this.txtPrecioCompetencia.TabIndex = 394;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 371);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(53, 17);
			this.label8.TabIndex = 397;
			this.label8.Text = "(+) Costo";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCosto
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCosto.Appearance = appearance20;
			this.txtCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCosto.Enabled = false;
			this.txtCosto.FormatString = "#,##0.00";
			this.txtCosto.Location = new System.Drawing.Point(144, 368);
			this.txtCosto.Name = "txtCosto";
			this.txtCosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCosto.PromptChar = ' ';
			this.txtCosto.Size = new System.Drawing.Size(112, 22);
			this.txtCosto.TabIndex = 396;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(8, 347);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(63, 17);
			this.label9.TabIndex = 399;
			this.label9.Text = "Precio Point";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPrecio
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecio.Appearance = appearance21;
			this.txtPrecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecio.Enabled = false;
			this.txtPrecio.FormatString = "#,##0.00";
			this.txtPrecio.Location = new System.Drawing.Point(144, 344);
			this.txtPrecio.Name = "txtPrecio";
			this.txtPrecio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecio.PromptChar = ' ';
			this.txtPrecio.Size = new System.Drawing.Size(112, 22);
			this.txtPrecio.TabIndex = 398;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(8, 467);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(81, 17);
			this.label10.TabIndex = 405;
			this.label10.Text = "Margen Minimo";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMargenMinimoPorcentaje
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenMinimoPorcentaje.Appearance = appearance22;
			this.txtMargenMinimoPorcentaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenMinimoPorcentaje.Enabled = false;
			this.txtMargenMinimoPorcentaje.FormatString = "#,##0.00";
			this.txtMargenMinimoPorcentaje.Location = new System.Drawing.Point(272, 464);
			this.txtMargenMinimoPorcentaje.Name = "txtMargenMinimoPorcentaje";
			this.txtMargenMinimoPorcentaje.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenMinimoPorcentaje.PromptChar = ' ';
			this.txtMargenMinimoPorcentaje.Size = new System.Drawing.Size(112, 22);
			this.txtMargenMinimoPorcentaje.TabIndex = 404;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(8, 419);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(109, 17);
			this.label11.TabIndex = 403;
			this.label11.Text = "(+) Comisión Tarjeta";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtComisionTarjeta
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComisionTarjeta.Appearance = appearance23;
			this.txtComisionTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComisionTarjeta.Enabled = false;
			this.txtComisionTarjeta.FormatString = "#,##0.00";
			this.txtComisionTarjeta.Location = new System.Drawing.Point(144, 416);
			this.txtComisionTarjeta.Name = "txtComisionTarjeta";
			this.txtComisionTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtComisionTarjeta.PromptChar = ' ';
			this.txtComisionTarjeta.Size = new System.Drawing.Size(112, 22);
			this.txtComisionTarjeta.TabIndex = 402;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(8, 395);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(72, 17);
			this.label12.TabIndex = 401;
			this.label12.Text = "(-) Protección";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtProteccion
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProteccion.Appearance = appearance24;
			this.txtProteccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProteccion.Enabled = false;
			this.txtProteccion.FormatString = "#,##0.00";
			this.txtProteccion.Location = new System.Drawing.Point(144, 392);
			this.txtProteccion.Name = "txtProteccion";
			this.txtProteccion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProteccion.PromptChar = ' ';
			this.txtProteccion.Size = new System.Drawing.Size(112, 22);
			this.txtProteccion.TabIndex = 400;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(384, 467);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(15, 17);
			this.label13.TabIndex = 406;
			this.label13.Text = "%";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(128, 467);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(10, 17);
			this.label14.TabIndex = 408;
			this.label14.Text = "$";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMargenMinimoValor
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenMinimoValor.Appearance = appearance25;
			this.txtMargenMinimoValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenMinimoValor.Enabled = false;
			this.txtMargenMinimoValor.FormatString = "#,##0.00";
			this.txtMargenMinimoValor.Location = new System.Drawing.Point(144, 464);
			this.txtMargenMinimoValor.Name = "txtMargenMinimoValor";
			this.txtMargenMinimoValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenMinimoValor.PromptChar = ' ';
			this.txtMargenMinimoValor.Size = new System.Drawing.Size(112, 22);
			this.txtMargenMinimoValor.TabIndex = 407;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(8, 299);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(123, 17);
			this.label15.TabIndex = 410;
			this.label15.Text = "(-) Precio Minimo Venta";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPrecioMinimoVenta
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioMinimoVenta.Appearance = appearance26;
			this.txtPrecioMinimoVenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioMinimoVenta.Enabled = false;
			this.txtPrecioMinimoVenta.FormatString = "#,##0.00";
			this.txtPrecioMinimoVenta.Location = new System.Drawing.Point(144, 296);
			this.txtPrecioMinimoVenta.Name = "txtPrecioMinimoVenta";
			this.txtPrecioMinimoVenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioMinimoVenta.PromptChar = ' ';
			this.txtPrecioMinimoVenta.Size = new System.Drawing.Size(112, 22);
			this.txtPrecioMinimoVenta.TabIndex = 409;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(8, 491);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(57, 17);
			this.label16.TabIndex = 412;
			this.label16.Text = "Descuento";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescuento
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescuento.Appearance = appearance27;
			this.txtDescuento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescuento.Enabled = false;
			this.txtDescuento.FormatString = "#,##0.00";
			this.txtDescuento.Location = new System.Drawing.Point(144, 488);
			this.txtDescuento.Name = "txtDescuento";
			this.txtDescuento.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescuento.PromptChar = ' ';
			this.txtDescuento.Size = new System.Drawing.Size(112, 22);
			this.txtDescuento.TabIndex = 411;
			// 
			// btnCalcular
			// 
			this.btnCalcular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCalcular.Image = ((System.Drawing.Image)(resources.GetObject("btnCalcular.Image")));
			this.btnCalcular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCalcular.Location = new System.Drawing.Point(272, 272);
			this.btnCalcular.Name = "btnCalcular";
			this.btnCalcular.Size = new System.Drawing.Size(72, 23);
			this.btnCalcular.TabIndex = 413;
			this.btnCalcular.Text = "Calcular";
			this.btnCalcular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCalcular.Click += new System.EventHandler(this.btnCalcular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(352, 272);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 414;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(272, 371);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(87, 17);
			this.label17.TabIndex = 416;
			this.label17.Text = "Costo Con I.V.A.";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCostoConIVA
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCostoConIVA.Appearance = appearance28;
			this.txtCostoConIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCostoConIVA.Enabled = false;
			this.txtCostoConIVA.FormatString = "#,##0.00";
			this.txtCostoConIVA.Location = new System.Drawing.Point(368, 368);
			this.txtCostoConIVA.Name = "txtCostoConIVA";
			this.txtCostoConIVA.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCostoConIVA.PromptChar = ' ';
			this.txtCostoConIVA.Size = new System.Drawing.Size(112, 22);
			this.txtCostoConIVA.TabIndex = 415;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 443);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(79, 17);
			this.label18.TabIndex = 418;
			this.label18.Text = "(=) Costo Final";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCostoFinal
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCostoFinal.Appearance = appearance29;
			this.txtCostoFinal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCostoFinal.Enabled = false;
			this.txtCostoFinal.FormatString = "#,##0.00";
			this.txtCostoFinal.Location = new System.Drawing.Point(144, 440);
			this.txtCostoFinal.Name = "txtCostoFinal";
			this.txtCostoFinal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCostoFinal.PromptChar = ' ';
			this.txtCostoFinal.Size = new System.Drawing.Size(112, 22);
			this.txtCostoFinal.TabIndex = 417;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(8, 320);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(25, 17);
			this.label19.TabIndex = 420;
			this.label19.Text = "(=) ";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCompetenciaPrecioMinimo
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCompetenciaPrecioMinimo.Appearance = appearance30;
			this.txtCompetenciaPrecioMinimo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCompetenciaPrecioMinimo.Enabled = false;
			this.txtCompetenciaPrecioMinimo.FormatString = "#,##0.00";
			this.txtCompetenciaPrecioMinimo.Location = new System.Drawing.Point(144, 320);
			this.txtCompetenciaPrecioMinimo.Name = "txtCompetenciaPrecioMinimo";
			this.txtCompetenciaPrecioMinimo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCompetenciaPrecioMinimo.PromptChar = ' ';
			this.txtCompetenciaPrecioMinimo.Size = new System.Drawing.Size(112, 22);
			this.txtCompetenciaPrecioMinimo.TabIndex = 419;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Location = new System.Drawing.Point(408, 467);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(148, 17);
			this.label20.TabIndex = 422;
			this.label20.Text = "Margen Igualar Competencia";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMargenIgualarCompetencia
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenIgualarCompetencia.Appearance = appearance31;
			this.txtMargenIgualarCompetencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenIgualarCompetencia.Enabled = false;
			this.txtMargenIgualarCompetencia.FormatString = "#,##0.00";
			this.txtMargenIgualarCompetencia.Location = new System.Drawing.Point(568, 464);
			this.txtMargenIgualarCompetencia.Name = "txtMargenIgualarCompetencia";
			this.txtMargenIgualarCompetencia.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenIgualarCompetencia.PromptChar = ' ';
			this.txtMargenIgualarCompetencia.Size = new System.Drawing.Size(112, 22);
			this.txtMargenIgualarCompetencia.TabIndex = 421;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.Location = new System.Drawing.Point(256, 491);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(15, 17);
			this.label21.TabIndex = 423;
			this.label21.Text = "%";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmFAC_DescuentosValidacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(692, 516);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.lblIce);
			this.Controls.Add(this.linkLabel1);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblPlazo);
			this.Controls.Add(this.lblBanco);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblPlanes);
			this.Controls.Add(this.lblTarjeta);
			this.Controls.Add(this.txtMargenIgualarCompetencia);
			this.Controls.Add(this.txtCompetenciaPrecioMinimo);
			this.Controls.Add(this.txtCostoFinal);
			this.Controls.Add(this.txtCostoConIVA);
			this.Controls.Add(this.btnCalcular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtDescuento);
			this.Controls.Add(this.txtPrecioMinimoVenta);
			this.Controls.Add(this.txtMargenMinimoValor);
			this.Controls.Add(this.txtMargenMinimoPorcentaje);
			this.Controls.Add(this.txtComisionTarjeta);
			this.Controls.Add(this.txtProteccion);
			this.Controls.Add(this.txtPrecio);
			this.Controls.Add(this.txtCosto);
			this.Controls.Add(this.txtPrecioCompetencia);
			this.Controls.Add(this.txtCompetencia);
			this.Controls.Add(this.cmbPlanesC);
			this.Controls.Add(this.cmbPlazosC);
			this.Controls.Add(this.cmbTarjetaC);
			this.Controls.Add(this.cmbBancoC);
			this.Controls.Add(this.optFormaPagoC);
			this.Controls.Add(this.cmbPlanes);
			this.Controls.Add(this.cmbPlazo);
			this.Controls.Add(this.cmbTarjeta);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.optFormaDePago);
			this.Controls.Add(this.txtidBanco);
			this.Controls.Add(this.txtidArticulo);
			this.Controls.Add(this.txtProducto);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbVendedor);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmFAC_DescuentosValidacion";
			this.Text = "Cotización De Descuentos";
			this.Load += new System.EventHandler(this.frmFAC_DescuentosValidacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProducto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFormaDePago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlanesC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPlazosC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjetaC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBancoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFormaPagoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCompetencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioCompetencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenMinimoPorcentaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComisionTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProteccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenMinimoValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioMinimoVenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCostoConIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCostoFinal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCompetenciaPrecioMinimo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenIgualarCompetencia)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmFAC_DescuentosValidacion_Load(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

			this.cmbBancoC.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Where Caja = 1 Order By Nombre");
			this.cmbBancoC.DisplayLayout.Bands[0].Columns["idBanco"].Hidden = true;
			this.cmbBancoC.DisplayLayout.Bands[0].Columns["Nombre"].Width = 224;

			this.cmbTarjetaC.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTarjeta, Nombre From TarjetasDeCredito Order By Nombre");
			this.cmbTarjetaC.DisplayLayout.Bands[0].Columns["idTarjeta"].Hidden = true;
			this.cmbTarjetaC.DisplayLayout.Bands[0].Columns["Nombre"].Width = 224;

			this.cmbPlanesC.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlan, Descripcion From Planes");
			this.cmbPlanesC.DisplayLayout.Bands[0].Columns["idPlan"].Hidden = true;
			this.cmbPlanesC.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 224;

			this.cmbPlazosC.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlazo, Descripcion From Plazos Order By Descripcion");
			this.cmbPlazosC.DisplayLayout.Bands[0].Columns["idPlazo"].Hidden = true;
			this.cmbPlazosC.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 224;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 1, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
		}
		
		int iRegion = 0;

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

			iRegion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Region From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));

			string sSQL = string.Format("Exec BTPP 0, 0, 0, 3, {0}, '{1}'", iRegion, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void txtProducto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				int iBodega = 0;
				if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

				using (BuscaArticulos miBusqueda = new BuscaArticulos(this.txtProducto.Text.ToString(), iBodega, 1, 0, DateTime.Today, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.txtProducto.Text = miBusqueda.sCodigo;
						this.txtidArticulo.Value = (int)miBusqueda.idArticulo;
						this.txtPrecioMinimoVenta.Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Precio3 From Articulo Where idArticulo = {0}", (int)this.txtidArticulo.Value));
						
					}
				}
			}
		}

		private void cmbBanco_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null)
			{
				this.cmbTarjeta.Value = System.DBNull.Value;
				this.cmbPlanes.Value = System.DBNull.Value;
				this.cmbPlazo.Value = System.DBNull.Value;
				//@idBanco Int, @idTarjeta Int, @idPlan Int, @Tipo Int, @Bodega Int = 0
				string sSQL = string.Format("Exec BTPP {0}, 0, 0, 0, {1}, '{2}', {3}", 
					(int)this.cmbBanco.Value, iRegion, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 3);
				this.cmbTarjeta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				this.cmbTarjeta.DisplayLayout.Bands[0].Columns["idTarjeta"].Hidden = true;
				this.cmbTarjeta.DisplayLayout.Bands[0].Columns["Nombre"].Width = 224;
			}
		}

		private void cmbTarjeta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null)
			{
				if (this.cmbTarjeta.ActiveRow != null) 
				{
					this.cmbPlanes.Value = System.DBNull.Value;
					this.cmbPlazo.Value = System.DBNull.Value;

					string sSQL = string.Format("Exec BTPP {0}, {1}, 0, 1, {2}, '{3}', {4}", 
						(int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value, iRegion, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 3);
					this.cmbPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

					this.cmbPlanes.DisplayLayout.Bands[0].Columns["idPlan"].Hidden = true;
					this.cmbPlanes.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 224;
				}
			}
		}

		private void cmbPlanes_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow != null)
			{
				if (this.cmbTarjeta.ActiveRow != null)
				{
					if (this.cmbPlanes.ActiveRow != null) 
					{
						this.cmbPlazo.Value = System.DBNull.Value;

						string sSQL = string.Format("Exec BTPP {0}, {1}, {2}, 2, {3}, '{4}', {5}", 
							(int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value, (int)this.cmbPlanes.Value, iRegion, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 3);
						this.cmbPlazo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

						this.cmbPlazo.DisplayLayout.Bands[0].Columns["idPlazo"].Hidden = true;
						this.cmbPlazo.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 224;

						if (this.cmbPlazo.Rows.Count == 1)
						{
							int idPlazo = 0;
							idPlazo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idPlazo From PlanesPlazos Where idPlan = {0} And idTarjeta = {1} And idBanco = {2}", 
								(int)this.cmbPlanes.Value, (int)this.cmbTarjeta.Value, (int)this.cmbBanco.Value)); 
							if (idPlazo > 0) this.cmbPlazo.Value = idPlazo;
						}
					}
				}
			}
		}

		private void optFormaDePago_ValueChanged(object sender, System.EventArgs e)
		{
			this.cmbBanco.Enabled = false;
			this.cmbTarjeta.Enabled = false;
			this.cmbPlanes.Enabled = false;
			this.cmbPlazo.Enabled = false;

			this.cmbBanco.Value = System.DBNull.Value;
			this.cmbTarjeta.Value = System.DBNull.Value;
			this.cmbPlanes.Value = System.DBNull.Value;
			this.cmbPlazo.Value = System.DBNull.Value;

			int idBanco = 0;
			int idTarjeta = 0;
			int idPlan = 0;
			int	idPlazo = 0;
			dComisionPorc = 0.00m;
			this.txtPrecio.Value = 0.00m;
			this.txtCostoFinal.Value = 0.00m;
			this.txtCosto.Value = 0.00m;
			this.txtCostoConIVA.Value = 0.00m;
			this.txtProteccion.Value = 0.00m;			
			this.txtComisionTarjeta.Value = 0.00m;
			this.txtMargenMinimoPorcentaje.Value = 0.00m;
			this.txtMargenMinimoValor.Value = 0.00m;
			this.txtPrecioMinimoVenta.Value = 0.00m;
			this.txtCompetenciaPrecioMinimo.Value = 0.00m;
			this.txtMargenIgualarCompetencia.Value = 0.00m;
			this.txtDescuento.Value = 0.00m;

			if ((int)this.optFormaDePago.Value != 1)
			{
				this.cmbBanco.Enabled = true;
				this.cmbTarjeta.Enabled = true;
				this.cmbPlanes.Enabled = true;
				this.cmbPlazo.Enabled = true;
			}
		}

		private void optFormaPagoC_ValueChanged(object sender, System.EventArgs e)
		{
			this.cmbBancoC.Enabled = false;
			this.cmbTarjetaC.Enabled = false;
			this.cmbPlanesC.Enabled = false;
			this.cmbPlazosC.Enabled = false;

			this.cmbBancoC.Value = System.DBNull.Value;
			this.cmbTarjetaC.Value = System.DBNull.Value;
			this.cmbPlanesC.Value = System.DBNull.Value;
			this.cmbPlazosC.Value = System.DBNull.Value;

			if ((int)this.optFormaPagoC.Value != 1)
			{
				this.cmbBancoC.Enabled = true;
				this.cmbTarjetaC.Enabled = true;
				this.cmbPlanesC.Enabled = true;
				this.cmbPlazosC.Enabled = true;
			}
		}
		decimal dComisionPorc = 0.00m;

		private void btnCalcular_Click(object sender, System.EventArgs e)
		{
			int idBanco = 0;
			int idTarjeta = 0;
			int idPlan = 0;
			int	idPlazo = 0;
			dComisionPorc = 0.00m;
			this.txtPrecio.Value = 0.00m;
			this.txtCostoFinal.Value = 0.00m;
			this.txtCosto.Value = 0.00m;
			this.txtCostoConIVA.Value = 0.00m;
			this.txtProteccion.Value = 0.00m;			
			this.txtComisionTarjeta.Value = 0.00m;
			this.txtMargenMinimoPorcentaje.Value = 0.00m;
			this.txtMargenMinimoValor.Value = 0.00m;
			this.txtPrecioMinimoVenta.Value = 0.00m;
			this.txtCompetenciaPrecioMinimo.Value = 0.00m;
			this.txtMargenIgualarCompetencia.Value = 0.00m;
			this.txtDescuento.Value = 0.00m;

			if (this.cmbBanco.ActiveRow != null) idBanco = (int)this.cmbBanco.Value;
			if (this.cmbTarjeta.ActiveRow != null) idTarjeta = (int)this.cmbTarjeta.Value;
			if (this.cmbPlanes.ActiveRow != null) idPlan = (int)this.cmbPlanes.Value;
			if (this.cmbPlazo.ActiveRow != null) idPlazo = (int)this.cmbPlazo.Value;
			
			string sSQL = string.Format("Exec FAC_DescuentosCotizador {0}, {1}, {2}, {3}, {4}, {5}", 
				(int)this.txtidArticulo.Value, (int)this.optFormaDePago.Value, idBanco, idTarjeta, idPlan, idPlazo);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtPrecio.Value = dr.GetDecimal(0);
				this.txtCosto.Value = dr.GetDecimal(1);
				this.txtCostoConIVA.Value = dr.GetDecimal(2);
				this.txtProteccion.Value = dr.GetDecimal(3);
				dComisionPorc = dr.GetDecimal(4);
				this.txtComisionTarjeta.Value = dr.GetDecimal(5);
				this.txtMargenMinimoPorcentaje.Value = dr.GetDecimal(6);
			}
			dr.Close();

			this.txtCostoFinal.Value = Convert.ToDecimal(this.txtCosto.Value) -
        Convert.ToDecimal(this.txtProteccion.Value) +
				Convert.ToDecimal(this.txtComisionTarjeta.Value);

			//decimal dMargen = Convert.ToDecimal(this.txtCostoFinal.Value) / 0.86m;
			this.txtPrecioMinimoVenta.Value = Convert.ToDecimal(this.txtCostoFinal.Value) / 0.86m; //Convert.ToDecimal(this.txtCostoFinal.Value) + dMargen;// + Convert.ToDecimal(this.txtMargenMinimoValor.Value);
			this.txtMargenMinimoValor.Value = Convert.ToDecimal(this.txtPrecioMinimoVenta.Value) - Convert.ToDecimal(this.txtCostoFinal.Value);// * Convert.ToDecimal(this.txtMargenMinimoPorcentaje.Value)) / 100.00m;
			
			this.txtCompetenciaPrecioMinimo.Value = Convert.ToDecimal(this.txtPrecioCompetencia.Value) - Convert.ToDecimal(this.txtPrecioMinimoVenta.Value);
			this.txtMargenIgualarCompetencia.Value = ((Convert.ToDecimal(this.txtPrecioCompetencia.Value) - Convert.ToDecimal(this.txtCostoFinal.Value)) / Convert.ToDecimal(this.txtPrecioCompetencia.Value)) * 100.00m;
			this.txtDescuento.Value = (Convert.ToDecimal(this.txtPrecio.Value) - Convert.ToDecimal(this.txtPrecioCompetencia.Value)) / Convert.ToDecimal(this.txtPrecio.Value) * 100.00m;
		}
	}
}
