using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPrendasPago.
	/// </summary>
	public class frmPrendasPago : System.Windows.Forms.Form
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinGrid.UltraGrid grdPrendas;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPrendasPago()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idformaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorFactura");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NuevoValor");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoPrenda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPrenda");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago_1", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("N", 1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPrendasPago));
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idformaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NuevoValor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoPrenda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorPrenda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.grdPrendas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdPrendas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel1.Location = new System.Drawing.Point(8, 11);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 15);
			this.ultraLabel1.TabIndex = 270;
			this.ultraLabel1.Text = "Buscar";
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(272, 7);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 269;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// txtBusqueda
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance1;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(64, 8);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(192, 21);
			this.txtBusqueda.TabIndex = 268;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-32, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1288, 8);
			this.groupBox1.TabIndex = 271;
			this.groupBox1.TabStop = false;
			// 
			// grdPrendas
			// 
			this.grdPrendas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdPrendas.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdPrendas.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdPrendas.DisplayLayout.Appearance = appearance2;
			this.grdPrendas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 70;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 95;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Width = 104;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Header.Caption = "Fecha Factura";
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Width = 116;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "Total Factura";
			ultraGridColumn5.Header.VisiblePosition = 8;
			ultraGridColumn5.Width = 115;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 61;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Cliente";
			ultraGridColumn7.Header.VisiblePosition = 11;
			ultraGridColumn7.Width = 159;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 12;
			ultraGridColumn8.Width = 107;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance4;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Nuevo Valor";
			ultraGridColumn9.Header.VisiblePosition = 13;
			ultraGridColumn9.Width = 89;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 136;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 76;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance5;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Valor Prenda";
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Width = 94;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn13.Header.Caption = "Forma De Pago";
			ultraGridColumn13.Header.VisiblePosition = 15;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn13.Width = 209;
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 56;
			ultraGridColumn15.DataType = typeof(bool);
			ultraGridColumn15.Header.Caption = "...";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Width = 27;
			ultraGridColumn16.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance6.Image = ((object)(resources.GetObject("appearance6.Image")));
			appearance6.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellButtonAppearance = appearance6;
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn16.Width = 31;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
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
																										 ultraGridColumn16});
			ultraGridBand1.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.grdPrendas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdPrendas.DisplayLayout.Override.ActiveRowAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.grdPrendas.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 9F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdPrendas.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.grdPrendas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdPrendas.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdPrendas.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdPrendas.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.grdPrendas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdPrendas.Location = new System.Drawing.Point(8, 56);
			this.grdPrendas.Name = "grdPrendas";
			this.grdPrendas.Size = new System.Drawing.Size(1208, 248);
			this.grdPrendas.TabIndex = 272;
			this.grdPrendas.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdPrendas_CellChange);
			this.grdPrendas.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdPrendas_ClickCellButton);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.Enabled = false;
			this.btnAceptar.Location = new System.Drawing.Point(8, 312);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 273;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Location = new System.Drawing.Point(88, 312);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 274;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource2;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "Descripcion";
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(1008, 128);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(200, 88);
			this.cmbFormaPago.TabIndex = 733;
			this.cmbFormaPago.ValueMember = "idFormaPago_1";
			this.cmbFormaPago.Visible = false;
			// 
			// ultraDataSource2
			// 
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// frmPrendasPago
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1224, 342);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.grdPrendas);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.txtBusqueda);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmPrendasPago";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Prendas Pendientes";
			this.Load += new System.EventHandler(this.frmPrendasPago_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdPrendas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		int iTarjeta = 0;
		int iBanco = 0;
		#endregion Variables

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmPrendasPago_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "1126");

			if (!Funcion.Permiso("1798", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Prendas Pendientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}
			if (miAcceso.Nuevo) this.btnAceptar.Enabled = true;
			
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idFormaPago As idFormaPago_1, FormaPago As Descripcion From CompraFormaPago Where idFormaPago in (1, 3)"));	
			this.grdPrendas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaPrendasPendientes"));
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ConsultaPrendasPendientes '{0}'", this.txtBusqueda.Text.ToString());
			this.grdPrendas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void grdPrendas_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Check
			if (e.Cell.Row.Cells["Sel"].Value != System.DBNull.Value)
			{
				if (e.Cell.Column.ToString() == "Sel")
				{
					this.grdPrendas.UpdateData();

					if ((int)this.grdPrendas.Rows.Count > 1)
					{
						int iCompra = 0;
						int icount = 0;
						iCompra = (int)e.Cell.Row.Cells["idCompra"].Value;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdPrendas.Rows.All)
						{								
							if ((bool)dr.Cells["Sel"].Value)
							{
								icount++;
							}
						}
						if (icount > 1)
						{
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.grdPrendas.Rows.All)
							{	
								if ((bool)dr.Cells["Sel"].Value)
								{
									if(iCompra != (int)dr.Cells["idCompra"].Value)
									{
										dr.Cells["Sel"].Value = false;
									}
								}
							}
						}
					}
				}
			}
			#endregion Check
		}

		private void Aceptar()
		{
			#region Valores
			int iFormaPago = 0;
			int iTurno = 0;
			int iBodega = 0;
			int iCliente = 0;
			int iTipoAnticipo = 0;
			int iAnticipo = 0;
			int iOrigen = 0;
			int iCuenta = 1211;
//			int iPlan = 0;
//			int iPlazo = 0;
//			int iAsiento = 0;
			int iCompra = 0;
			decimal dValor = 0.00m;
			decimal dValorPrenda = 0.00m;
			DateTime dFechaRT = DateTime.Today;
			DateTime dtFecha = DateTime.Today;
			bool bCheck = false;
			string sSecuencia = "";
			string sMensaje = "";
			string sDocUniBanco = "";
			string sCuentaCorriente = "";
			string sNumero = "";
			string sAutorizacion = "";
			string sNumVoucher = "";
			string sLote = "";
			string sReferencia = "";

			iTipoAnticipo = 1;
			iOrigen = 2;
			iFormaPago = (int)grdPrendas.ActiveRow.Cells["idFormaPago_1"].Value;
			iBodega = (int)grdPrendas.ActiveRow.Cells["Bodega"].Value;
			iCompra = (int)grdPrendas.ActiveRow.Cells["idCompra"].Value;
			iCliente = (int)grdPrendas.ActiveRow.Cells["idCliente"].Value;
			dValor = (decimal)grdPrendas.ActiveRow.Cells["NuevoValor"].Value;
			bCheck = (bool)grdPrendas.ActiveRow.Cells["Sel"].Value;
			dValorPrenda = (decimal)grdPrendas.ActiveRow.Cells["ValorPrenda"].Value;
			dtFecha = Funcion.dtEscalarSQL(cdsSeteoF, "Select dbo.FacturacionRetornaFecha()");  //FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
			sMensaje = "CONTABILIZACION PRENDA";

			dFechaRT = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
			iTurno = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaTurno '{0}'", Convert.ToDateTime(dFechaRT).ToString("yyyyMMdd HH:mm")));
			#endregion Valores
			
			if (bCheck)
			{
				if (dValor == 0)
				{
					MessageBox.Show("Ingrese el nuevo valor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.grdPrendas.Focus();				
					return;
				}
				else if (dValor < 0)
				{
					MessageBox.Show("No puede ingresar valores en negativo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.grdPrendas.Focus();				
					return;
				}
				else if(dValorPrenda != dValor)
				{
					MessageBox.Show("El valor ingresado es distinto al valor prenda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.grdPrendas.Focus();				
					return;
				}
				else if (iFormaPago == 0)
				{
					MessageBox.Show("Seleccione la Forma de Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.grdPrendas.Focus();
					return;
				}
				else
				{
					using (VentaFPagoLot miCobro = new VentaFPagoLot (iTipoAnticipo, iFormaPago, dValor, iOrigen, 4))
					{
						if (DialogResult.OK == miCobro.ShowDialog())
						{
							#region Proceso
							using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))					
							{
								oConexion.Open();

								SqlTransaction oTransaction = null;
								SqlCommand oCmdActualiza = oConexion.CreateCommand();
								oCmdActualiza.Connection = oConexion;
								oCmdActualiza.CommandType = CommandType.Text;
								oCmdActualiza.CommandTimeout = 120;
						
								try
								{							
									oTransaction = oConexion.BeginTransaction();
									oCmdActualiza.Transaction = oTransaction;

									#region Numeracion
									string sSQLNumero = string.Format("Exec NumeracionLocales 24, {0}, 0", iBodega);
									oCmdActualiza.CommandText = sSQLNumero;
									sSecuencia = oCmdActualiza.ExecuteScalar().ToString();
									#endregion Numeracion

									#region Guarda Anticipo
									string sSQLGraba = string.Format("Exec GuardaAnticipos {0}, {1}, '{2}', {3}, '{4}', {5}, {6}, {7}, '{8}', {9}, {10}, {11}, '{12}', {13}, '{14}', '{15}', {16}", 
										0, 0, dtFecha.ToString("yyyyMMdd"), iBodega, sSecuencia, iCliente, iFormaPago, 0, 
										sMensaje, 3, Convert.ToDecimal(dValor), iTipoAnticipo, sDocUniBanco,
										0, 0, dtFecha.ToString("yyyyMMdd"), iTurno);
									oCmdActualiza.CommandText = sSQLGraba;
									iAnticipo = (int)oCmdActualiza.ExecuteScalar();
									#endregion Guarda Anticipo

									#region Guarda Cobro
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in miCobro.ultraGrid1.Rows.All)
									{
										int idDetFormaPago = 0;
										int idBanco = 0;
										int idTarjeta = 0;
										int idCuenta = 0;
										int idPlan = 0;
										int idPlazo = 0;
										decimal dComTarjeta = 0.00m;
										DateTime dtFecCaducidad = DateTime.Today;
							
										if (dr.Cells["idDetFormaPago"].Value != System.DBNull.Value) idDetFormaPago = (int)dr.Cells["idDetFormaPago"].Value;
										if (dr.Cells["idBanco"].Value != System.DBNull.Value) idBanco = (int)dr.Cells["idBanco"].Value;
										if (dr.Cells["idTarjeta"].Value != System.DBNull.Value) idTarjeta = (int)dr.Cells["idTarjeta"].Value;
										if (dr.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)dr.Cells["idCuenta"].Value;
										if (dr.Cells["idPlan"].Value != System.DBNull.Value) idPlan = (int)dr.Cells["idPlan"].Value;							
										if (dr.Cells["idPlazo"].Value != System.DBNull.Value) idPlazo = (int)dr.Cells["idPlazo"].Value;
										if ((int)dr.Cells["idFormaPago"].Value == 3 || (int)dr.Cells["idFormaPago"].Value == 11) if (dr.Cells["FechaCaducidad"].Value != System.DBNull.Value) dtFecCaducidad = (DateTime)dr.Cells["FechaCaducidad"].Value;
										if (dr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = (decimal )dr.Cells["ValorUsoTarjeta"].Value;

										string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, 0", 
											idDetFormaPago, 
											iAnticipo, iOrigen,
											(int)dr.Cells["idFormaPago"].Value, (decimal )dr.Cells["Valor"].Value,
											idBanco, idTarjeta,
											dr.Cells["CuentaCorriente"].Value.ToString(), dr.Cells["Numero"].Value.ToString(),
											dtFecCaducidad.ToString("yyyyMMdd"),								
											dr.Cells["NumVoucher"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString(),
											idCuenta, 
											dr.Cells["Lote"].Value.ToString(), dr.Cells["Referencia"].Value.ToString(),
											idPlan, idPlazo, dComTarjeta);
										oCmdActualiza.CommandText = sSQLCobro;
										oCmdActualiza.ExecuteNonQuery();
									}
									#endregion Guarda Cobro

									#region Saldo
									string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 24, 1", iAnticipo);
									oCmdActualiza.CommandText = sSQLSaldo;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Saldo

									#region Saldo Anticipo
									string sSQLA = string.Format("Exec ActualizarAnticipos {0}, {1}, {2}, '{3}', {4}", iCliente, dValor, iCompra, sSecuencia, iBodega);
									oCmdActualiza.CommandText = sSQLA;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Saldo Anticipo

									#region Actualiza Estado Prenda
									string sSQLAE = string.Format("Exec ActulizaEstadoPrenda {0}", iCompra);
									oCmdActualiza.CommandText = sSQLAE;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Actualiza Estado Prenda

									#region Asiento
									//							string sSQLAsiento = string.Format("Exec CreaAsientoAnticipo {0}, {1}, {2}, {3}", iAsiento, iAnticipo, 1, iOrigen);
									//							iAsiento = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento, true);							
									#endregion Asiento

									oTransaction.Commit();
							
									oConexion.Close();
								
									#region Impresion
									using (frmNumeracionRecibo Numeracion = new frmNumeracionRecibo (sSecuencia))
									{
										if (DialogResult.OK == Numeracion.ShowDialog())
										{
											Facturacion.ImpresionRecibos(iAnticipo, 1, cdsSeteoF);
										}
									}						
									#endregion Impresion																
								}
								catch (Exception ex)
								{
									oTransaction.Rollback();

									MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
								}
								finally
								{
									oConexion.Close();
								}
							}
							#endregion Proceso
						}
					}
				}
			}
			else
			{
				MessageBox.Show("Seleccione el documento para generar el pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.grdPrendas.Focus();				
				return;
			}
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			Aceptar();
			this.btnBuscar_Click(sender, e);
		}

		private void grdPrendas_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "N")
			{
				using (frmDescripcionPrenda Busqueda = new frmDescripcionPrenda((int)e.Cell.Row.Cells["idCompra"].Value ))
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{						
					}
				}
			}
		}
	}
}
