using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmBuscaNCPendientes.
	/// </summary>
	public class frmBuscaNCPendientes : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnBuscar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;

		int IdCliente = 0;
		int idTipoFactura = 0;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoFacturas;
		private System.Windows.Forms.Label lblFacturas;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridNotasDeCredito;
		private System.Windows.Forms.GroupBox groupBox2;
		decimal DSaldoFac = 0.00m;

		public frmBuscaNCPendientes(int idCliente, decimal dSaldoFac, int IdTipoFactura)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			IdCliente = idCliente;
			DSaldoFac = dSaldoFac;
			idTipoFactura = IdTipoFactura;
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
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizado");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoAnticipo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 10, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label6 = new System.Windows.Forms.Label();
			this.txtSaldoFacturas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblFacturas = new System.Windows.Forms.Label();
			this.uGridNotasDeCredito = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoFacturas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridNotasDeCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.Location = new System.Drawing.Point(8, 285);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.TabIndex = 3;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Location = new System.Drawing.Point(88, 285);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 4;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Location = new System.Drawing.Point(280, 4);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 1;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// txtBusqueda
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance1;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(72, 5);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(184, 21);
			this.txtBusqueda.TabIndex = 0;
			this.txtBusqueda.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBusqueda_KeyPress);
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.ultraLabel4.Location = new System.Drawing.Point(8, 8);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(37, 15);
			this.ultraLabel4.TabIndex = 259;
			this.ultraLabel4.Text = "Buscar";
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
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 246);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(850, 8);
			this.groupBox1.TabIndex = 264;
			this.groupBox1.TabStop = false;
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(522, 264);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(166, 16);
			this.label6.TabIndex = 635;
			this.label6.Text = "Valor Seleccionado En Facturas";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldoFacturas
			// 
			this.txtSaldoFacturas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoFacturas.Appearance = appearance2;
			this.txtSaldoFacturas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoFacturas.Enabled = false;
			this.txtSaldoFacturas.FormatString = "#,##0.00";
			this.txtSaldoFacturas.Location = new System.Drawing.Point(738, 262);
			this.txtSaldoFacturas.Name = "txtSaldoFacturas";
			this.txtSaldoFacturas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoFacturas.PromptChar = ' ';
			this.txtSaldoFacturas.Size = new System.Drawing.Size(96, 21);
			this.txtSaldoFacturas.TabIndex = 636;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(522, 288);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(207, 16);
			this.label5.TabIndex = 633;
			this.label5.Text = "Valor Seleccionado En Notas de Crédito";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValor
			// 
			this.txtValor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance3;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(738, 286);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(96, 21);
			this.txtValor.TabIndex = 634;
			// 
			// lblFacturas
			// 
			this.lblFacturas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblFacturas.AutoSize = true;
			this.lblFacturas.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFacturas.ForeColor = System.Drawing.Color.Firebrick;
			this.lblFacturas.Location = new System.Drawing.Point(626, 6);
			this.lblFacturas.Name = "lblFacturas";
			this.lblFacturas.Size = new System.Drawing.Size(0, 18);
			this.lblFacturas.TabIndex = 637;
			this.lblFacturas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridNotasDeCredito
			// 
			this.uGridNotasDeCredito.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridNotasDeCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridNotasDeCredito.DataSource = this.ultraDataSource1;
			appearance4.BackColor = System.Drawing.Color.White;
			this.uGridNotasDeCredito.DisplayLayout.Appearance = appearance4;
			this.uGridNotasDeCredito.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 20;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 44;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 75;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Número";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "R.U.C.";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 110;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Proveedor";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 260;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.Caption = "Emisión";
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance6;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Saldo";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance7;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Width = 80;
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
																										 ultraGridColumn11});
			ultraGridBand1.Header.Caption = "Documentos";
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance8;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridNotasDeCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridNotasDeCredito.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.uGridNotasDeCredito.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridNotasDeCredito.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridNotasDeCredito.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridNotasDeCredito.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridNotasDeCredito.DisplayLayout.Override.CellAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8.5F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridNotasDeCredito.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridNotasDeCredito.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridNotasDeCredito.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridNotasDeCredito.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridNotasDeCredito.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridNotasDeCredito.Location = new System.Drawing.Point(9, 56);
			this.uGridNotasDeCredito.Name = "uGridNotasDeCredito";
			this.uGridNotasDeCredito.Size = new System.Drawing.Size(826, 184);
			this.uGridNotasDeCredito.TabIndex = 2;
			this.uGridNotasDeCredito.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridNotasDeCredito_CellChange);
			this.uGridNotasDeCredito.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridNotasDeCredito_AfterCellUpdate);
			this.uGridNotasDeCredito.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridNotasDeCredito_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn8.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
																																 ultraDataColumn11});
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(936, 8);
			this.groupBox2.TabIndex = 638;
			this.groupBox2.TabStop = false;
			// 
			// frmBuscaNCPendientes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(844, 318);
			this.ControlBox = false;
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.uGridNotasDeCredito);
			this.Controls.Add(this.lblFacturas);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtSaldoFacturas);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.txtBusqueda);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmBuscaNCPendientes";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Notas de crédito pendientes de cancelar";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmBuscaNCPendientes_KeyDown);
			this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frmBuscaNCPendientes_KeyPress);
			this.Load += new System.EventHandler(this.frmBuscaNCPendientes_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoFacturas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridNotasDeCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmBuscaNCPendientes_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.dsGeneralModulo(string.Format("Exec ConsultaNCProvCXP '{0}', {1}, {2}", this.txtBusqueda.Text.ToString(), IdCliente, idTipoFactura), this.uGridNotasDeCredito);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmBuscaNCPendientes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void txtBusqueda_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnBuscar_Click(sender, e);
		}

		private void Suma()
		{
			this.txtValor.Value = 0;
						
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridNotasDeCredito.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value && dr.Cells["idCompra"].Value != System.DBNull.Value)
				{
					this.txtValor.Value = Convert.ToDecimal(this.txtValor.Value) + Convert.ToDecimal(dr.Cells["SaldoAnticipo"].Value);
				}
			}
		}

		private void CuentaFacSel()
		{
			this.lblFacturas.Text = "";
			int iCont = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridNotasDeCredito.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value) iCont ++;
			}

			this.lblFacturas.Text = iCont.ToString() + " DOCUMENTOS SELECCIONADOS.";
		}

		private void frmBuscaNCPendientes_Load(object sender, System.EventArgs e)
		{
			this.txtBusqueda.Focus();

			this.btnBuscar_Click(sender, e);

			this.txtSaldoFacturas.Value = DSaldoFac;
		}

		private void Aceptar()
		{
			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridNotasDeCredito.Rows.All)
				if ((bool) dr.Cells["Sel"].Value) iCont ++;

			if (iCont == 0)
			{
				MessageBox.Show("Seleccione la(s) Factura(s).", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridNotasDeCredito.Focus();
				return;
			}

			this.DialogResult = DialogResult.OK;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void uGridNotasDeCredito_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void uGridNotasDeCredito_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridNotasDeCredito_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCompra"].Value == System.DBNull.Value) return;

			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridNotasDeCredito.UpdateData();

				Suma();

				CuentaFacSel();
			}
		}
	}
}
