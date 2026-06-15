using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;
using System.Data.SqlClient;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ImportacionProgramacion.
	/// </summary>
	public class ImportacionProgramacion : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btProgramacion;
		private System.Windows.Forms.Button btFlujo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPago;
		private Infragistics.Win.Misc.UltraLabel ultraLabel17;
		private Infragistics.Win.Misc.UltraLabel ultraLabel13;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCtaSwift;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBeneficiario;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje;
		private Infragistics.Win.Misc.UltraLabel Porcentaje;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ImportacionProgramacion()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacionTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idcompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("programado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("autorizado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("saldopro");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btProgramacion = new System.Windows.Forms.Button();
			this.btFlujo = new System.Windows.Forms.Button();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel17 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel13 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.txtCtaSwift = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.txtBeneficiario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.txtPorcentaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.Porcentaje = new Infragistics.Win.Misc.UltraLabel();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCtaSwift)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBeneficiario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 61;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 154;
			ultraGridColumn3.Header.Caption = "Número";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 109;
			ultraGridColumn4.Header.Caption = "Importación";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 266;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 115;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance2;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 109;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 252;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(872, 320);
			this.ultraGrid1.TabIndex = 163;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			this.ultraGrid1.Click += new System.EventHandler(this.ultraGrid1_Click);
			// 
			// ultraDataSource2
			// 
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// ultraDataSource1
			// 
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15});
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
			// 
			// btProgramacion
			// 
			this.btProgramacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btProgramacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btProgramacion.Location = new System.Drawing.Point(448, 336);
			this.btProgramacion.Name = "btProgramacion";
			this.btProgramacion.Size = new System.Drawing.Size(104, 24);
			this.btProgramacion.TabIndex = 6;
			this.btProgramacion.Text = "Programar Pago";
			this.btProgramacion.Click += new System.EventHandler(this.btProgramacion_Click);
			// 
			// btFlujo
			// 
			this.btFlujo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btFlujo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btFlujo.Location = new System.Drawing.Point(448, 368);
			this.btFlujo.Name = "btFlujo";
			this.btFlujo.Size = new System.Drawing.Size(104, 24);
			this.btFlujo.TabIndex = 7;
			this.btFlujo.Text = "Flujo de Pago";
			this.btFlujo.Click += new System.EventHandler(this.btFlujo_Click);
			// 
			// txtPago
			// 
			this.txtPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtPago.Location = new System.Drawing.Point(304, 336);
			this.txtPago.Name = "txtPago";
			this.txtPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPago.PromptChar = ' ';
			this.txtPago.Size = new System.Drawing.Size(104, 21);
			this.txtPago.TabIndex = 1;
			this.txtPago.Click += new System.EventHandler(this.txtPago_Click);
			this.txtPago.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPago_KeyPress);
			this.txtPago.Enter += new System.EventHandler(this.txtPago_Enter);
			// 
			// ultraLabel17
			// 
			this.ultraLabel17.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel17.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel17.Location = new System.Drawing.Point(248, 336);
			this.ultraLabel17.Name = "ultraLabel17";
			this.ultraLabel17.Size = new System.Drawing.Size(48, 17);
			this.ultraLabel17.TabIndex = 199;
			this.ultraLabel17.Text = "Pago";
			// 
			// ultraLabel13
			// 
			this.ultraLabel13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel13.AutoSize = true;
			this.ultraLabel13.Location = new System.Drawing.Point(8, 358);
			this.ultraLabel13.Name = "ultraLabel13";
			this.ultraLabel13.Size = new System.Drawing.Size(36, 14);
			this.ultraLabel13.TabIndex = 202;
			this.ultraLabel13.Text = "Fecha";
			// 
			// cmbBanco
			// 
			this.cmbBanco.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DisplayLayout.AutoFitColumns = true;
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(112, 384);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(296, 21);
			this.cmbBanco.TabIndex = 3;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBanco_KeyPress);
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel7.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel7.Location = new System.Drawing.Point(8, 380);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(48, 17);
			this.ultraLabel7.TabIndex = 203;
			this.ultraLabel7.Text = "Banco";
			// 
			// txtCtaSwift
			// 
			this.txtCtaSwift.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtCtaSwift.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCtaSwift.Location = new System.Drawing.Point(112, 408);
			this.txtCtaSwift.MaxLength = 20;
			this.txtCtaSwift.Name = "txtCtaSwift";
			this.txtCtaSwift.Size = new System.Drawing.Size(296, 21);
			this.txtCtaSwift.TabIndex = 4;
			this.txtCtaSwift.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCtaSwift_KeyPress);
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel5.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel5.Location = new System.Drawing.Point(8, 405);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(72, 17);
			this.ultraLabel5.TabIndex = 206;
			this.ultraLabel5.Text = "Cuenta Swift";
			// 
			// txtBeneficiario
			// 
			this.txtBeneficiario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtBeneficiario.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBeneficiario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBeneficiario.Location = new System.Drawing.Point(112, 432);
			this.txtBeneficiario.Name = "txtBeneficiario";
			this.txtBeneficiario.Size = new System.Drawing.Size(296, 21);
			this.txtBeneficiario.TabIndex = 5;
			this.txtBeneficiario.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBeneficiario_KeyPress);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Location = new System.Drawing.Point(8, 430);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(72, 17);
			this.ultraLabel1.TabIndex = 208;
			this.ultraLabel1.Text = "Beneficiario";
			// 
			// cmbFecha
			// 
			this.cmbFecha.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbFecha.DateTime = new System.DateTime(2013, 7, 25, 0, 0, 0, 0);
			this.cmbFecha.FormatString = "dd/MM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(112, 360);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.PromptChar = ' ';
			this.cmbFecha.Size = new System.Drawing.Size(160, 21);
			this.cmbFecha.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbFecha.TabIndex = 2;
			this.cmbFecha.Value = new System.DateTime(2013, 7, 25, 0, 0, 0, 0);
			this.cmbFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbFecha_KeyPress);
			// 
			// txtPorcentaje
			// 
			this.txtPorcentaje.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtPorcentaje.FormatString = "#,##0.00";
			this.txtPorcentaje.Location = new System.Drawing.Point(112, 336);
			this.txtPorcentaje.Name = "txtPorcentaje";
			this.txtPorcentaje.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje.PromptChar = ' ';
			this.txtPorcentaje.Size = new System.Drawing.Size(104, 21);
			this.txtPorcentaje.TabIndex = 0;
			this.txtPorcentaje.Click += new System.EventHandler(this.txtPorcentaje_Click);
			this.txtPorcentaje.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPorcentaje_KeyPress);
			this.txtPorcentaje.Validated += new System.EventHandler(this.txtPorcentaje_Validated);
			this.txtPorcentaje.ValueChanged += new System.EventHandler(this.txtPorcentaje_ValueChanged);
			this.txtPorcentaje.Enter += new System.EventHandler(this.txtPorcentaje_Enter);
			// 
			// Porcentaje
			// 
			this.Porcentaje.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.Porcentaje.AutoSize = true;
			this.Porcentaje.BackColor = System.Drawing.Color.Transparent;
			this.Porcentaje.Location = new System.Drawing.Point(8, 336);
			this.Porcentaje.Name = "Porcentaje";
			this.Porcentaje.Size = new System.Drawing.Size(58, 14);
			this.Porcentaje.TabIndex = 199;
			this.Porcentaje.Text = "Porcentaje";
			// 
			// ImportacionProgramacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(888, 462);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.txtBeneficiario);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtCtaSwift);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.ultraLabel13);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.ultraLabel17);
			this.Controls.Add(this.btFlujo);
			this.Controls.Add(this.btProgramacion);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtPorcentaje);
			this.Controls.Add(this.Porcentaje);
			this.KeyPreview = true;
			this.Name = "ImportacionProgramacion";
			this.Text = "Programación de Pagos de Importación para Proveedores del Asia";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ImportacionProgramacion_KeyDown);
			this.Load += new System.EventHandler(this.ImportacionProgramacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCtaSwift)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBeneficiario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ImportacionProgramacion_Load(object sender, System.EventArgs e)
		{
			this.cmbFecha.Value = DateTime.Today;

			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF , "Select idBanco, Nombre From Banco");
			this.cmbBanco.DisplayLayout.Bands[0].Columns["idBanco"].Hidden = true;

			//string stExec = string.Format("Exec ImportacionPointPrograma2");
			string sSQL = string.Format("Exec ListaPagosPendientesImportacion");
			this.ultraGrid1.DataSource  = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		CultureInfo us = new CultureInfo("en-US");

		private void btProgramacion_Click(object sender, System.EventArgs e)
		{	
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No tiene pagos Pendientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
			{
				//				double Saldo = (double)dr.Cells["saldopro"].Value;
				//			string stSelect = string.Format("Insert INTO PAGO (pendiente, pointpa, IDCOMPRA, PAGO, FECHA, PointidBancoExt, PointCtaSwift,PointBeneficiario,pointpautorizado) values " + "({0}, {1}, {2}, {3},'{4:yyyyMMdd}',{5},'{6}','{7}', {8})",1, 1,IdCompra.ToString(),dPago.ToString("0.00", us), dtFecha, idBanco, txtCtaSwift.Text.ToString(), txtBeneficiario.Text.ToString(), 0);
				//			string sSQL = string.Format("Exec InsertaPagosImportacion {0}, {1}, '{2}', {3}, '{4}', '{5}'", IdCompra, dPago, dtFecha, idBanco, this.txtCtaSwift.Text.ToString(), this.txtBeneficiario.Text.ToString()); 
				//			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				//			string stExec = string.Format("Exec ImportacionPointPrograma2");
				//			ultraGrid1.DataSource  = Funcion.dvProcedimiento(cdsSeteoF, stExec);

				int IdCompra = (int) dr.Cells["idCompra"].Value;
				int IdImportacion = (int) dr.Cells["idImportacionTipo"].Value;
				double dTotal = (double)dr.Cells["Total"].Value;
				DateTime dtFecha = (DateTime) cmbFecha.Value;
				double dPago = double.Parse(txtPago.Value.ToString());
				string sImportacion = dr.Cells["Nombre"].Value.ToString();
				int idBanco = 0;
				if (this.cmbBanco.Text.ToString() != "") idBanco = (int)this.cmbBanco.Value;

				if (dPago > dTotal)
				{
					MessageBox.Show("No se puede programar pago mayor a :" + dTotal.ToString(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				string sSQL = string.Format("Exec InsertaAnticipoProveedorImportacion {0}, {1}, '{2}', {3}, '{4}', {5}, '{6}', '{7}'", IdCompra, IdImportacion, dtFecha.ToString("yyyyMMdd"), dPago, sImportacion, idBanco, this.txtCtaSwift.Text.ToString(), this.txtBeneficiario.Text.ToString());
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				
				MessageBox.Show("Autorizaciones de Pagos de Anticipo y Cancelación de Importación Programados.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information); 

				sSQL = string.Format("Exec ListaPagosPendientesImportacion");
				ultraGrid1.DataSource  = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				this.txtPorcentaje.Value = 0.0;
				this.txtPago.Value = 0.0;
				this.cmbFecha.Value = DateTime.Today;
				this.cmbBanco.Value = 0;
				this.txtCtaSwift.Text = "";
				this.txtBeneficiario.Text = "";
			}
		}

		private void btFlujo_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "";
			Reporte miRep = new Reporte("ImportacionFlujo.rpt", stFiltro);
			miRep.MdiParent = MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;	
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_Click(object sender, System.EventArgs e)
		{
			
		}

		private void txtCtaSwift_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtBeneficiario.Select();
		}

		private void txtPorcentaje_Enter(object sender, System.EventArgs e)
		{
			this.txtPorcentaje.SelectAll();
		}

		private void txtPorcentaje_Click(object sender, System.EventArgs e)
		{
			this.txtPorcentaje.SelectAll();
		}

		private void txtPago_Click(object sender, System.EventArgs e)
		{
			this.txtPago.SelectAll();
		}

		private void txtPago_Enter(object sender, System.EventArgs e)
		{
			this.txtPago.SelectAll();
		}

		private void txtPorcentaje_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtPago.Select();
		}

		private void txtPago_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbFecha.Select();
		}

		private void cmbFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbBanco.Select();
		}

		private void cmbBanco_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCtaSwift.Select();
		}

		private void txtBeneficiario_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btProgramacion.Focus();
		}

		private void txtPorcentaje_Validated(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null || this.ultraGrid1.ActiveRow.Cells["idCompra"].Value == DBNull.Value) return;
			
			double dSaldo = (double)this.ultraGrid1.ActiveRow.Cells["Total"].Value;
			this.txtPago.Value = (dSaldo * (double)this.txtPorcentaje.Value)/100;
		}

		private void txtPorcentaje_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ImportacionProgramacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (ultraGrid1.ActiveRow == null || this.ultraGrid1.ActiveRow.Cells["idCompra"].Value == DBNull.Value)
			{
				return;
			}
			#region Codigo Anterior
			//			double dProgramado = (double)this.ultraGrid1.ActiveRow.Cells["Programado"].Value;
			//			if (dProgramado == 0)
			//			{
			//				this.txtPorcentaje.Value = 40.00;
			//				double dSaldo = (double)this.ultraGrid1.ActiveRow.Cells["Saldo"].Value;
			//				this.txtPago.Value = (dSaldo * (double)this.txtPorcentaje.Value)/ 100;
			//			}
			//			else
			//			{
			//				this.txtPorcentaje.Value = 60.00;
			//				double dSaldo = (double)this.ultraGrid1.ActiveRow.Cells["Saldo"].Value;
			//				this.txtPago.Value = (dSaldo * (double)this.txtPorcentaje.Value)/ 100;
			//			}
			#endregion Codigo Anterior

			this.cmbFecha.Value = null;
			this.cmbBanco.Value = null;
			this.txtCtaSwift.Text = "";
			this.txtBeneficiario.Text = "";

			int idImportacion = (int) this.ultraGrid1.ActiveRow.Cells["idImportacionTipo"].Value;
			string sSQLFecha = string.Format("Exec FechaInicioFinImportacion {0}", idImportacion);
			SqlDataReader drFecha = Funcion.rEscalarSQL(cdsSeteoF, sSQLFecha, true);
			drFecha.Read();
			if (drFecha.HasRows)
			{
				this.cmbFecha.Value = (DateTime)drFecha.GetValue(0);
				this.cmbFecha.MinDate = (DateTime)drFecha.GetValue(0);
			}
			drFecha.Close();
						
			int idCompra = (int) this.ultraGrid1.ActiveRow.Cells["idCompra"].Value;
			string sSQL = string.Format("Exec InformacionClienteImportacion {0}", idCompra);
			SqlDataReader drImp = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drImp.Read();
			if (drImp.HasRows)
			{
				this.cmbBanco.Value = (int)drImp.GetValue(0);
				this.txtCtaSwift.Text = drImp.GetValue(1).ToString();
				this.txtBeneficiario.Text = drImp.GetValue(2).ToString();
			}
			drImp.Close();

			this.txtPorcentaje.Value = 40.00;
			double dSaldo = (double)this.ultraGrid1.ActiveRow.Cells["Total"].Value;
			this.txtPago.Value = (dSaldo * (double)this.txtPorcentaje.Value)/100;
		}
	}
}
