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
	/// Descripción breve de frmActualizaciondeSeriales.
	/// </summary>
	public class frmActualizaciondeSeriales : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnLimpiar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial;
		private System.Windows.Forms.Label label72;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSerial;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdArticulo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmActualizaciondeSeriales()
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
		private void UnloadMe()
		{
			this.Close();
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasDelSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Disponible");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasDelSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Disponible");
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label2 = new System.Windows.Forms.Label();
			this.btnLimpiar = new System.Windows.Forms.Button();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label72 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtIdSerial = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Appearance = appearance1;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 76;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 98;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 154;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Número";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 89;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 89;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 219;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 64;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Local";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 39;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 199;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Sistema";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 154;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 55;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.Caption = "D";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 20;
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
																										 ultraGridColumn13});
			ultraGridBand1.Header.Caption = "MOVIMIENTOS";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(8, 176);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(1048, 184);
			this.ultraGrid2.TabIndex = 304;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(16, 104);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 16);
			this.label2.TabIndex = 320;
			this.label2.Text = "Observaciones";
			// 
			// btnLimpiar
			// 
			this.btnLimpiar.Location = new System.Drawing.Point(736, 136);
			this.btnLimpiar.Name = "btnLimpiar";
			this.btnLimpiar.TabIndex = 315;
			this.btnLimpiar.Text = "&Limpiar";
			this.btnLimpiar.Click += new System.EventHandler(this.btnLimpiar_Click);
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.AcceptsTab = true;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance8;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(104, 104);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(520, 56);
			this.txtObservaciones.TabIndex = 313;
			this.txtObservaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtObservaciones_KeyDown);
			// 
			// btnBuscar
			// 
			this.btnBuscar.Location = new System.Drawing.Point(368, 40);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 311;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Location = new System.Drawing.Point(816, 136);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 316;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Location = new System.Drawing.Point(656, 136);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 314;
			this.btnGuardar.Text = "&Bloquear";
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// txtCodigo
			// 
			this.txtCodigo.AcceptsTab = true;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance9;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(104, 68);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(256, 21);
			this.txtCodigo.TabIndex = 312;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.Validated += new System.EventHandler(this.txtCodigo_Validated);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 16);
			this.label1.TabIndex = 319;
			this.label1.Text = "Articulo";
			// 
			// txtSerial
			// 
			this.txtSerial.AcceptsTab = true;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial.Appearance = appearance10;
			this.txtSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial.Location = new System.Drawing.Point(104, 37);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(256, 21);
			this.txtSerial.TabIndex = 310;
			this.txtSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerial_KeyDown);
			this.txtSerial.Validated += new System.EventHandler(this.txtSerial_Validated);
			// 
			// label72
			// 
			this.label72.AutoSize = true;
			this.label72.BackColor = System.Drawing.Color.Transparent;
			this.label72.Location = new System.Drawing.Point(16, 40);
			this.label72.Name = "label72";
			this.label72.Size = new System.Drawing.Size(33, 16);
			this.label72.TabIndex = 318;
			this.label72.Text = "Serial";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(16, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(43, 16);
			this.lblBodega.TabIndex = 317;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 180;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(104, 6);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(256, 21);
			this.cmbBodega.TabIndex = 0;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
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
			// txtIdSerial
			// 
			this.txtIdSerial.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSerial.Enabled = false;
			this.txtIdSerial.FormatString = "#,##0.00";
			this.txtIdSerial.Location = new System.Drawing.Point(1016, -24);
			this.txtIdSerial.Name = "txtIdSerial";
			this.txtIdSerial.PromptChar = ' ';
			this.txtIdSerial.Size = new System.Drawing.Size(16, 21);
			this.txtIdSerial.TabIndex = 322;
			this.txtIdSerial.Visible = false;
			// 
			// txtIdArticulo
			// 
			this.txtIdArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdArticulo.Enabled = false;
			this.txtIdArticulo.FormatString = "#,##0.00";
			this.txtIdArticulo.Location = new System.Drawing.Point(1032, -24);
			this.txtIdArticulo.Name = "txtIdArticulo";
			this.txtIdArticulo.PromptChar = ' ';
			this.txtIdArticulo.Size = new System.Drawing.Size(16, 21);
			this.txtIdArticulo.TabIndex = 321;
			this.txtIdArticulo.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(bool);
			ultraDataColumn13.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn13});
			// 
			// frmActualizaciondeSeriales
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1064, 366);
			this.Controls.Add(this.txtIdSerial);
			this.Controls.Add(this.txtIdArticulo);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btnLimpiar);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtSerial);
			this.Controls.Add(this.label72);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.ultraGrid2);
			this.KeyPreview = true;
			this.Name = "frmActualizaciondeSeriales";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Bloqueo de Series";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmActualizaciondeSeriales_KeyDown);
			this.Load += new System.EventHandler(this.frmActualizaciondeSeriales_Load);
			this.Validated += new System.EventHandler(this.frmActualizaciondeSeriales_Validated);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmActualizaciondeSeriales_Load(object sender, System.EventArgs e)
		{
			
			if (!Funcion.Permiso("2004", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar al Bloqueo de Seriales", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			string sSQL = string.Format("Exec ListaBodegas {0}, 47", MenuLatinium.IdUsuario);
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			
			this.txtIdSerial.Value = 0;
			this.txtIdArticulo.Value = 0;
			this.txtCodigo.Text = "";	
			this.txtObservaciones.Text = "";

			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ReporteDetalleSeriales 0");

			if (this.txtSerial.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese una Serial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Exec ISMBuscaSerial '{0}'", this.txtSerial.Text.ToString()), true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdSerial.Value = dr.GetInt32(0);
				this.txtIdArticulo.Value = dr.GetInt32(1);
				this.txtCodigo.Text = dr.GetString(2);
			}
			dr.Close();

			if ((int)this.txtIdSerial.Value > 0)
				this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ReporteDetalleSeriales {0}", (int)this.txtIdSerial.Value));
			//				else
			//				{
			//					dr.Close();
			//					MessageBox.Show(string.Format("La serial '{0}' no existe, si desea registrarla ingrese el Articulo", this.txtSerial.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				}
			
		}

		private void btnLimpiar_Click(object sender, System.EventArgs e)
		{
			this.txtIdSerial.Value = 0;
			this.txtSerial.Text = "";	
			this.cmbBodega.Text ="";
			this.txtIdArticulo.Value = 0;
			this.txtCodigo.Text = "";	
			this.txtObservaciones.Text = "";	
			this.txtSerial.Focus();
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				int iBodega = 0;
				if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

				using (BuscaArticulos miBusqueda = new BuscaArticulos(this.txtCodigo.Text.ToString(), iBodega, 0))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.txtCodigo.Text = miBusqueda.sCodigo;
						this.txtIdArticulo.Value = (int)miBusqueda.idArticulo;
					}
				}
			}
		}

		private void txtCodigo_Validated(object sender, System.EventArgs e)
		{
			if (this.txtCodigo.Text.ToString().Length > 0)
			{
				string sSQL = string.Format("Select idArticulo, Codigo From Articulo Where (Codigo = '{0}' Or Articulo = '{0}')", this.txtCodigo.Text.ToString().Trim()); 
				SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drArticulo.Read();
				if (drArticulo.HasRows)
				{
					this.txtIdArticulo.Value = drArticulo.GetInt32(0);
					this.txtCodigo.Text = drArticulo.GetString(1);				
					drArticulo.Close();
				}										
				else
				{
					drArticulo.Close();
					MessageBox.Show("No existe el Articulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtIdArticulo.Value = 0;
				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{	
			if (this.cmbBodega.ActiveRow == null)	
			{
				MessageBox.Show("Seleccione una Bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
				return;
			}
			if (this.txtSerial.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese una Serial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerial.Focus();
				return;
			}
//			if (!Validacion.vbTexto(this.txtMotivo,2,1000,"Motivo")) 
//			{
//				this.txtMotivo.Focus();
//				return;
//			}
//			if (!Validacion.vbTexto(this.txtSerialCorrecto,2,1000,"Serial Correcto")) 
//			{
//				this.txtSerialCorrecto.Focus();
//				return;
//			}
			if (this.txtObservaciones.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese una Nota", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtObservaciones.Focus();
				return;
			}
			
			#region Serial Si Existe
			if (Convert.ToInt32(this.txtIdSerial.Value) > 0)
			{
				if (Convert.ToInt32(this.txtIdArticulo.Value) == 0)
				{
					MessageBox.Show("Borro el Articulo, click en el Boton Buscar para Actualizar la informacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					this.btnBuscar_Click(sender, e);
					return;				
				}
				
				int iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ISMValidaSerialArticulo {0}, {1}", Convert.ToInt32(this.txtIdSerial.Value), Convert.ToInt32(this.txtIdArticulo.Value)));
				
				//if (iCont == 0 && DialogResult.No == MessageBox.Show("CAMBIO DE ARTICULO,\n\nSi registra esta serial con este articulo se modificara el historial de la serial a este Articulo,\n\n¿Desea Continuar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

				if (DialogResult.Yes == MessageBox.Show("Esta seguro de Bloquear esta Serie", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				{					
				//	Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleSeriales Set Disponible = 0,Observaciones = 'BLOQUEADO POR AUDITORIA - {0}' Where idSerial = {1}",this.txtObservaciones.Text, Convert.ToInt32(this.txtIdSerial.Value)));
					
				//	Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleSeriales Set Estado = 25 Where idSerial = {0}", Convert.ToInt32(this.txtIdSerial.Value)));

				//	Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Seriales Set Estado = 25 Where idSerial = {0}", Convert.ToInt32(this.txtIdSerial.Value)));

				//	if (iCont == 0) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Seriales Set idArticulo = {0} Where idSerial = {1}", (int)this.txtIdArticulo.Value, Convert.ToInt32(this.txtIdSerial.Value)));

			
					string sSQLBloqueaSerial = string.Format("Exec BloqueaSeriales {0}, 0, 0, {1}, '{2}', 1 ,'BLOQUEADO POR AUDITORIA',25",
						Convert.ToInt32(this.txtIdSerial.Value),Convert.ToInt32(this.cmbBodega.Value), this.txtObservaciones.Text.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueaSerial, true);

					MessageBox.Show("Bloqueado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
					this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ReporteDetalleSeriales {0}", (int)this.txtIdSerial.Value));
				}
			}
			#endregion Serial Si Existe
		}

		private void frmActualizaciondeSeriales_Validated(object sender, System.EventArgs e)
		{
			this.txtIdSerial.Value = 0;
			this.txtIdArticulo.Value = 0;
			this.txtCodigo.Text = "";			
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ReporteDetalleSeriales 0");		
		}

		private void frmActualizaciondeSeriales_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void txtSerial_Validated(object sender, System.EventArgs e)
		{
			this.btnBuscar_Click(sender, e);
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtSerial.Focus();
		}

		private void txtSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCodigo.Focus();
		}

//		private void txtMotivo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
//		{
//			if (e.KeyCode == Keys.Enter) this.txtSerialCorrecto.Focus();
//		}

//		private void txtSerialCorrecto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
//		{
//			if (e.KeyCode == Keys.Enter) this.txtObservaciones.Focus();
//		}

		private void txtObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}
	}
}
