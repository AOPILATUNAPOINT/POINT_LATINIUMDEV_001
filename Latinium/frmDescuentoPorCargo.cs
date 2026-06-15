using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmDescuentoPorCargo.
	/// </summary>
	public class frmDescuentoPorCargo : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECargo;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstados;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDescuentos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDescuentoPorCargo()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDescuentoPorCargo));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDescuentoFormaPagoCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPrecio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenMinimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDescuentoFormaPagoCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio", -1, "cmbEstados");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenMinimo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNECargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label13 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnExportar = new System.Windows.Forms.Button();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnVer = new System.Windows.Forms.Button();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.uGridDescuentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnGuardar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDescuentos)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource9
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbNECargo
			// 
			this.cmbNECargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECargo.DataSource = this.ultraDataSource9;
			ultraGridBand1.AddButtonCaption = "RolTipo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 208;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbNECargo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbNECargo.DisplayMember = "Descripcion";
			this.cmbNECargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECargo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECargo.Location = new System.Drawing.Point(120, 8);
			this.cmbNECargo.Name = "cmbNECargo";
			this.cmbNECargo.Size = new System.Drawing.Size(232, 22);
			this.cmbNECargo.TabIndex = 34;
			this.cmbNECargo.ValueMember = "idCargo";
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
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(16, 11);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(35, 16);
			this.label13.TabIndex = 33;
			this.label13.Text = "Cargo";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(84, 16);
			this.label1.TabIndex = 35;
			this.label1.Text = "Forma De Pago";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.WaitCursor;
			this.cmbFormaPago.DataSource = this.ultraDataSource1;
			ultraGridBand2.AddButtonCaption = "RolTipo";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 232;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "Descripcion";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(120, 40);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(232, 22);
			this.cmbFormaPago.TabIndex = 38;
			this.cmbFormaPago.ValueMember = "idTipoPrecio";
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(657, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 23);
			this.btnExportar.TabIndex = 736;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn5.DefaultValue = 0;
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn7.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn10.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(384, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 735;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// cmbEstados
			// 
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource1;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 160;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEstados.DisplayMember = "Descripcion";
			this.cmbEstados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbEstados.Location = new System.Drawing.Point(496, 168);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(200, 112);
			this.cmbEstados.TabIndex = 734;
			this.cmbEstados.ValueMember = "idTipoPrecio";
			this.cmbEstados.Visible = false;
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargos.DataSource = this.ultraDataSource9;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 160;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCargos.DisplayMember = "Descripcion";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(136, 152);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(200, 112);
			this.cmbCargos.TabIndex = 733;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
			// 
			// uGridDescuentos
			// 
			this.uGridDescuentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDescuentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDescuentos.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDescuentos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 308;
			ultraGridColumn10.Header.Caption = "Cargo";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 180;
			ultraGridColumn11.Header.Caption = "Forma De Pago";
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn11.Width = 180;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance2;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 120;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance3;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Width = 120;
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.Width = 100;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			appearance4.FontData.SizeInPoints = 8F;
			ultraGridBand5.Override.HeaderAppearance = appearance4;
			this.uGridDescuentos.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.uGridDescuentos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDescuentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDescuentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridDescuentos.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDescuentos.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDescuentos.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDescuentos.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridDescuentos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDescuentos.Location = new System.Drawing.Point(8, 88);
			this.uGridDescuentos.Name = "uGridDescuentos";
			this.uGridDescuentos.Size = new System.Drawing.Size(721, 624);
			this.uGridDescuentos.TabIndex = 738;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(736, 8);
			this.groupBox1.TabIndex = 739;
			this.groupBox1.TabStop = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(656, 40);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 740;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// frmDescuentoPorCargo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(737, 718);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.cmbCargos);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbNECargo);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.uGridDescuentos);
			this.Name = "frmDescuentoPorCargo";
			this.Text = "Descuento Por Cargo Forma De Pago";
			this.Load += new System.EventHandler(this.frmDescuentoPorCargo_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDescuentos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmDescuentoPorCargo_Load(object sender, System.EventArgs e)
		{
			this.cmbNECargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCargo, Descripcion From RolCargo Where Descuento = 1");
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoPrecio, Precio As Descripcion From ClientePrecio Where idTipoPrecio In (2, 3, 4, 92)");

			this.cmbCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCargo, Descripcion From RolCargo Where Descuento = 1");
			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoPrecio, Precio As Descripcion From ClientePrecio Where idTipoPrecio In (2, 3, 4, 92)");

			Consulta ();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Consulta ()
		{
			int idCargo = 0;
			int idTipoPrecio = 0;

			if (this.cmbNECargo.ActiveRow != null) idCargo = (int)this.cmbNECargo.Value;
			if (this.cmbFormaPago.ActiveRow != null) idTipoPrecio = (int)this.cmbFormaPago.Value;

			string sSQL = string.Format("Exec DescuentoConsultaGeneral {0}, {1}", idCargo, idTipoPrecio);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDescuentos);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Consulta ();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			int idDescuentoFormaPagoCargo = 0;
			int idTipoPrecio = 0;
			int idCargo = 0;
			decimal dMargenMinimo = 0;
			decimal dDescuento = 0;
			Boolean bEstado = false;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drEstudios in this.uGridDescuentos.Rows.All)
			{	
				if (drEstudios.Cells["idDescuentoFormaPagoCargo"].Value != System.DBNull.Value) idDescuentoFormaPagoCargo = (int)drEstudios.Cells["idDescuentoFormaPagoCargo"].Value;							
				if (drEstudios.Cells["idTipoPrecio"].Value != System.DBNull.Value) idTipoPrecio = (int)drEstudios.Cells["idTipoPrecio"].Value;	
				if (drEstudios.Cells["idCargo"].Value != System.DBNull.Value) idCargo = (int)drEstudios.Cells["idCargo"].Value;	
				if (drEstudios.Cells["MargenMinimo"].Value != System.DBNull.Value) dMargenMinimo = (decimal)drEstudios.Cells["MargenMinimo"].Value;
				if (drEstudios.Cells["Descuento"].Value != System.DBNull.Value) dDescuento = (decimal)drEstudios.Cells["Descuento"].Value;
				if (drEstudios.Cells["Estado"].Value != System.DBNull.Value) bEstado = (Boolean)drEstudios.Cells["Estado"].Value;

				string sSQL = string.Format("Exec DescuentosGuarda {0}, {1}, {2}, {3}, {4}, {5}", 
					idDescuentoFormaPagoCargo, idTipoPrecio, idCargo, dMargenMinimo, dDescuento, bEstado);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
			
      MessageBox.Show("Registros Actualizados Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);	

			Consulta ();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDescuentos);
		}
	}
}
