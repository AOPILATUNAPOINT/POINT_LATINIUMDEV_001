using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmBuscaPromociones.
	/// </summary>
	public class frmBuscaPromociones : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btExportar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbTipo;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridPromociones;
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.RadioGroup rdgEstado;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmBuscaPromociones()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDesde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaHasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioMinimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioMaximo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotaReal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotaCalculada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Variable");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPromocion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDesde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaHasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioMinimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioMaximo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotaReal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotaCalculada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Variable");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			this.uGridPromociones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btExportar = new DevExpress.XtraEditors.SimpleButton();
			this.cmbTipo = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.rdgEstado = new DevExpress.XtraEditors.RadioGroup();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.uGridPromociones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgEstado.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridPromociones
			// 
			this.uGridPromociones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPromociones.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridPromociones.DisplayLayout.Appearance = appearance1;
			this.uGridPromociones.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 19;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 56;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 97;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Header.Caption = "Inicio";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 69;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn5.Header.Caption = "Fin";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 71;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 349;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 64;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 73;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 74;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 82;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 120;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 123;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 109;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 56;
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
																										 ultraGridColumn14});
			ultraGridBand1.Header.Caption = "";
			this.uGridPromociones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPromociones.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridPromociones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridPromociones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridPromociones.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPromociones.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPromociones.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPromociones.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPromociones.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridPromociones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPromociones.Location = new System.Drawing.Point(6, 80);
			this.uGridPromociones.Name = "uGridPromociones";
			this.uGridPromociones.Size = new System.Drawing.Size(716, 216);
			this.uGridPromociones.TabIndex = 285;
			this.uGridPromociones.DoubleClick += new System.EventHandler(this.uGridPromociones_DoubleClick);
			this.uGridPromociones.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridPromociones_KeyPress);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn4.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
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
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
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
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(96, 304);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 24);
			this.btCancelar.TabIndex = 294;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(8, 304);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(80, 24);
			this.btAceptar.TabIndex = 293;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btExportar
			// 
			this.btExportar.Location = new System.Drawing.Point(640, 304);
			this.btExportar.Name = "btExportar";
			this.btExportar.Size = new System.Drawing.Size(80, 24);
			this.btExportar.TabIndex = 301;
			this.btExportar.Text = "&Exportar";
			this.btExportar.Visible = false;
			this.btExportar.Click += new System.EventHandler(this.btExportar_Click);
			// 
			// cmbTipo
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance8;
			this.cmbTipo.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem1.DataValue = "ValueListItem1";
			valueListItem1.DisplayText = "COMBO";
			valueListItem2.DataValue = "ValueListItem0";
			valueListItem2.DisplayText = "PROMOCIÓN";
			valueListItem3.DataValue = "ValueListItem4";
			valueListItem3.DisplayText = "BARATAZO";
			valueListItem4.DataValue = "ValueListItem3";
			valueListItem4.DisplayText = "COMBO PAYJOY";
			this.cmbTipo.Items.Add(valueListItem1);
			this.cmbTipo.Items.Add(valueListItem2);
			this.cmbTipo.Items.Add(valueListItem3);
			this.cmbTipo.Items.Add(valueListItem4);
			this.cmbTipo.Location = new System.Drawing.Point(72, 8);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(144, 22);
			this.cmbTipo.TabIndex = 297;
			this.cmbTipo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTipo_KeyPress);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Enabled = false;
			this.lblCliente.Location = new System.Drawing.Point(8, 43);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 16);
			this.lblCliente.TabIndex = 296;
			this.lblCliente.Text = "Nombre";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance9;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(72, 40);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(392, 22);
			this.txtDescripcion.TabIndex = 295;
			this.txtDescripcion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescripcion_KeyPress);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(480, 40);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(80, 24);
			this.btBuscar.TabIndex = 298;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
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
			// rdgEstado
			// 
			this.rdgEstado.EditValue = 0;
			this.rdgEstado.Location = new System.Drawing.Point(624, 8);
			this.rdgEstado.Name = "rdgEstado";
			// 
			// rdgEstado.Properties
			// 
			this.rdgEstado.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgEstado.Properties.Appearance.Options.UseBackColor = true;
			this.rdgEstado.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Activos"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Finalizados")});
			this.rdgEstado.Size = new System.Drawing.Size(96, 64);
			this.rdgEstado.TabIndex = 299;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(8, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(26, 16);
			this.label1.TabIndex = 300;
			this.label1.Text = "Tipo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmBuscaPromociones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(728, 336);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.rdgEstado);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btExportar);
			this.Controls.Add(this.uGridPromociones);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmBuscaPromociones";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Buscar";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmBuscaPromociones_KeyDown);
			this.Load += new System.EventHandler(this.frmBuscaPromociones_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridPromociones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgEstado.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
			Close();
		}

		private void cmbTipo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDescripcion.Select();
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			string sEstado = "";

			if ((int)this.rdgEstado.EditValue == 1) sEstado = "ACTIVO";
			if ((int)this.rdgEstado.EditValue == 2) sEstado = "FINALIZADO";

			string sSQL = string.Format("Exec BuscaPromociones '{0}', '{1}', '{2}'", this.cmbTipo.Text.ToString().Trim(), this.txtDescripcion.Text.ToString().Trim(), sEstado);
			this.uGridPromociones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.uGridPromociones.Select();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void Aceptar()
		{
			if (this.uGridPromociones.Rows.Count == 0)	
			{
				MessageBox.Show("No hay Promociones realice una busqueda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDescripcion.Focus();
				return;
			}

			if (this.uGridPromociones.ActiveRow == null)	
			{
				MessageBox.Show("Seleccione una Promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridPromociones.Focus();
				return;
			}

			if (this.uGridPromociones.ActiveRow.Cells["idPromocion"].Value == System.DBNull.Value)	
			{
				MessageBox.Show("Seleccione una Promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridPromociones.Focus();
				return;
			}

			this.DialogResult = DialogResult.OK;		
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void frmBuscaPromociones_Load(object sender, System.EventArgs e)
		{
			Acceso miAcceso = new Acceso(cdsSeteoF, "0608");
			if (miAcceso.BExportar) this.btExportar.Visible = true;
			this.cmbTipo.Select();
		}

		private void txtDescripcion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btBuscar_Click(sender, e);
		}

		private void frmBuscaPromociones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();
		}

		private void uGridPromociones_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) Aceptar();
		}

		private void uGridPromociones_DoubleClick(object sender, System.EventArgs e)
		{
			Aceptar();
		}

		private void btExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridPromociones);
		}
	}
}

