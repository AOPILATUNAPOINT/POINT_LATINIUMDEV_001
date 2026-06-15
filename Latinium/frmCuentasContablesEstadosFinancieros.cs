using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCuentasContablesEstadosFinancieros.
	/// </summary>
	public class frmCuentasContablesEstadosFinancieros : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optGrupos;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuentas;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuentas;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCuentas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCentrosDeCosto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCentrosDeCosto;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optReportes;
		private System.Windows.Forms.Button btnVer;
		/// <summary>
		/// Variable del diseñador requerida.
		string Codigo=null;
		int Tipo=0;
		int IDCUENTA=0;
		string prueba=null;
		string Descrip=null;
		string codigoh=null;
		DateTime Desde= DateTime.Now;
		DateTime Hasta= DateTime.Now;
		string CodigoCuenta = null;
		// Booleano que permite, identificar si se puede editar, los dos puntos. 
		bool Cerrado = false;
		private System.Windows.Forms.Label lbl_registros;
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCuentasContablesEstadosFinancieros(int idCuenta,int tipo,string codigo,DateTime desde,DateTime hasta,string DescripcionH,string codigoH, string codigoCuenta, bool cerrado)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			Tipo=tipo;
			Codigo=codigo;
			IDCUENTA=idCuenta;
			Desde=desde;
			Hasta=hasta;
			codigoh=codigoH;
			Descrip=DescripcionH;
			CodigoCuenta = codigoCuenta;
			Cerrado = cerrado;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCuentasContablesEstadosFinancieros));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Selec3");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Selec3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			this.label1 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.optGrupos = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtCuentas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkCuentas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridCuentas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label17 = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridCentrosDeCosto = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.chkCentrosDeCosto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.optReportes = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.lbl_registros = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.optGrupos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCentrosDeCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optReportes)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(56, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(49, 16);
			this.label1.TabIndex = 752;
			this.label1.Text = "Cuentas ";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(48, 96);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(127, 16);
			this.label14.TabIndex = 749;
			this.label14.Text = "Buscar Cuenta Contable";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(304, 48);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 751;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// optGrupos
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optGrupos.Appearance = appearance1;
			this.optGrupos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optGrupos.ItemAppearance = appearance2;
			this.optGrupos.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Activo";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Pasivo";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Patrimonio";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Ingresos";
			valueListItem6.DataValue = 5;
			valueListItem6.DisplayText = "Gastos";
			this.optGrupos.Items.Add(valueListItem1);
			this.optGrupos.Items.Add(valueListItem2);
			this.optGrupos.Items.Add(valueListItem3);
			this.optGrupos.Items.Add(valueListItem4);
			this.optGrupos.Items.Add(valueListItem5);
			this.optGrupos.Items.Add(valueListItem6);
			this.optGrupos.ItemSpacingVertical = 10;
			this.optGrupos.Location = new System.Drawing.Point(44, 46);
			this.optGrupos.Name = "optGrupos";
			this.optGrupos.Size = new System.Drawing.Size(424, 24);
			this.optGrupos.TabIndex = 750;
			this.optGrupos.Visible = false;
			this.optGrupos.ValueChanged += new System.EventHandler(this.optGrupos_ValueChanged);
			// 
			// txtCuentas
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuentas.Appearance = appearance3;
			this.txtCuentas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCuentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuentas.Enabled = false;
			this.txtCuentas.Location = new System.Drawing.Point(184, 88);
			this.txtCuentas.Name = "txtCuentas";
			this.txtCuentas.Size = new System.Drawing.Size(200, 21);
			this.txtCuentas.TabIndex = 748;
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
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(36, 38);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(640, 8);
			this.groupBox1.TabIndex = 747;
			this.groupBox1.TabStop = false;
			// 
			// chkCuentas
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkCuentas.Appearance = appearance4;
			this.chkCuentas.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCuentas.Enabled = false;
			this.chkCuentas.Location = new System.Drawing.Point(56, 152);
			this.chkCuentas.Name = "chkCuentas";
			this.chkCuentas.Size = new System.Drawing.Size(13, 13);
			this.chkCuentas.TabIndex = 746;
			this.chkCuentas.CheckedChanged += new System.EventHandler(this.chkCuentas_CheckedChanged);
			// 
			// uGridCuentas
			// 
			this.uGridCuentas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCuentas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCuentas.DataSource = this.ultraDataSource2;
			appearance5.BackColor = System.Drawing.Color.White;
			this.uGridCuentas.DisplayLayout.Appearance = appearance5;
			this.uGridCuentas.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 122;
			ultraGridColumn2.Header.Caption = "Código";
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 271;
			ultraGridColumn3.Header.Caption = "Descripción";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 240;
			ultraGridColumn4.Header.Caption = "Selec";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 116;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.Header.Caption = "Cuentas Contables";
			ultraGridBand1.HeaderVisible = true;
			this.uGridCuentas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCuentas.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.uGridCuentas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCuentas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCuentas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridCuentas.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCuentas.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.uGridCuentas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuentas.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuentas.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuentas.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.uGridCuentas.Enabled = false;
			this.uGridCuentas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCuentas.Location = new System.Drawing.Point(48, 120);
			this.uGridCuentas.Name = "uGridCuentas";
			this.uGridCuentas.Size = new System.Drawing.Size(648, 280);
			this.uGridCuentas.TabIndex = 745;
			this.uGridCuentas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCuentas_AfterCellUpdate);
			this.uGridCuentas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridCuentas_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(bool);
			ultraDataColumn4.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(236, 10);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(37, 16);
			this.label17.TabIndex = 740;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn7.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(44, 10);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(40, 16);
			this.label18.TabIndex = 739;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance12;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(292, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 738;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance13;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(100, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 737;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// uGridCentrosDeCosto
			// 
			this.uGridCentrosDeCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCentrosDeCosto.DataSource = this.ultraDataSource1;
			appearance14.BackColor = System.Drawing.Color.White;
			this.uGridCentrosDeCosto.DisplayLayout.Appearance = appearance14;
			this.uGridCentrosDeCosto.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "N.";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 32;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 300;
			ultraGridColumn7.Header.Caption = "...";
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn7.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand2.Header.Caption = "Centros de Costo";
			ultraGridBand2.HeaderVisible = true;
			this.uGridCentrosDeCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCentrosDeCosto.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.uGridCentrosDeCosto.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCentrosDeCosto.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCentrosDeCosto.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridCentrosDeCosto.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 10F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCentrosDeCosto.DisplayLayout.Override.HeaderAppearance = appearance17;
			this.uGridCentrosDeCosto.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance18.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance18.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.RowAlternateAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.uGridCentrosDeCosto.Enabled = false;
			this.uGridCentrosDeCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCentrosDeCosto.Location = new System.Drawing.Point(164, 222);
			this.uGridCentrosDeCosto.Name = "uGridCentrosDeCosto";
			this.uGridCentrosDeCosto.Size = new System.Drawing.Size(341, 200);
			this.uGridCentrosDeCosto.TabIndex = 741;
			this.uGridCentrosDeCosto.Visible = false;
			// 
			// chkCentrosDeCosto
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkCentrosDeCosto.Appearance = appearance21;
			this.chkCentrosDeCosto.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCentrosDeCosto.Enabled = false;
			this.chkCentrosDeCosto.Location = new System.Drawing.Point(492, 174);
			this.chkCentrosDeCosto.Name = "chkCentrosDeCosto";
			this.chkCentrosDeCosto.Size = new System.Drawing.Size(13, 13);
			this.chkCentrosDeCosto.TabIndex = 744;
			// 
			// optReportes
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.optReportes.Appearance = appearance22;
			this.optReportes.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optReportes.ItemAppearance = appearance23;
			this.optReportes.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem7.DataValue = 0;
			valueListItem7.DisplayText = "Todos";
			valueListItem8.DataValue = 8;
			valueListItem8.DisplayText = "Seleccionados";
			this.optReportes.Items.Add(valueListItem7);
			this.optReportes.Items.Add(valueListItem8);
			this.optReportes.ItemSpacingVertical = 10;
			this.optReportes.Location = new System.Drawing.Point(112, 48);
			this.optReportes.Name = "optReportes";
			this.optReportes.Size = new System.Drawing.Size(176, 24);
			this.optReportes.TabIndex = 743;
			this.optReportes.ValueChanged += new System.EventHandler(this.optReportes_ValueChanged);
			// 
			// btnVer
			// 
			this.btnVer.Enabled = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(392, 88);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 742;
			this.btnVer.Text = "Buscar";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// lbl_registros
			// 
			this.lbl_registros.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lbl_registros.ForeColor = System.Drawing.Color.Red;
			this.lbl_registros.Location = new System.Drawing.Point(496, 72);
			this.lbl_registros.Name = "lbl_registros";
			this.lbl_registros.Size = new System.Drawing.Size(192, 23);
			this.lbl_registros.TabIndex = 753;
			this.lbl_registros.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmCuentasContablesEstadosFinancieros
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(728, 438);
			this.Controls.Add(this.lbl_registros);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.txtCuentas);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.chkCuentas);
			this.Controls.Add(this.uGridCuentas);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridCentrosDeCosto);
			this.Controls.Add(this.chkCentrosDeCosto);
			this.Controls.Add(this.optReportes);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.optGrupos);
			this.Name = "frmCuentasContablesEstadosFinancieros";
			this.Load += new System.EventHandler(this.frmCuentasContablesEstadosFinancieros_Load);
			((System.ComponentModel.ISupportInitialize)(this.optGrupos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCentrosDeCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optReportes)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCuentasContablesEstadosFinancieros_Load(object sender, System.EventArgs e)
		{
			//			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			//			this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			//
			//			DateTime dtFecha = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
			//
			//			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			//			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			//			
			//			this.dtDesde.CalendarInfo.MaxDate = dtFecha;
			//			this.dtHasta.CalendarInfo.MaxDate = dtFecha;
			
			this.dtDesde.Value = Desde;
			this.dtHasta.Value = Hasta;

			//this.uGridCentrosDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre, CONVERT(Bit, 0) As Sel From Proyecto Where Visible = 1 Order By Nombre");

			this.optReportes.CheckedIndex = 0;
			this.optGrupos.CheckedIndex = 0;
			
			ListaCuentas();
		}

		private void ListaCuentas()
		{
			if((int)this.optReportes.Value==8)
			{
				this.uGridCuentas.DataSource = 
					FuncionesProcedimientos.dtGeneral(string.Format("Exec CuentasEstadosFinancieros2 '{0}', '{1}', {2}, '{3}', '{4}', {5}", 
					CodigoCuenta,
					this.txtCuentas.Text.ToString().Trim(),
					(int)this.optReportes.Value, 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
					Cerrado
					)); 
				this.txtCuentas.Enabled = true;
				this.chkCuentas.Enabled = true;
				this.uGridCuentas.Enabled = true;
				this.btnVer.Enabled = true;
			}
			else
			{
				this.uGridCuentas.DataSource = 
					FuncionesProcedimientos.dtGeneral(string.Format("Exec CuentasEstadosFinancieros2 '{0}', '{1}', {2}, '{3}', '{4}', {5}", 
					CodigoCuenta, 
					this.txtCuentas.Text.ToString().Trim(),
					(int)this.optReportes.Value ,
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
					Cerrado
					)); 
			}
			//Desabilitar los botones para solo lectura, al ya encontrarse cerrado
			if (Cerrado == true)
			{
				this.txtCuentas.Enabled = true;
				this.chkCuentas.Enabled = false;
				//this.uGridCuentas.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnVer.Enabled = true;
			}

		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			ListaCuentas();
		}

		private void optGrupos_ValueChanged(object sender, System.EventArgs e)
		{
			ListaCuentas();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			RegistrosReporteFinal();
		}

		private void RegistrosReporteFinal()
		{
			string sCuentas = "";
			int iCont=0;
			string sSQLVD = "";
			try
			{		

				//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuentas.Rows.All)	
				//					if ((bool) dr.Cells["Selec"].Value)
				//					{
				//						sCuentas = dr.Cells["idCuenta"].Value.ToString();
				//						#region Duplicados
				//						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT COUNT(*) FROM ReporteEstadosAB WHERE Cuenta = '{0}' AND DescripcionCuenta = '{1}'AND Estado=1",sCuentas,Descrip)) > 0)
				//						{
				//							sSQLVD = string.Format("Exec ValidaDuplicadosEstadosFinancieros '{0}'", 
				//							sCuentas,Descrip);					
				//							Funcion.EjecutaSQL(cdsSeteoF, sSQLVD, true);
				//						}
				//						iCont++;
				//						#endregion Duplicados
				//					
				//					}
				//				#region REPETIDOS
				//				if (Funcion.iEscalarSQL(cdsSeteoF, "select count(*) from  ValidaReporteEstadosAB") > 0)
				//				{
				//					frmValidacionEstadosFinancieros VRB = new frmValidacionEstadosFinancieros();
				//					VRB.ShowDialog();
				//
				//					Funcion.EjecutaSQL(cdsSeteoF, "Delete From ValidaReporteEstadosAB");
				//
				//					this.Cursor = Cursors.Default;
				//
				//					return;
				//				}
				//
				//				#endregion REPETIDOS
				iCont=Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CuentasContablesSelecCodigoCuenta '{0}'", CodigoCuenta));

//				string sSQL = string.Format("Exec EstadoMovimientoContabilidad2 '{0}', '{1}', {2},{3},{4}",
//					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
//					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
//					(int)IDCUENTA,
//					codigoh,MenuLatinium.stUsuario);
//				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				this.lbl_registros.Text = iCont.ToString()+ " GENERADOS";
				MessageBox.Show(string.Format("Proceso Finalizado Correctamente, {0} Generados", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Demo", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;							
			}
		}


		private void optReportes_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int)this.optReportes.Value < 3)
			{
				this.txtCuentas.Enabled = true;
				this.chkCuentas.Enabled = true;
				this.uGridCuentas.Enabled = true;
				this.btnVer.Enabled = true;
			}
		}

		private void chkCuentas_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuentas.Rows.All)
				dr.Cells["Selec3"].Value = this.chkCuentas.Checked;
		}

		private void uGridCuentas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.Header.Caption;

			if (e.Cell.Column.Header.Caption.ToString() == sColumna)
			{
				ActualizaValores((int)e.Cell.Row.Cells["idCuenta"].Value,(bool)e.Cell.Row.Cells["Selec3"].Value);
			}
		}
		private void ActualizaValores(int idCuenta, bool sSel)
		{
			string sSQL = string.Format("Exec ActualizaEstadosFinancieros2 {0}, {1}", idCuenta,sSel);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void uGridCuentas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			if (Cerrado == true)
			{
							e.Layout.Bands[0].Columns["Selec3"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			}

		}
	}
}
