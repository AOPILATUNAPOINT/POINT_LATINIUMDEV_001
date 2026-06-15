using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCuentasEstadoFinanciero.
	/// </summary>
	public class frmCuentasEstadoFinanciero : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCentrosDeCosto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCentrosDeCosto;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCuentas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuentas;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optGrupos;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optReportes;
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
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCuentas;
		private System.Windows.Forms.Button selec;
		private System.Windows.Forms.GroupBox groupBox2;
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCuentasEstadoFinanciero()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmCuentasEstadoFinanciero(int idCuenta,int tipo,string codigo,DateTime desde,DateTime hasta,string DescripcionH,string codigoH)
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCuentasEstadoFinanciero));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Selec");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Selec");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridCentrosDeCosto = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnVer = new System.Windows.Forms.Button();
			this.optReportes = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.chkCentrosDeCosto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkCuentas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label14 = new System.Windows.Forms.Label();
			this.txtCuentas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.optGrupos = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.uGridCuentas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.selec = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCentrosDeCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optReportes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optGrupos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuentas)).BeginInit();
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
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(200, 12);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(37, 17);
			this.label17.TabIndex = 657;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 12);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(39, 17);
			this.label18.TabIndex = 656;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 10);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 655;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 10);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 654;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// uGridCentrosDeCosto
			// 
			this.uGridCentrosDeCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCentrosDeCosto.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridCentrosDeCosto.DisplayLayout.Appearance = appearance3;
			this.uGridCentrosDeCosto.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.Caption = "N.";
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 32;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 300;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn3.Width = 20;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			ultraGridBand1.Header.Caption = "Centros de Costo";
			ultraGridBand1.HeaderVisible = true;
			this.uGridCentrosDeCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCentrosDeCosto.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridCentrosDeCosto.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCentrosDeCosto.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCentrosDeCosto.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridCentrosDeCosto.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCentrosDeCosto.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridCentrosDeCosto.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCentrosDeCosto.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridCentrosDeCosto.Enabled = false;
			this.uGridCentrosDeCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCentrosDeCosto.Location = new System.Drawing.Point(128, 224);
			this.uGridCentrosDeCosto.Name = "uGridCentrosDeCosto";
			this.uGridCentrosDeCosto.Size = new System.Drawing.Size(341, 200);
			this.uGridCentrosDeCosto.TabIndex = 658;
			this.uGridCentrosDeCosto.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// btnVer
			// 
			this.btnVer.Enabled = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(352, 79);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 659;
			this.btnVer.Text = "Buscar";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// optReportes
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.optReportes.Appearance = appearance10;
			this.optReportes.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optReportes.CheckedIndex = 0;
			this.optReportes.ItemAppearance = appearance11;
			this.optReportes.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 8;
			valueListItem2.DisplayText = "Seleccionados";
			this.optReportes.Items.Add(valueListItem1);
			this.optReportes.Items.Add(valueListItem2);
			this.optReportes.ItemSpacingVertical = 10;
			this.optReportes.Location = new System.Drawing.Point(440, 48);
			this.optReportes.Name = "optReportes";
			this.optReportes.Size = new System.Drawing.Size(168, 24);
			this.optReportes.TabIndex = 672;
			this.optReportes.Text = "Todos";
			this.optReportes.ValueChanged += new System.EventHandler(this.optReportes_ValueChanged);
			// 
			// chkCentrosDeCosto
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkCentrosDeCosto.Appearance = appearance12;
			this.chkCentrosDeCosto.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCentrosDeCosto.Enabled = false;
			this.chkCentrosDeCosto.Location = new System.Drawing.Point(456, 176);
			this.chkCentrosDeCosto.Name = "chkCentrosDeCosto";
			this.chkCentrosDeCosto.Size = new System.Drawing.Size(13, 13);
			this.chkCentrosDeCosto.TabIndex = 673;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn7.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// chkCuentas
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkCuentas.Appearance = appearance13;
			this.chkCuentas.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCuentas.Enabled = false;
			this.chkCuentas.Location = new System.Drawing.Point(37, 155);
			this.chkCuentas.Name = "chkCuentas";
			this.chkCuentas.Size = new System.Drawing.Size(13, 13);
			this.chkCuentas.TabIndex = 675;
			this.chkCuentas.CheckedChanged += new System.EventHandler(this.chkCuentas_CheckedChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(619, 8);
			this.groupBox1.TabIndex = 676;
			this.groupBox1.TabStop = false;
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(8, 83);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(124, 17);
			this.label14.TabIndex = 678;
			this.label14.Text = "Buscar Cuenta Contable";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCuentas
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuentas.Appearance = appearance14;
			this.txtCuentas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCuentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuentas.Enabled = false;
			this.txtCuentas.Location = new System.Drawing.Point(144, 80);
			this.txtCuentas.Name = "txtCuentas";
			this.txtCuentas.Size = new System.Drawing.Size(200, 22);
			this.txtCuentas.TabIndex = 677;
			this.txtCuentas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuentas_KeyDown);
			// 
			// optGrupos
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.optGrupos.Appearance = appearance15;
			this.optGrupos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optGrupos.CheckedIndex = 0;
			this.optGrupos.ItemAppearance = appearance16;
			this.optGrupos.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "Todos";
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "Activo";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "Pasivo";
			valueListItem6.DataValue = 3;
			valueListItem6.DisplayText = "Patrimonio";
			valueListItem7.DataValue = 4;
			valueListItem7.DisplayText = "Ingresos";
			valueListItem8.DataValue = 5;
			valueListItem8.DisplayText = "Gastos";
			this.optGrupos.Items.Add(valueListItem3);
			this.optGrupos.Items.Add(valueListItem4);
			this.optGrupos.Items.Add(valueListItem5);
			this.optGrupos.Items.Add(valueListItem6);
			this.optGrupos.Items.Add(valueListItem7);
			this.optGrupos.Items.Add(valueListItem8);
			this.optGrupos.ItemSpacingVertical = 10;
			this.optGrupos.Location = new System.Drawing.Point(8, 48);
			this.optGrupos.Name = "optGrupos";
			this.optGrupos.Size = new System.Drawing.Size(424, 24);
			this.optGrupos.TabIndex = 679;
			this.optGrupos.Text = "Todos";
			this.optGrupos.ValueChanged += new System.EventHandler(this.optGrupos_ValueChanged);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(512, 80);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 734;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// uGridCuentas
			// 
			this.uGridCuentas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCuentas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCuentas.DataSource = this.ultraDataSource2;
			appearance17.BackColor = System.Drawing.Color.White;
			this.uGridCuentas.DisplayLayout.Appearance = appearance17;
			this.uGridCuentas.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 83;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 204;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 348;
			ultraGridColumn7.Header.Caption = "...";
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 32;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand2.Header.Caption = "Cuentas Contables";
			ultraGridBand2.HeaderVisible = true;
			this.uGridCuentas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCuentas.DisplayLayout.Override.ActiveRowAppearance = appearance18;
			this.uGridCuentas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCuentas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCuentas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance19.BackColor = System.Drawing.Color.Transparent;
			this.uGridCuentas.DisplayLayout.Override.CardAreaAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 10F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCuentas.DisplayLayout.Override.HeaderAppearance = appearance20;
			this.uGridCuentas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance21.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance21.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuentas.DisplayLayout.Override.RowAlternateAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuentas.DisplayLayout.Override.RowSelectorAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuentas.DisplayLayout.Override.SelectedRowAppearance = appearance23;
			this.uGridCuentas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCuentas.Location = new System.Drawing.Point(8, 128);
			this.uGridCuentas.Name = "uGridCuentas";
			this.uGridCuentas.Size = new System.Drawing.Size(605, 304);
			this.uGridCuentas.TabIndex = 674;
			this.uGridCuentas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCuentas_AfterCellUpdate);
			this.uGridCuentas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridCuentas_InitializeLayout);
			// 
			// selec
			// 
			this.selec.Image = ((System.Drawing.Image)(resources.GetObject("selec.Image")));
			this.selec.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.selec.Location = new System.Drawing.Point(432, 79);
			this.selec.Name = "selec";
			this.selec.Size = new System.Drawing.Size(75, 24);
			this.selec.TabIndex = 736;
			this.selec.Text = "Activos";
			this.selec.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.selec.Click += new System.EventHandler(this.selec_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 112);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(616, 8);
			this.groupBox2.TabIndex = 737;
			this.groupBox2.TabStop = false;
			// 
			// frmCuentasEstadoFinanciero
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(619, 438);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.selec);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.optGrupos);
			this.Controls.Add(this.txtCuentas);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.chkCuentas);
			this.Controls.Add(this.uGridCuentas);
			this.Controls.Add(this.chkCentrosDeCosto);
			this.Controls.Add(this.optReportes);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridCentrosDeCosto);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCuentasEstadoFinanciero";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cuentas Estado Financiero";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCuentasEstadoFinanciero_KeyDown);
			this.Load += new System.EventHandler(this.frmCuentasEstadoFinanciero_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCentrosDeCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optReportes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optGrupos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuentas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmCuentasEstadoFinanciero_Load(object sender, System.EventArgs e)
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
					FuncionesProcedimientos.dtGeneral(string.Format("Exec CuentasEstadosFinancieros {0}, '{1}'", 
					(int)this.optReportes.Value, this.txtCuentas.Text.ToString().Trim())); 
				this.txtCuentas.Enabled = true;
				this.chkCuentas.Enabled = true;
				this.uGridCuentas.Enabled = true;
				this.btnVer.Enabled = true;
			}
			else
			{
				this.uGridCuentas.DataSource = 
					FuncionesProcedimientos.dtGeneral(string.Format("Exec CuentasEstadosFinancieros {0}, '{1}'", 
					(int)this.optGrupos.Value, this.txtCuentas.Text.ToString().Trim())); 
			}
			}

		private void txtCuentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) ListaCuentas();
		}

		private void optGrupos_ValueChanged(object sender, System.EventArgs e)
		{
			ListaCuentas();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			ListaCuentas();
		}

		private void chkCuentas_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuentas.Rows.All)
				dr.Cells["Selec"].Value = this.chkCuentas.Checked;
		}

		private void btnImpresion_Click(object sender, System.EventArgs e)
		{	
			if (!Validacion.vbValidaFechasDesdeHasta(this.dtDesde, this.dtHasta)) return;

			string sReporte = "MovimientosDeContabilidad.rpt";

			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramFieldDesde = new ParameterField ();
			ParameterDiscreteValue discreteValDesde = new ParameterDiscreteValue ();
			paramFieldDesde.ParameterFieldName = "@Desde";
			discreteValDesde.Value = Convert.ToDateTime(this.dtDesde.Value);
			paramFieldDesde.CurrentValues.Add (discreteValDesde);
			paramFields.Add (paramFieldDesde);

			ParameterField paramFieldHasta = new ParameterField ();
			ParameterDiscreteValue discreteValHasta = new ParameterDiscreteValue ();
			paramFieldHasta.ParameterFieldName = "@Hasta";
			discreteValHasta.Value = Convert.ToDateTime(this.dtHasta.Value);
			paramFieldHasta.CurrentValues.Add (discreteValHasta);
			paramFields.Add (paramFieldHasta);
			
			ParameterField paramFieldEmpresa = new ParameterField ();
			ParameterDiscreteValue discreteValEmpresa = new ParameterDiscreteValue ();
			paramFieldEmpresa.ParameterFieldName = "@idEmpresa";
			discreteValEmpresa.Value = MenuLatinium.stIdDB; 
			paramFieldEmpresa.CurrentValues.Add (discreteValEmpresa);
			paramFields.Add (paramFieldEmpresa);

			if ((int)this.optReportes.Value == 1)
			{
				string sCuentas = "";

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuentas.Rows.All)	
					if ((bool) dr.Cells["Sel"].Value)
						if (sCuentas.Length < 1) sCuentas = dr.Cells["idCuenta"].Value.ToString(); else sCuentas = sCuentas + ", " + dr.Cells["idCuenta"].Value.ToString();
				
				if (sCuentas.Length == 0)
				{
					MessageBox.Show("Seleccione al menos una Cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				ParameterField paramFieldCuentas = new ParameterField ();
				ParameterDiscreteValue discreteValCuentas = new ParameterDiscreteValue ();
				paramFieldCuentas.ParameterFieldName = "@Cuentas";
				discreteValCuentas.Value = sCuentas;
				paramFieldCuentas.CurrentValues.Add (discreteValCuentas);
				paramFields.Add (paramFieldCuentas);
			}

			if ((int)this.optReportes.Value == 2) sReporte = "BalancesDeComprobacion.rpt";
			
			if ((int)this.optReportes.Value == 3) sReporte = "BalancesGenerales.rpt";
			
			if ((int)this.optReportes.Value == 4) sReporte = "BalancesDeResultados.rpt";
			
			Reporte Reporte = new Reporte(sReporte, "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
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

		private void frmCuentasEstadoFinanciero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
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
						iCont=Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT COUNT(*) FROM Cuenta where Selec = 1"));

						string sSQL = string.Format("Exec EstadoMovimientoContabilidad '{0}', '{1}', {2},{3},{4}",
							Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
							Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
							(int)IDCUENTA,
							codigoh,MenuLatinium.stUsuario);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);


				
				MessageBox.Show(string.Format("Proceso Finalizado Correctamente, {0} Generados", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				
			}
		
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Demo", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;							
			}
		}

		private void uGridCuentas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.Header.Caption;

			if (e.Cell.Column.ToString() == sColumna)
			{
				ActualizaValores((int)e.Cell.Row.Cells["idCuenta"].Value,(bool)e.Cell.Row.Cells["Selec"].Value);
			}
		}
		private void ActualizaValores(int idCuenta, bool sSel)
		{
			string sSQL = string.Format("Exec ActualizaEstadosFinancieros {0}, {1}", idCuenta,sSel);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void selec_Click(object sender, System.EventArgs e)
		{
			using (frmValidacionEstadosFinancieros CCL = new frmValidacionEstadosFinancieros(Codigo,Descrip))
			{		
				if (DialogResult.OK == CCL.ShowDialog())
				{
					 
				}
			}
		}

		private void uGridCuentas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		
	}
}
