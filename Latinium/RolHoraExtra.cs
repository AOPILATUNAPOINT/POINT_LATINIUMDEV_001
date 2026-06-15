using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolHoraExtra.
	/// </summary>
	public class RolHoraExtra : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPersonal;
		private C1.Data.C1DataSet cdsPersonal;
		private Infragistics.Win.Misc.UltraButton btPersonal;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataView cdvEntradaSalida;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private Infragistics.Win.Misc.UltraButton btImprimir;
		private Infragistics.Win.Misc.UltraButton btSalir;
		private Infragistics.Win.Misc.UltraButton btCalcular;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.Misc.UltraButton btCargaHoras;
		private C1.Data.C1DataSet cdsEntradaSalida;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAlmuerzo;
		private Infragistics.Win.Misc.UltraButton btIngresaRol;
		private DevExpress.XtraEditors.LookUpEdit cmbPersonal;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton cmdError;
		private System.ComponentModel.IContainer components;

		public RolHoraExtra()
		{
			InitializeComponent();
		}
		int IdPersonal = 0;
		public RolHoraExtra(int idPersonal)
		{
			InitializeComponent();
			IdPersonal = idPersonal;
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolEntradaSalida", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolEntradaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaHorario");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Atrazo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraExtra");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraSimple");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraDoble");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraNocturna");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Doble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Minutos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MinutosD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaH");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almuerzo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			this.cdsEntradaSalida = new C1.Data.C1DataSet();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkPersonal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.btPersonal = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdvEntradaSalida = new C1.Data.C1DataView();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.btImprimir = new Infragistics.Win.Misc.UltraButton();
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.btCalcular = new Infragistics.Win.Misc.UltraButton();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btCargaHoras = new Infragistics.Win.Misc.UltraButton();
			this.chkAlmuerzo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btIngresaRol = new Infragistics.Win.Misc.UltraButton();
			this.cmbPersonal = new DevExpress.XtraEditors.LookUpEdit();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmdError = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsEntradaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEntradaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsEntradaSalida
			// 
			this.cdsEntradaSalida.BindingContextCtrl = this;
			this.cdsEntradaSalida.DataLibrary = "LibPersonal";
			this.cdsEntradaSalida.DataLibraryUrl = "";
			this.cdsEntradaSalida.DataSetDef = "dsRolEntradaSalida";
			this.cdsEntradaSalida.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEntradaSalida.Name = "cdsEntradaSalida";
			this.cdsEntradaSalida.SchemaClassName = "LibPersonal.DataClass";
			this.cdsEntradaSalida.SchemaDef = null;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(32, 26);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(35, 15);
			this.ultraLabel1.TabIndex = 1;
			this.ultraLabel1.Text = "Desde";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(32, 60);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(33, 15);
			this.ultraLabel2.TabIndex = 2;
			this.ultraLabel2.Text = "Hasta";
			// 
			// cmbDesde
			// 
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(80, 26);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 3;
			this.cmbDesde.Value = new System.DateTime(2008, 2, 21, 0, 0, 0, 0);
			// 
			// cmbHasta
			// 
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(80, 60);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 4;
			this.cmbHasta.Value = new System.DateTime(2008, 2, 21, 0, 0, 0, 0);
			// 
			// chkPersonal
			// 
			this.chkPersonal.Checked = true;
			this.chkPersonal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPersonal.Location = new System.Drawing.Point(216, 26);
			this.chkPersonal.Name = "chkPersonal";
			this.chkPersonal.Size = new System.Drawing.Size(96, 21);
			this.chkPersonal.TabIndex = 5;
			this.chkPersonal.Text = "Todo Personal";
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// btPersonal
			// 
			this.btPersonal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPersonal.Location = new System.Drawing.Point(488, 26);
			this.btPersonal.Name = "btPersonal";
			this.btPersonal.Size = new System.Drawing.Size(96, 25);
			this.btPersonal.TabIndex = 7;
			this.btPersonal.Text = "Generar";
			this.toolTip1.SetToolTip(this.btPersonal, "Calculo de Linea / Calculo por Fecha");
			this.btPersonal.Click += new System.EventHandler(this.btPersonal_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdvEntradaSalida
			// 
			this.cdvEntradaSalida.BindingContextCtrl = this;
			this.cdvEntradaSalida.DataSet = this.cdsEntradaSalida;
			this.cdvEntradaSalida.TableName = "";
			this.cdvEntradaSalida.TableViewName = "RolEntradaSalida";
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrabar.Location = new System.Drawing.Point(24, 431);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(75, 25);
			this.btGrabar.TabIndex = 8;
			this.btGrabar.Text = "Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btImprimir
			// 
			this.btImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btImprimir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btImprimir.Location = new System.Drawing.Point(120, 431);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(75, 25);
			this.btImprimir.TabIndex = 9;
			this.btImprimir.Text = "Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// btSalir
			// 
			this.btSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btSalir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSalir.Location = new System.Drawing.Point(848, 431);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(75, 25);
			this.btSalir.TabIndex = 10;
			this.btSalir.Text = "Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// btCalcular
			// 
			this.btCalcular.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCalcular.Location = new System.Drawing.Point(592, 24);
			this.btCalcular.Name = "btCalcular";
			this.btCalcular.Size = new System.Drawing.Size(75, 25);
			this.btCalcular.TabIndex = 13;
			this.btCalcular.Text = "Calcular";
			this.toolTip1.SetToolTip(this.btCalcular, "Calcula linea Seleccionado / Calcula Todo");
			this.btCalcular.Click += new System.EventHandler(this.btCalcular_Click);
			this.btCalcular.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btCalcular_MouseUp);
			// 
			// btCargaHoras
			// 
			this.btCargaHoras.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCargaHoras.Location = new System.Drawing.Point(488, 56);
			this.btCargaHoras.Name = "btCargaHoras";
			this.btCargaHoras.Size = new System.Drawing.Size(96, 25);
			this.btCargaHoras.TabIndex = 14;
			this.btCargaHoras.Text = "Leer Reloj";
			this.btCargaHoras.Click += new System.EventHandler(this.btCargaHoras_Click);
			// 
			// chkAlmuerzo
			// 
			this.chkAlmuerzo.Location = new System.Drawing.Point(216, 60);
			this.chkAlmuerzo.Name = "chkAlmuerzo";
			this.chkAlmuerzo.Size = new System.Drawing.Size(96, 22);
			this.chkAlmuerzo.TabIndex = 15;
			this.chkAlmuerzo.Text = "Almuerzo";
			this.chkAlmuerzo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkAlmuerzo_MouseUp);
			// 
			// btIngresaRol
			// 
			this.btIngresaRol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btIngresaRol.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btIngresaRol.Location = new System.Drawing.Point(216, 431);
			this.btIngresaRol.Name = "btIngresaRol";
			this.btIngresaRol.Size = new System.Drawing.Size(75, 25);
			this.btIngresaRol.TabIndex = 16;
			this.btIngresaRol.Text = "Ingresa Rol";
			this.btIngresaRol.Click += new System.EventHandler(this.btIngresaRol_Click);
			// 
			// cmbPersonal
			// 
			this.cmbPersonal.Location = new System.Drawing.Point(320, 24);
			this.cmbPersonal.Name = "cmbPersonal";
			// 
			// cmbPersonal.Properties
			// 
			this.cmbPersonal.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																				new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbPersonal.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending),
																																																						new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Codigo", "Codigo", 20, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Default, DevExpress.Data.ColumnSortOrder.None)});
			this.cmbPersonal.Properties.DisplayMember = "Nombre";
			this.cmbPersonal.Properties.NullText = "";
			this.cmbPersonal.Properties.PopupWidth = 300;
			this.cmbPersonal.Properties.ShowFooter = false;
			this.cmbPersonal.Properties.ShowHeader = false;
			this.cmbPersonal.Properties.ValueMember = "Codigo";
			this.cmbPersonal.Size = new System.Drawing.Size(160, 20);
			this.cmbPersonal.TabIndex = 119;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "RolEntradaSalida";
			this.ultraGrid1.DataSource = this.cdsEntradaSalida;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 62;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Format = "ddd dd/MMM/yyyy HH:mm";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.MaskInput = "{LOC}dd/mm/yyyy hh:mm";
			ultraGridColumn3.Width = 113;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 48;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 61;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 70;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 48;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "H. Extra";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 43;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance7;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "H. Simple";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 48;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance8;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "H. Doble";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 52;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance9;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "H. Nocturna";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 48;
			ultraGridColumn12.Header.VisiblePosition = 16;
			ultraGridColumn12.Width = 66;
			ultraGridColumn13.Header.VisiblePosition = 15;
			ultraGridColumn13.Width = 55;
			ultraGridColumn14.Header.Caption = "M. Ex.";
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Width = 43;
			ultraGridColumn15.Header.Caption = "M. D.";
			ultraGridColumn15.Header.VisiblePosition = 13;
			ultraGridColumn15.Width = 45;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Format = "HH:mm";
			ultraGridColumn16.Header.Caption = "H. Doble";
			ultraGridColumn16.Header.VisiblePosition = 14;
			ultraGridColumn16.Width = 51;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Format = "HH:mm";
			ultraGridColumn17.Header.Caption = "Horario";
			ultraGridColumn17.Header.VisiblePosition = 12;
			ultraGridColumn17.Width = 52;
			ultraGridColumn18.Header.VisiblePosition = 17;
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
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 10F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 96);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(920, 319);
			this.ultraGrid1.TabIndex = 120;
			// 
			// cmdError
			// 
			this.cmdError.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmdError.Location = new System.Drawing.Point(592, 56);
			this.cmdError.Name = "cmdError";
			this.cmdError.Size = new System.Drawing.Size(75, 25);
			this.cmdError.TabIndex = 121;
			this.cmdError.Text = "Errores";
			this.toolTip1.SetToolTip(this.cmdError, "Corregir Errores");
			this.cmdError.Click += new System.EventHandler(this.cmdError_Click);
			// 
			// RolHoraExtra
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(952, 471);
			this.Controls.Add(this.cmdError);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.btIngresaRol);
			this.Controls.Add(this.chkAlmuerzo);
			this.Controls.Add(this.btCargaHoras);
			this.Controls.Add(this.btCalcular);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.btPersonal);
			this.Controls.Add(this.chkPersonal);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Name = "RolHoraExtra";
			this.Text = "Horas Extras";
			this.Load += new System.EventHandler(this.RolHoraExtra_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsEntradaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEntradaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void RolHoraExtra_Load(object sender, System.EventArgs e)
		{
			cmbDesde.Value = DateTime.Today.AddDays(-DateTime.Today.Day + 1);
			cmbHasta.Value = DateTime.Today;
			cmbPersonal.Properties.DataSource = Funcion.dvProcedimiento(cdsEntradaSalida,
				"Select Nombre, IsNull(Codigo, '') As Codigo, idPersonal From Personal Where Len(Codigo) > 0 Order by Nombre");
			if (IdPersonal > 0)
			{
				chkPersonal.Checked = false;
				string stSelec = "Select Codigo From Personal Where idPersonal = " + IdPersonal;
				cmbPersonal.EditValue = Funcion.sEscalarSQL(cdsEntradaSalida, stSelec);
				btPersonal_Click(this, e);
			}
		}

		private void btPersonal_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(cmbPersonal, "");
			if (!chkPersonal.Checked)
			{
				if (cmbPersonal.EditValue == null || cmbPersonal.EditValue.ToString() == "")
				{
					errorProvider.SetError(cmbPersonal, "Ingrese Personal");
					return;
				}
			}
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stFiltro = "Fecha >= '" ;
			stFiltro += dtDesde.ToString("yyyyMMdd");
			stFiltro += "' And Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "' ";

			if (chkAlmuerzo.CheckState == System.Windows.Forms.CheckState.Indeterminate) stFiltro += " ";
			else if (chkAlmuerzo.Checked == true) stFiltro += " And Almuerzo = 1 ";
			else stFiltro += " And Almuerzo = 0 ";

			if (!chkPersonal.Checked) 
			{
				string stCodigo = cmbPersonal.EditValue.ToString();
				stFiltro += "And Codigo = '" + stCodigo + "' ";
			}

			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsEntradaSalida, "RolEntradaSalida", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			this.cdsEntradaSalida.Fill(fcTotal, false);
			errorProvider.SetError(ultraGrid1, "");
			if (this.cdvEntradaSalida.Count == 0)
			{
				errorProvider.SetError(ultraGrid1, "No existen registros que cumplan la condicion de busqueda");
				return;
			}
		}

		private void btCalcular_Click(object sender, System.EventArgs e)
		{
			cdsEntradaSalida.Update();
			if (ultraGrid1.ActiveRow == null) 
			{
				MessageBox.Show("Seleccione linea a calcular las Horas Extras", "Información");
				return;
			}
			string stExec = "Exec RolHoraExtra " 
				+ ultraGrid1.ActiveRow.Cells["idRolEntradaSalida"].Value.ToString();
			Funcion.EjecutaSQL(cdsEntradaSalida, stExec, true);
			btPersonal_Click(this, e);
			//MessageBox.Show("Horas Extras Calculadas", "Información");
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsEntradaSalida.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Información");
			}
		}

		private void btCalcular_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;

			if (e.Button != MouseButtons.Right) return;
			Cursor = Cursors.WaitCursor;
			string stExec = "Exec RolHoraExtraTodo '" + dtDesde.ToString("yyyyMMdd")
				+ "', '" + dtHasta.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsPersonal, stExec, true);
			Cursor = Cursors.Default;
			MessageBox.Show("Calculo entre fechas generado");
		}

		private void btCargaHoras_Click(object sender, System.EventArgs e)
		{
			RolReloj miReloj = new RolReloj();
			miReloj.MdiParent = this.MdiParent;
			miReloj.Show();
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stFiltro = "{RolEntradaSalida.Fecha} >= #"
				+ dtDesde.ToString("MM/dd/yyyy") + "# And {RolEntradaSalida.Fecha} < #"
				+ dtHasta.AddDays(1).ToString("MM/dd/yyyy") + "# ";
			if (!chkPersonal.Checked) 
			{
				string stCodigo = cmbPersonal.EditValue.ToString();
				stFiltro += "And {RolEntradaSalida.Codigo} = '" + stCodigo + "' ";
			}
			if (chkAlmuerzo.Checked) stFiltro += "And {RolEntradaSalida.Almuerzo} ";
			else stFiltro += "And Not {RolEntradaSalida.Almuerzo} ";

			Reporte miRep = new Reporte("Rol Hora Extra.rpt", stFiltro);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void chkAlmuerzo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkAlmuerzo.CheckedValue = null;
		}

		private void btIngresaRol_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime) cmbDesde.Value;
      if (DialogResult.No == MessageBox.Show("Desea Ingresar Horas Extras del " + dtFecha.ToString("MM/yyyy"), 
        "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stExec = "exec RolHoraExtraLlenaRol " + dtFecha.Month.ToString() 
				+ ", " + dtFecha.Year.ToString();
			Funcion.EjecutaSQL(cdsEntradaSalida, stExec, true);
			MessageBox.Show("Ingreso de Horas extras a Rol de pago Generado", "Información");
		}

		private void cmdError_Click(object sender, System.EventArgs e)
		{
			int idLinea = 0;
			if (ultraGrid1.ActiveRow != null) 
			{
				idLinea = (int) ultraGrid1.ActiveRow.Cells["idRolEntradaSalida"].Value;
			}

			RolHoraError miRevision = new RolHoraError(idLinea);
			miRevision.MdiParent = this.MdiParent;
			miRevision.Show();
		}
	}
}
