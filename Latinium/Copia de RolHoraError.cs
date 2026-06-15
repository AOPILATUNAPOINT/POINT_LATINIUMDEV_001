using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolHoraError.
	/// </summary>
	public class RolHoraError : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnHasta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDesde;
		private C1.Data.C1DataSet cdsEntradaSalida;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEntradaSalida;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEntradaDia;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private C1.Data.C1DataView cdvEntradaSalida;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPersonal;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAlmuerzo;
		private DevExpress.XtraEditors.LookUpEdit cmbPersonal;
		private Infragistics.Win.Misc.UltraButton btRangoDias;
		private Infragistics.Win.Misc.UltraButton btFiltroDia;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolHoraError()
		{
			InitializeComponent();
		}

		public RolHoraError(int IdLinea)
		{
			InitializeComponent();
			idLinea = IdLinea;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entrada");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolEntradaSalida", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolEntradaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Atrazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraExtra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraSimple");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraDoble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraNocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Doble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Minutos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MinutosD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaH");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almuerzo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsEntradaSalida = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btRangoDias = new Infragistics.Win.Misc.UltraButton();
			this.spnHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnDesde = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsEntradaSalida = new C1.Data.C1DataSet();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.btFiltroDia = new Infragistics.Win.Misc.UltraButton();
			this.udsEntradaDia = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdvEntradaSalida = new C1.Data.C1DataView();
			this.chkPersonal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.chkAlmuerzo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbPersonal = new DevExpress.XtraEditors.LookUpEdit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEntradaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEntradaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEntradaDia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEntradaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.udsEntradaSalida;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 88;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Format = "dd/MMM/yyyy";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Location = new System.Drawing.Point(24, 112);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(272, 276);
			this.ultraGrid1.TabIndex = 0;
			// 
			// udsEntradaSalida
			// 
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			this.udsEntradaSalida.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// btRangoDias
			// 
			this.btRangoDias.Location = new System.Drawing.Point(24, 52);
			this.btRangoDias.Name = "btRangoDias";
			this.btRangoDias.Size = new System.Drawing.Size(152, 24);
			this.btRangoDias.TabIndex = 1;
			this.btRangoDias.Text = "Rango de Registro Diarios";
			this.btRangoDias.Click += new System.EventHandler(this.btRangoDias_Click);
			// 
			// spnHasta
			// 
			this.spnHasta.Location = new System.Drawing.Point(240, 52);
			this.spnHasta.MaskInput = "nn";
			this.spnHasta.MaxValue = 99;
			this.spnHasta.MinValue = 0;
			this.spnHasta.Name = "spnHasta";
			this.spnHasta.PromptChar = ' ';
			this.spnHasta.Size = new System.Drawing.Size(48, 22);
			this.spnHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnHasta.TabIndex = 107;
			this.spnHasta.Value = 9;
			// 
			// spnDesde
			// 
			this.spnDesde.Location = new System.Drawing.Point(184, 52);
			this.spnDesde.MaskInput = "nn";
			this.spnDesde.MaxValue = 23;
			this.spnDesde.MinValue = 0;
			this.spnDesde.Name = "spnDesde";
			this.spnDesde.PromptChar = ' ';
			this.spnDesde.Size = new System.Drawing.Size(48, 22);
			this.spnDesde.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDesde.TabIndex = 106;
			this.spnDesde.Value = 1;
			// 
			// cdsEntradaSalida
			// 
			this.cdsEntradaSalida.BindingContextCtrl = this;
			this.cdsEntradaSalida.DataLibrary = "LibPersonal";
			this.cdsEntradaSalida.DataLibraryUrl = "";
			this.cdsEntradaSalida.DataSetDef = "dsRolEntradaSalida";
			this.cdsEntradaSalida.FillOnRequest = false;
			this.cdsEntradaSalida.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEntradaSalida.Name = "cdsEntradaSalida";
			this.cdsEntradaSalida.SchemaClassName = "LibPersonal.DataClass";
			this.cdsEntradaSalida.SchemaDef = null;
			// 
			// btGrabar
			// 
			this.btGrabar.Location = new System.Drawing.Point(528, 52);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(104, 24);
			this.btGrabar.TabIndex = 108;
			this.btGrabar.Text = "Grabar";
			this.btGrabar.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// btFiltroDia
			// 
			this.btFiltroDia.Location = new System.Drawing.Point(312, 52);
			this.btFiltroDia.Name = "btFiltroDia";
			this.btFiltroDia.Size = new System.Drawing.Size(112, 24);
			this.btFiltroDia.TabIndex = 109;
			this.btFiltroDia.Text = "Filtro del Dia";
			this.btFiltroDia.Click += new System.EventHandler(this.btFiltroDia_Click);
			// 
			// udsEntradaDia
			// 
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(bool);
			this.udsEntradaDia.Band.Columns.AddRange(new object[] {
																															ultraDataColumn4,
																															ultraDataColumn5,
																															ultraDataColumn6});
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "RolEntradaSalida";
			this.ultraGrid2.DataSource = this.cdsEntradaSalida;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance6;
			ultraGridBand2.AddButtonCaption = "Band 0";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 42;
			ultraGridColumn6.Format = "dd/MMM/yyyy HH:mm";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.MaskInput = "{LOC}dd/mm/yyyy hh:mm";
			ultraGridColumn6.Width = 125;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Width = 69;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 14;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 15;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 16;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 17;
			ultraGridColumn21.Width = 53;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(312, 86);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(320, 302);
			this.ultraGrid2.TabIndex = 110;
			// 
			// cdvEntradaSalida
			// 
			this.cdvEntradaSalida.BindingContextCtrl = this;
			this.cdvEntradaSalida.DataSet = this.cdsEntradaSalida;
			this.cdvEntradaSalida.TableName = "";
			this.cdvEntradaSalida.TableViewName = "RolEntradaSalida";
			// 
			// chkPersonal
			// 
			this.chkPersonal.Checked = true;
			this.chkPersonal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPersonal.Location = new System.Drawing.Point(24, 17);
			this.chkPersonal.Name = "chkPersonal";
			this.chkPersonal.Size = new System.Drawing.Size(96, 22);
			this.chkPersonal.TabIndex = 111;
			this.chkPersonal.Text = "Todo Personal";
			// 
			// cmbHasta
			// 
			dateButton1.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton1);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(512, 17);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 116;
			this.cmbHasta.Value = new System.DateTime(2008, 2, 21, 0, 0, 0, 0);
			// 
			// cmbDesde
			// 
			dateButton2.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton2);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(336, 17);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 115;
			this.cmbDesde.Value = new System.DateTime(2008, 2, 21, 0, 0, 0, 0);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(464, 17);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(33, 15);
			this.ultraLabel2.TabIndex = 114;
			this.ultraLabel2.Text = "Hasta";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(288, 17);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(35, 15);
			this.ultraLabel1.TabIndex = 113;
			this.ultraLabel1.Text = "Desde";
			// 
			// chkAlmuerzo
			// 
			this.chkAlmuerzo.Location = new System.Drawing.Point(32, 86);
			this.chkAlmuerzo.Name = "chkAlmuerzo";
			this.chkAlmuerzo.Size = new System.Drawing.Size(128, 22);
			this.chkAlmuerzo.TabIndex = 117;
			this.chkAlmuerzo.Text = "Mostrar Almuerzo";
			this.chkAlmuerzo.ThreeState = true;
			this.chkAlmuerzo.MouseUp += new System.Windows.Forms.MouseEventHandler(this.chkAlmuerzo_MouseUp);
			// 
			// cmbPersonal
			// 
			this.cmbPersonal.Location = new System.Drawing.Point(120, 16);
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
			this.cmbPersonal.TabIndex = 118;
			// 
			// RolHoraError
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(664, 401);
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.chkAlmuerzo);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.chkPersonal);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.btFiltroDia);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.spnHasta);
			this.Controls.Add(this.spnDesde);
			this.Controls.Add(this.btRangoDias);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "RolHoraError";
			this.Text = "Errores en Entradas Salidas";
			this.Load += new System.EventHandler(this.RolHoraError_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEntradaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEntradaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEntradaDia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEntradaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsEntradaSalida.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al Grabar");
			}
		}

		int idLinea = 0;
		private void RolHoraError_Load(object sender, System.EventArgs e)
		{
			chkAlmuerzo.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			cmbDesde.Value = DateTime.Today.AddDays(-2);
			cmbHasta.Value = DateTime.Today;
			cmbPersonal.Properties.DataSource = Funcion.dvProcedimiento(cdsEntradaSalida,
				"Select Nombre, IsNull(Codigo, '') As Codigo, idPersonal From Personal Where Len(Codigo) > 0 Order by Nombre");
			if (idLinea > 0)
			{
				string stSelect = "Select LTRIM(RTRIM(Codigo)) From RolEntradaSalida Where idRolEntradaSalida = " + idLinea;
				cmbPersonal.EditValue = Funcion.sEscalarSQL(cdsEntradaSalida, stSelect);
				chkPersonal.Checked = false;
				stSelect = "Select Fecha From RolEntradaSalida Where idRolEntradaSalida = " + idLinea;
				DateTime dtFecha = Funcion.dtEscalarSQL(cdsEntradaSalida, stSelect);
				cmbDesde.Value = dtFecha;
				cmbHasta.Value = dtFecha;
				btRangoDias_Click(this, e);
			}
		}

		private void chkAlmuerzo_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			chkAlmuerzo.CheckedValue = null;
		}

		private void btRangoDias_Click(object sender, System.EventArgs e)
		{
			string stExec = "Exec RolHoraError " + spnDesde.Value.ToString()
				+ ", " + spnHasta.Value.ToString() + ", ";
			if (!chkPersonal.Checked) 
			{
				string stCodigo = cmbPersonal.EditValue.ToString();
				stExec += "'" + stCodigo + "' ";
			}
			else
				stExec += "''";

			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			stExec += ", '" + dtDesde.ToString("yyyyMMdd") + "', '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'";

			if (chkAlmuerzo.CheckState == System.Windows.Forms.CheckState.Indeterminate) stExec += ", 2";
			else if (chkAlmuerzo.Checked == true) stExec += ", 1";
			else stExec += ", 0";

			SqlDataReader dr = Funcion.rEscalarSQL(cdsEntradaSalida, stExec, true);
			udsEntradaSalida.Rows.Clear();
			int iNumFilas = -1;
			while (dr.Read())
			{
				iNumFilas ++;
				udsEntradaSalida.Rows.Add();
				udsEntradaSalida.Rows[iNumFilas]["Codigo"]=dr.GetString(0);
				udsEntradaSalida.Rows[iNumFilas]["Cuenta"]=dr.GetInt32(1);
				udsEntradaSalida.Rows[iNumFilas]["Fecha"]=dr.GetDateTime(2);
			}
			dr.Close();
		}

		private void btFiltroDia_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null) 
			{
				MessageBox.Show("Seleccione Fila con datos de grilla de la izquierda", "Información");
				return;
			}
			string stCodigo = ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
			DateTime dtFecha = (DateTime) ultraGrid1.ActiveRow.Cells["Fecha"].Value;
			string stFiltro = "Fecha >= '" ;
			stFiltro += dtFecha.ToString("yyyyMMdd");
			stFiltro += "' And Fecha < '" + dtFecha.AddDays(1).ToString("yyyyMMdd");
			stFiltro += "' And Codigo = '" + stCodigo + "'";

			if (chkAlmuerzo.CheckState == System.Windows.Forms.CheckState.Indeterminate) stFiltro += " ";
			else if (chkAlmuerzo.Checked == true) stFiltro += " And almuerzo = 1";
			else stFiltro += " And almuerzo = 0";

			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsEntradaSalida, "RolEntradaSalida", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			this.cdsEntradaSalida.Fill(fcTotal, false);
			if (this.cdvEntradaSalida.Count == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

		}
	}
}
