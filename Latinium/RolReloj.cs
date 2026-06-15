using System;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolReloj.
	/// </summary>
	public class RolReloj : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsEntradaSalida;
		private C1.Data.C1DataSet cdsSeteoRol;
		private C1.Data.C1DataView cdvEntradaSalida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnDesde;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnHasta;
		private Infragistics.Win.Misc.UltraButton btLunch;
		private Infragistics.Win.Misc.UltraButton btError;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolReloj()
		{
			InitializeComponent();
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraDoble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraNocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nocturna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Doble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Minutos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MinutosD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaH");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almuerzo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsEntradaSalida = new C1.Data.C1DataSet();
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsSeteoRol = new C1.Data.C1DataSet();
			this.cdvEntradaSalida = new C1.Data.C1DataView();
			this.spnDesde = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btLunch = new Infragistics.Win.Misc.UltraButton();
			this.btError = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEntradaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRol)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEntradaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnHasta)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbHasta
			// 
			dateButton1.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton1);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(272, 17);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 8;
			this.cmbHasta.Value = new System.DateTime(2008, 2, 21, 0, 0, 0, 0);
			// 
			// cmbDesde
			// 
			dateButton2.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton2);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(96, 17);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 7;
			this.cmbDesde.Value = new System.DateTime(2008, 2, 21, 0, 0, 0, 0);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(224, 17);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(33, 15);
			this.ultraLabel2.TabIndex = 6;
			this.ultraLabel2.Text = "Hasta";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(48, 17);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(35, 15);
			this.ultraLabel1.TabIndex = 5;
			this.ultraLabel1.Text = "Desde";
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(16, 414);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(48, 25);
			this.btAceptar.TabIndex = 10;
			this.btAceptar.Text = "S&ubir";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(392, 414);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 25);
			this.btCancelar.TabIndex = 11;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
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
			ultraGridColumn2.Width = 75;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Format = "ddd, dd/MMM/yyyy HH:mm";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.MaskInput = "{LOC}dd/mm/yyyy hh:mm";
			ultraGridColumn3.Width = 183;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 143;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 105;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 95;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 105;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 108;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 43;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 47;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 57;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 46;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 48;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 35;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 41;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 63;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 76;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 58;
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
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 60);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(480, 336);
			this.ultraGrid1.TabIndex = 12;
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
			this.cdsEntradaSalida.BeforeFill += new C1.Data.FillEventHandler(this.cdsEntradaSalida_BeforeFill);
			// 
			// btFiltro
			// 
			this.btFiltro.Location = new System.Drawing.Point(408, 17);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(48, 26);
			this.btFiltro.TabIndex = 101;
			this.btFiltro.Text = "Filtro";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cdsSeteoRol
			// 
			this.cdsSeteoRol.BindingContextCtrl = this;
			this.cdsSeteoRol.DataLibrary = "LibPersonal";
			this.cdsSeteoRol.DataLibraryUrl = "";
			this.cdsSeteoRol.DataSetDef = "dsSeteoRol";
			this.cdsSeteoRol.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoRol.Name = "cdsSeteoRol";
			this.cdsSeteoRol.SchemaClassName = "LibPersonal.DataClass";
			this.cdsSeteoRol.SchemaDef = null;
			// 
			// cdvEntradaSalida
			// 
			this.cdvEntradaSalida.BindingContextCtrl = this;
			this.cdvEntradaSalida.DataSet = this.cdsEntradaSalida;
			this.cdvEntradaSalida.TableName = "";
			this.cdvEntradaSalida.TableViewName = "RolEntradaSalida";
			// 
			// spnDesde
			// 
			this.spnDesde.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.spnDesde.Location = new System.Drawing.Point(128, 414);
			this.spnDesde.MaskInput = "nn";
			this.spnDesde.MaxValue = 23;
			this.spnDesde.MinValue = 0;
			this.spnDesde.Name = "spnDesde";
			this.spnDesde.PromptChar = ' ';
			this.spnDesde.Size = new System.Drawing.Size(48, 22);
			this.spnDesde.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnDesde.TabIndex = 103;
			this.spnDesde.Value = 12;
			// 
			// spnHasta
			// 
			this.spnHasta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.spnHasta.Location = new System.Drawing.Point(176, 414);
			this.spnHasta.MaskInput = "nn";
			this.spnHasta.MaxValue = 23;
			this.spnHasta.MinValue = 0;
			this.spnHasta.Name = "spnHasta";
			this.spnHasta.PromptChar = ' ';
			this.spnHasta.Size = new System.Drawing.Size(48, 22);
			this.spnHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnHasta.TabIndex = 105;
			this.spnHasta.Value = 14;
			// 
			// btLunch
			// 
			this.btLunch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btLunch.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btLunch.Location = new System.Drawing.Point(232, 414);
			this.btLunch.Name = "btLunch";
			this.btLunch.Size = new System.Drawing.Size(48, 25);
			this.btLunch.TabIndex = 106;
			this.btLunch.Text = "&Lunch";
			this.btLunch.Click += new System.EventHandler(this.btLunch_Click);
			// 
			// btError
			// 
			this.btError.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btError.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btError.Location = new System.Drawing.Point(288, 414);
			this.btError.Name = "btError";
			this.btError.Size = new System.Drawing.Size(48, 25);
			this.btError.TabIndex = 107;
			this.btError.Text = "&Error";
			this.btError.Click += new System.EventHandler(this.btError_Click);
			// 
			// RolReloj
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(512, 454);
			this.Controls.Add(this.btError);
			this.Controls.Add(this.btLunch);
			this.Controls.Add(this.spnHasta);
			this.Controls.Add(this.spnDesde);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Name = "RolReloj";
			this.Text = "Reloj";
			this.Load += new System.EventHandler(this.RolReloj_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsEntradaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRol)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEntradaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnHasta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void Filtro()
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;

			string stFiltro = "Fecha >= '" + dtDesde.ToString("yyyyMMdd") + 
				"' And Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'";
			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsEntradaSalida, "RolEntradaSalida", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);
			this.cdsEntradaSalida.Fill(fcTotal, false);
			ultraGrid1.Refresh();
		}

		private void RolReloj_Load(object sender, System.EventArgs e)
		{
			cmbDesde.Value = DateTime.Today.AddDays(-1);
			cmbHasta.Value = DateTime.Today.AddDays(-1);
//			cmbHoy.Value = DateTime.Today;
			Filtro();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stConeccion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
			string stCodigo = Funcion.sEscalarSQL(cdsEntradaSalida, "Select IsNull(RelojCodigo, '') From SeteoRol");
			if (stCodigo.Length == 0) // Lee desde SQL
			{
				string stExec = String.Format("Exec RolRelojLee '{0}', '{1}'", dtDesde.ToString("yyyyMMdd"),
						dtHasta.ToString("yyyyMMdd"));
				string strBase = "RelojSQL";
//				string strBase = "Practica01";
				string strServidor = @".\SqlExpress";
				string strUsuario = "infoelect";

				stConeccion = @"Password=Bl45o6$9;"
					+ "Persist Security Info=True;User ID=" + strUsuario.Trim()
					+ ";Initial Catalog=" + strBase.Trim()
					+ ";Data Source=" + strServidor.Trim();

        SqlDataReader dr = Funcion.rInstanciaSQL(stConeccion, stExec, true);
				while(dr.Read())
				{
					string stLeeLinea = string.Format("Exec RolHoraExtraIngresa '{0}', '{1}', {2}",
						dr.GetValue(0).ToString(), dr.GetDateTime(1).ToString("yyyyMMdd HH:mm:ss"), dr.GetInt32(2));
					Funcion.EjecutaSQL(cdsEntradaSalida, stLeeLinea, true);
				}
				dr.Close();
				MessageBox.Show("Datos Subidos");
				return;
			}

			string stRelojB = Funcion.sEscalarSQL(cdsEntradaSalida, "Select IsNull(RelojBiometrico, '') From SeteoRol");

			if (stRelojB.Trim().Length == 0) stRelojB = "c:\\Bio-Office\\OA100.mdb";
			if (!File.Exists(stRelojB)) 
			{
				MessageBox.Show("No existe Archivo " + stRelojB + " para leer los datos del Reloj Biometrico",
					"Información");
				SeteoR miSeteo = new SeteoR();
				miSeteo.MdiParent = this.MdiParent;
				miSeteo.Show();
				return;
			}
			stConeccion += stRelojB;
			OleDbConnection aConnection = new OleDbConnection(stConeccion);

			string stCodigo2 = Funcion.sEscalarSQL(cdsEntradaSalida, "Select IsNull(RelojCodigo2, '') From SeteoRol");
			string stHora = Funcion.sEscalarSQL(cdsEntradaSalida, "Select IsNull(RelojHora, '') From SeteoRol");
			string stEntradaS = Funcion.sEscalarSQL(cdsEntradaSalida, "Select IsNull(RelojEntrada, '') From SeteoRol");
			string stTabla = Funcion.sEscalarSQL(cdsEntradaSalida, "Select IsNull(RelojTabla, '') From SeteoRol");
			string stTabla2 = Funcion.sEscalarSQL(cdsEntradaSalida, "Select IsNull(RelojTabla2, '') From SeteoRol");
			if (stTabla2.Length > 0)
			{
				stTabla2 = " Inner Join " + stTabla2 + " t2 On t1." + stCodigo
					+ " = t2." + stCodigo + " ";
				stCodigo = "t2." + stCodigo2;
			}

			int iTipoIO = Funcion.iEscalarSQL(cdsEntradaSalida, "Select IsNull(RelojTipoIO, '') From SeteoRol");
			string stTextoIO = Funcion.sEscalarSQL(cdsEntradaSalida, "Select IsNull(RelojTextoIO, '') From SeteoRol");
			string stSelect = "select " + stCodigo + ", t1." + stHora + ", t1."
				+ stEntradaS + " from " + stTabla + " As t1 " + stTabla2
				+ "Where t1." + stHora + " >= #"
				+ dtDesde.ToString("MM-dd-yyyy") + "# And t1." + stHora + " < #" 
				+ dtHasta.AddDays(1).ToString("MM-dd-yyyy") + "#";

			OleDbCommand aCommand = new OleDbCommand(stSelect, aConnection);
			string stLee = "";
			try
			{
				aConnection.Open();
				OleDbDataReader aReader = aCommand.ExecuteReader();
				while(aReader.Read())
				{
					string stEntrada = "', 0";
					if (iTipoIO == 0)	if (aReader.GetBoolean(2) == true) stEntrada = "', 1";
					if (iTipoIO == 1)	if (aReader.GetInt32(2) == 1) stEntrada = "', 1";
					if (iTipoIO == 2)	if (aReader.GetString(2) == stTextoIO) stEntrada = "', 1";
					stLee = "Exec RolHoraExtraIngresa " + aReader.GetValue(0).ToString() + ", '"
						+ aReader.GetDateTime(1).ToString("yyyyMMdd HH:mm:ss") + stEntrada;
					Funcion.EjecutaSQL(cdsEntradaSalida, stLee, true);
				}
				aReader.Close();
				aConnection.Close();
			}
			catch(OleDbException ex)
			{
				MessageBox.Show("Error: {0}", ex.Errors[0].Message);
			}
			MessageBox.Show("Datos Subidos");
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			cdsEntradaSalida.Update();
		}

		private void btNuevo_Click(object sender, System.EventArgs e)
		{

		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			Filtro();	
		}

		private void cdsEntradaSalida_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsEntradaSalida.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void btLunch_Click(object sender, System.EventArgs e)
		{
			string stExec = "Exec RolAlmuerzo " + spnDesde.Value.ToString()
				+ ", " + spnHasta.Value.ToString();
			Funcion.EjecutaSQL(cdsEntradaSalida, stExec, true);
			MessageBox.Show("Horas de Almuerzo procesadas", "Información");
		}

		private void btError_Click(object sender, System.EventArgs e)
		{
			RolHoraError miRolError = new RolHoraError();
			miRolError.MdiParent = this.MdiParent;
			miRolError.Show();
		}

	}
}
