using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de EstadoCuente.
	/// </summary>
	public class EstadoCuenta : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsEstadoCuenta;
		private C1.Data.C1DataView cdvEstadoCuenta;
		private C1.Data.C1DataSet cdsCuentaCorr;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private C1.Data.C1DataSet cdsSeguridad;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.CheckBox chkFiltroFecha;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private System.Windows.Forms.CheckBox chkNoConfirmado;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.Misc.UltraButton btPreConci;
		private System.Windows.Forms.CheckBox chkAnulado;
		private DevExpress.XtraEditors.DateEdit dtFechaIni;
		private DevExpress.XtraEditors.DateEdit dtFechaFin;
		private DevExpress.XtraEditors.SimpleButton btChequeAntes;
		private DevExpress.XtraEditors.SimpleButton btFiltro;
		private DevExpress.XtraEditors.SimpleButton btExcel;
		private DevExpress.XtraEditors.SimpleButton btConciliaAuto;
		private DevExpress.XtraEditors.SimpleButton btAsiento;
		private System.ComponentModel.IContainer components;

		public EstadoCuenta()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("EstadoCuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaConf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalH");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fechaComp");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCostoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProyectoBR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNoDeducible");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CuentaCorriente", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Automatico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VisibleLocales");
			this.cdsEstadoCuenta = new C1.Data.C1DataSet();
			this.cdvEstadoCuenta = new C1.Data.C1DataView();
			this.cdsCuentaCorr = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.chkFiltroFecha = new System.Windows.Forms.CheckBox();
			this.btPreConci = new Infragistics.Win.Misc.UltraButton();
			this.chkNoConfirmado = new System.Windows.Forms.CheckBox();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.chkAnulado = new System.Windows.Forms.CheckBox();
			this.dtFechaIni = new DevExpress.XtraEditors.DateEdit();
			this.dtFechaFin = new DevExpress.XtraEditors.DateEdit();
			this.btChequeAntes = new DevExpress.XtraEditors.SimpleButton();
			this.btFiltro = new DevExpress.XtraEditors.SimpleButton();
			this.btExcel = new DevExpress.XtraEditors.SimpleButton();
			this.btConciliaAuto = new DevExpress.XtraEditors.SimpleButton();
			this.btAsiento = new DevExpress.XtraEditors.SimpleButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsEstadoCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEstadoCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsEstadoCuenta
			// 
			this.cdsEstadoCuenta.BindingContextCtrl = this;
			this.cdsEstadoCuenta.DataLibrary = "LibContabilidad";
			this.cdsEstadoCuenta.DataLibraryUrl = "";
			this.cdsEstadoCuenta.DataSetDef = "dsEstadoCuenta";
			this.cdsEstadoCuenta.FillOnRequest = false;
			this.cdsEstadoCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsEstadoCuenta.Name = "cdsEstadoCuenta";
			this.cdsEstadoCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsEstadoCuenta.SchemaDef = null;
			this.cdsEstadoCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsEstadoCuenta_BeforeFill);
			// 
			// cdvEstadoCuenta
			// 
			this.cdvEstadoCuenta.BindingContextCtrl = this;
			this.cdvEstadoCuenta.DataSet = this.cdsEstadoCuenta;
			this.cdvEstadoCuenta.TableName = "";
			this.cdvEstadoCuenta.TableViewName = "EstadoCuenta";
			// 
			// cdsCuentaCorr
			// 
			this.cdsCuentaCorr.BindingContextCtrl = this;
			this.cdsCuentaCorr.DataLibrary = "LibContabilidad";
			this.cdsCuentaCorr.DataLibraryUrl = "";
			this.cdsCuentaCorr.DataSetDef = "dsCuentaCorriente";
			this.cdsCuentaCorr.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuentaCorr.Name = "cdsCuentaCorr";
			this.cdsCuentaCorr.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuentaCorr.SchemaDef = null;
			this.cdsCuentaCorr.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuentaCorr_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 16);
			this.label1.TabIndex = 3;
			this.label1.Text = "Desde";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(192, 9);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 5;
			this.label2.Text = "Hasta:";
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(640, 457);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(56, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "Cuenta Nº";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvEstadoCuenta;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Conf.";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 54;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 53;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance3;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 53;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn5.Header.Caption = "Código";
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 65;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.Caption = "Descripción";
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Width = 73;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Format = "dd/MMM/yyyy";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 67;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Width = 73;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 63;
			ultraGridColumn10.Format = "dd/MMM/yyyy";
			ultraGridColumn10.Header.Caption = "F. Confirma";
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 67;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn16.Header.VisiblePosition = 17;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn17.Header.Caption = "Anulado";
			ultraGridColumn17.Header.VisiblePosition = 18;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 49;
			ultraGridColumn18.Header.VisiblePosition = 19;
			ultraGridColumn18.Width = 67;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn19.Header.VisiblePosition = 15;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn20.Header.VisiblePosition = 16;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 73;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 67;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 84;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Hidden = true;
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
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance7.BackColor = System.Drawing.Color.Gainsboro;
			appearance7.BackColor2 = System.Drawing.SystemColors.ControlLight;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 43);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(880, 395);
			this.ultraGrid1.TabIndex = 3;
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance10;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCuentaCorr, "CuentaCorriente.idCuenta"));
			this.cmbCuenta.DataMember = "CuentaCorriente";
			this.cmbCuenta.DataSource = this.cdsCuentaCorr;
			ultraGridColumn29.Header.VisiblePosition = 6;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 9;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 4;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 5;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn34.Header.VisiblePosition = 7;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 1;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 8;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 2;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 3;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 10;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 11;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 12;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuenta.DisplayMember = "NumCuenta";
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Location = new System.Drawing.Point(704, 457);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.ReadOnly = true;
			this.cmbCuenta.Size = new System.Drawing.Size(144, 22);
			this.cmbCuenta.TabIndex = 14;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuenta_InitializeLayout);
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "CuentaCorriente";
			this.barraDato1.DataNombreId = "idCuentaCorriente";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsCuentaCorr;
			this.barraDato1.DataTabla = "CuentaCorriente";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(8, 456);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(280, 25);
			this.barraDato1.TabIndex = 15;
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// chkFiltroFecha
			// 
			this.chkFiltroFecha.Checked = true;
			this.chkFiltroFecha.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkFiltroFecha.Location = new System.Drawing.Point(400, 9);
			this.chkFiltroFecha.Name = "chkFiltroFecha";
			this.chkFiltroFecha.Size = new System.Drawing.Size(128, 25);
			this.chkFiltroFecha.TabIndex = 17;
			this.chkFiltroFecha.Text = "Filtro Fecha/Conf";
			this.toolTip1.SetToolTip(this.chkFiltroFecha, "Filtra por fecha del asiento o fecha de Confirmacion");
			// 
			// btPreConci
			// 
			this.btPreConci.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btPreConci.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btPreConci.Location = new System.Drawing.Point(488, 456);
			this.btPreConci.Name = "btPreConci";
			this.btPreConci.Size = new System.Drawing.Size(96, 25);
			this.btPreConci.TabIndex = 21;
			this.btPreConci.Text = "PreConciliacion ";
			this.toolTip1.SetToolTip(this.btPreConci, "Pre Conciliacion ");
			this.btPreConci.Click += new System.EventHandler(this.btPreConci_Click);
			// 
			// chkNoConfirmado
			// 
			this.chkNoConfirmado.Checked = true;
			this.chkNoConfirmado.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkNoConfirmado.Location = new System.Drawing.Point(528, 9);
			this.chkNoConfirmado.Name = "chkNoConfirmado";
			this.chkNoConfirmado.Size = new System.Drawing.Size(160, 25);
			this.chkNoConfirmado.TabIndex = 19;
			this.chkNoConfirmado.Text = "Todos los No Confirmados";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// chkAnulado
			// 
			this.chkAnulado.Location = new System.Drawing.Point(688, 8);
			this.chkAnulado.Name = "chkAnulado";
			this.chkAnulado.Size = new System.Drawing.Size(80, 25);
			this.chkAnulado.TabIndex = 22;
			this.chkAnulado.Text = "Anulados";
			this.chkAnulado.CheckedChanged += new System.EventHandler(this.chkAnulado_CheckedChanged);
			// 
			// dtFechaIni
			// 
			this.dtFechaIni.EditValue = null;
			this.dtFechaIni.Location = new System.Drawing.Point(64, 8);
			this.dtFechaIni.Name = "dtFechaIni";
			// 
			// dtFechaIni.Properties
			// 
			this.dtFechaIni.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaIni.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaIni.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dtFechaIni.Properties.EditFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaIni.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dtFechaIni.Properties.Mask.EditMask = "dd/MMM/yyyy";
			this.dtFechaIni.Properties.NullDate = "";
			this.dtFechaIni.Size = new System.Drawing.Size(104, 20);
			this.dtFechaIni.TabIndex = 65;
			// 
			// dtFechaFin
			// 
			this.dtFechaFin.EditValue = null;
			this.dtFechaFin.Location = new System.Drawing.Point(240, 8);
			this.dtFechaFin.Name = "dtFechaFin";
			// 
			// dtFechaFin.Properties
			// 
			this.dtFechaFin.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																			 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.dtFechaFin.Properties.DisplayFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaFin.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dtFechaFin.Properties.EditFormat.FormatString = "dd/MMM/yyyy";
			this.dtFechaFin.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
			this.dtFechaFin.Properties.Mask.EditMask = "dd/MMM/yyyy";
			this.dtFechaFin.Properties.NullDate = "";
			this.dtFechaFin.Size = new System.Drawing.Size(104, 20);
			this.dtFechaFin.TabIndex = 66;
			this.dtFechaFin.EditValueChanged += new System.EventHandler(this.dtFechaFin_EditValueChanged);
			// 
			// btChequeAntes
			// 
			this.btChequeAntes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btChequeAntes.Location = new System.Drawing.Point(856, 456);
			this.btChequeAntes.Name = "btChequeAntes";
			this.btChequeAntes.Size = new System.Drawing.Size(24, 23);
			this.btChequeAntes.TabIndex = 67;
			this.btChequeAntes.Text = "...";
			this.btChequeAntes.ToolTip = "Ingreso de cheques del anio pasado";
			this.btChequeAntes.Click += new System.EventHandler(this.btChequeAntes_Click);
			// 
			// btFiltro
			// 
			this.btFiltro.Location = new System.Drawing.Point(824, 8);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(40, 23);
			this.btFiltro.TabIndex = 68;
			this.btFiltro.Text = "Filtro";
			this.btFiltro.ToolTip = "Filtro / Actualizacion de Fechas";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			this.btFiltro.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btFiltro_MouseDown);
			// 
			// btExcel
			// 
			this.btExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btExcel.Location = new System.Drawing.Point(344, 456);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(56, 23);
			this.btExcel.TabIndex = 69;
			this.btExcel.Text = "Excel";
			this.btExcel.ToolTip = "Concialiacion Automatica";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// btConciliaAuto
			// 
			this.btConciliaAuto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btConciliaAuto.Location = new System.Drawing.Point(312, 456);
			this.btConciliaAuto.Name = "btConciliaAuto";
			this.btConciliaAuto.Size = new System.Drawing.Size(24, 23);
			this.btConciliaAuto.TabIndex = 70;
			this.btConciliaAuto.Text = "C";
			this.btConciliaAuto.ToolTip = "Concialiacion Automatica";
			this.btConciliaAuto.Click += new System.EventHandler(this.btConciliaAuto_Click);
			// 
			// btAsiento
			// 
			this.btAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiento.Location = new System.Drawing.Point(288, 456);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(24, 23);
			this.btAsiento.TabIndex = 71;
			this.btAsiento.Text = "A";
			this.btAsiento.ToolTip = "Asiento de Fila Seleccionada";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// EstadoCuenta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(896, 496);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btConciliaAuto);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.btChequeAntes);
			this.Controls.Add(this.dtFechaFin);
			this.Controls.Add(this.dtFechaIni);
			this.Controls.Add(this.chkAnulado);
			this.Controls.Add(this.btPreConci);
			this.Controls.Add(this.chkNoConfirmado);
			this.Controls.Add(this.chkFiltroFecha);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "EstadoCuenta";
			this.Text = "Estado de Cuenta";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.EstadoCuenta_Closing);
			this.Load += new System.EventHandler(this.EstadoCuente_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsEstadoCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvEstadoCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void EstadoCuente_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Estado Cuenta'";
			Funcion.EjecutaSQL(cdsCuentaCorr, stAudita, true);
			if (Funcion.bEjecutaSQL(cdsCuentaCorr, "Exec SeteoGExiste 'ACGECB'"))
			{
				ultraGrid1.DisplayLayout.Bands[0].Columns["Deposito"].Hidden = true;
				ultraGrid1.DisplayLayout.Bands[0].Columns["FechaRef"].Hidden = true;
				ultraGrid1.DisplayLayout.Bands[0].Columns["FechaComp"].Hidden = true;
				ultraGrid1.DisplayLayout.Bands[0].Columns["Codigo"].Hidden = true;
			}
			//Cambio para que aparesca el boton de preconciliacion si esta activo el codigo
			bool bActiva = Funcion.bEjecutaSQL(cdsCuentaCorr, "Exec SeteoGExiste 'ASECPC'");
			if(bActiva == true) btPreConci.Visible=true;
			///

			string stSelect = "Delete AsientoDetalle from AsientoDetalle "
				+ "left outer join Asiento On AsientoDetalle.idAsiento "
				+ "= Asiento.idAsiento Where IsNull(Asiento.idAsiento, -1) = -1";
		
			Funcion.EjecutaSQL(cdsCuentaCorr, stSelect);

			miAcceso =new Acceso("0416", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);

			DateTime dtFecha1 =DateTime.Today;
			dtFecha1 = dtFecha1.AddMonths(-1);
			dtFecha1 = dtFecha1.AddDays(1 - dtFecha1.Day);
			dtFechaIni.EditValue = dtFecha1;
			dtFechaFin.EditValue = dtFecha1.AddMonths(1).AddDays(-1);
			
			barraDato1.PosUltima();
			barraDato1.HabilitaControles(false);
			ModificaGrilla(false);
			chkNoConfirmado.Enabled = true;
			chkAnulado.Enabled = true;
			Cursor = Cursors.Default;
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}


		private void cdsEstadoCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsEstadoCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);		
		}

		private void cdsCuentaCorr_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuentaCorr.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void Filtro()
		{
			chkFiltroFecha.Enabled = true;
			string stFiltro = "";
			if (cmbCuenta.Value != null && cmbCuenta.Value != System.DBNull.Value && (int) cmbCuenta.Value > 0)
				stFiltro = "idCuenta = " + this.cmbCuenta.Value.ToString();
			else
			{
				MessageBox.Show("Filtro no aplicable");
				return;
			}
//			string stConfirmado = "'))";
			string stTipoFecha = "Fecha";

			DateTime dtIni = (DateTime) this.dtFechaIni.EditValue;
			DateTime dtFin = (DateTime) this.dtFechaFin.EditValue;
			if (!chkFiltroFecha.Checked)
			{
				stTipoFecha = "FechaConf";
			}
			else
			{
				string stFechaLinea = "Select IsNull(FechaLinea, 0) From SeteoC";
				bool bFechaLinea = Funcion.bEscalarSQL(cdsCuentaCorr, stFechaLinea, true);
				if (bFechaLinea == true)
				{
					stTipoFecha = "FechaComp";
				}
			}

			if (!chkAnulado.Checked) stFiltro += " And Asiento.Borrar = 0 ";

			stFiltro += " And ((" + stTipoFecha + " >= '" + dtIni.ToString("yyyyMMdd")
				+ "' And " + stTipoFecha + " < '" + dtFin.AddDays(1).ToString("yyyyMMdd");
			if (chkNoConfirmado.Checked)
				stFiltro += "') Or (Confirmado = 0 And " + stTipoFecha + " < '" + dtFin.AddDays(1).ToString("yyyyMMdd") + "'))";
			else
				stFiltro += "'))";

			try
			{
				C1.Data.FilterCondition fcPadre = new 
					C1.Data.FilterCondition(cdsEstadoCuenta, "EstadoCuenta", stFiltro);
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcPadre);

				cdsEstadoCuenta.Fill(fcTotal, false);
				errorProvider.SetError(dtFechaFin, "");
				if (this.cdvEstadoCuenta.Count == 0)
				{
					errorProvider.SetError(dtFechaFin, "No existen registros en ese periodo");
					return;
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error en Filtro");
			}
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Filtrar Articulos", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Filtro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			this.cdsEstadoCuenta.Update();
			this.barraDato1.Edicion(false);
			ModificaGrilla(false);
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir Estado de Cuenta",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			EstadoCuentaImprime miImprime = new EstadoCuentaImprime();
			miImprime.MdiParent = this.MdiParent;
			miImprime.dtDesde.EditValue = this.dtFechaIni.EditValue;
			miImprime.dtHasta.EditValue = this.dtFechaFin.EditValue;
			miImprime.cmbCuentaCorr.Value = this.cmbCuenta.Value;
			miImprime.Show();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar Estado de Cuenta",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDato1.EditarRegistro();
			ModificaGrilla(true);
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			MessageBox.Show("Para crear una nueva cuenta use la opción\n Contabilidad Cuentas");
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
			this.cdsEstadoCuenta.RejectChanges();
			ModificaGrilla(false);
		}

		private void btChequeAntes_Click(object sender, System.EventArgs e)
		{
			EstadoCuentaAntes miEstado = new EstadoCuentaAntes((int) cmbCuenta.Value);
			miEstado.MdiParent = this.MdiParent;
			miEstado.Show();
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.Cancel = true;
			MessageBox.Show("No es posible borrar lineas del Estado de Cuenta", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			Filtro();
		}

		private void btFiltro_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				if (DialogResult.No == MessageBox.Show("Desea actualizar fechas de confirmacion con fecha de asiento en transacciones no confirmadas?",
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
				string stSelect = "Update AsientoDetalle Set AsientoDetalle.FechaConf = "
					+ "Asiento.Fecha From Asiento Inner Join AsientoDetalle "
					+ "On AsientoDetalle.idAsiento = Asiento.idAsiento "
					+ "Where AsientoDetalle.idCuenta = " + cmbCuenta.Value.ToString()
					+ " And AsientoDetalle.Confirmado = 0 ";
				Funcion.EjecutaSQL(cdsEstadoCuenta, stSelect, true);
				barraDato1.ProximoId(5);
				MessageBox.Show("Fechas de Confirmacion Actualizadas", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null) return;
			Asiento miAsiento = new Asiento((int) ultraGrid1.ActiveRow.Cells["idAsiento"].Value);
			miAsiento.MdiParent = this.MdiParent;
			miAsiento.Show();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (BuscaEstadoCuenta miBusqueda = new BuscaEstadoCuenta((int) cmbCuenta.Value))
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					if (miBusqueda.idBusca.Text.Trim().Length == 0)
					{
						return;
					}
					int IdDetAsi = int.Parse(miBusqueda.idBusca.Text);
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						if (dr.Cells["idDetAsiento"].Value != System.DBNull.Value
							&& (int) dr.Cells["idDetAsiento"].Value == IdDetAsi)
						{
							ultraGrid1.ActiveRow = dr;
							return;
						}
					}
				}
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		private void btConciliaAuto_Click(object sender, System.EventArgs e)
		{
			int IdCuenta = (int) cmbCuenta.Value;
			EstadoCuentaAuto miEstado = new EstadoCuentaAuto(IdCuenta);
			miEstado.MdiParent = this.MdiParent;
			miEstado.Show();
		}

		private void EstadoCuenta_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Estado Cuenta'";
			Funcion.EjecutaSQL(cdsCuentaCorr, stAudita, true);
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\";
			ultraGridExcelExporter1.Export(ultraGrid1, "EstadoCuenta.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "EstadoCuenta.xls");
		}

		private void btPreConci_Click(object sender, System.EventArgs e)
		{
			string stidCuenta = "";
			if (cmbCuenta.Value != null && cmbCuenta.Value != System.DBNull.Value && (int) cmbCuenta.Value > 0)
				stidCuenta =  this.cmbCuenta.Value.ToString();

			P_EstadoCuentaAlterno miEstado = new P_EstadoCuentaAlterno(stidCuenta);
			miEstado.MdiParent = this.MdiParent;
			miEstado.Show();

		}

		private void chkAnulado_CheckedChanged(object sender, System.EventArgs e)
		{
			ultraGrid1.DisplayLayout.Bands[0].Columns["Borrar"].Hidden = !chkAnulado.Checked;		
		}

		private void dtFechaFin_EditValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbCuenta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

	}
}
