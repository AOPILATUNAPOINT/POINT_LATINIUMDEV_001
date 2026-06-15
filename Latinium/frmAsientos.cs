using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmAsientos.
	/// </summary>
	public class frmAsientos : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label lblCheque;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox5;
		private C1.Data.C1DataSet cdsSeteoC;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbOrigen;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroDeCosto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCentroDeCostoL;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRUC;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtChequeDeposito;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDctoOrigen;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHaber;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDebe;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label lblDescuadrado;
		private System.Windows.Forms.Label lblPeriodo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnDistribucion;
		private System.Windows.Forms.ContextMenu mnuDistribucion;
		private System.Windows.Forms.MenuItem mnuAgencias;
		private System.Windows.Forms.MenuItem mnuAdministrativos;
		private System.Windows.Forms.MenuItem mnuGeneral;
		private System.Windows.Forms.Button btnDuplicar;
		private System.Windows.Forms.MenuItem mnuAgenciasIg;
		private System.Windows.Forms.MenuItem mnuAdministrativosIg;
		private System.Windows.Forms.MenuItem mnuGeneralIg;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmAsientos(int idAsiento)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			this.txtIdAsiento.Value = idAsiento;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmAsientos()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAsientos));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("AsientoTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("AsientoTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigenAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigenAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbCentroDeCostoL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFactura");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debe", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debe", 8, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Haber", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Haber", 9, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Haber");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFactura");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			this.label10 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.lblCheque = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCentroDeCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbOrigen = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbCentroDeCostoL = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtHaber = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDebe = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRUC = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtChequeDeposito = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDctoOrigen = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblDescuadrado = new System.Windows.Forms.Label();
			this.lblPeriodo = new System.Windows.Forms.Label();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnDistribucion = new System.Windows.Forms.Button();
			this.mnuDistribucion = new System.Windows.Forms.ContextMenu();
			this.mnuAgencias = new System.Windows.Forms.MenuItem();
			this.mnuAdministrativos = new System.Windows.Forms.MenuItem();
			this.mnuGeneral = new System.Windows.Forms.MenuItem();
			this.mnuAgenciasIg = new System.Windows.Forms.MenuItem();
			this.mnuAdministrativosIg = new System.Windows.Forms.MenuItem();
			this.mnuGeneralIg = new System.Windows.Forms.MenuItem();
			this.btnDuplicar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCostoL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHaber)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDebe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRUC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtChequeDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDctoOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.ForeColor = System.Drawing.Color.Black;
			this.label10.Location = new System.Drawing.Point(472, 43);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(36, 16);
			this.label10.TabIndex = 149;
			this.label10.Text = "R.U.C.";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(224, 106);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(85, 16);
			this.label9.TabIndex = 147;
			this.label9.Text = "Centro de Costo";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 106);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(34, 16);
			this.label5.TabIndex = 146;
			this.label5.Text = "Fecha";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCheque
			// 
			this.lblCheque.AutoSize = true;
			this.lblCheque.Location = new System.Drawing.Point(8, 75);
			this.lblCheque.Name = "lblCheque";
			this.lblCheque.Size = new System.Drawing.Size(56, 16);
			this.lblCheque.TabIndex = 142;
			this.lblCheque.Text = "Cheque N.";
			this.lblCheque.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(776, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(44, 16);
			this.label6.TabIndex = 141;
			this.label6.Text = "Numero";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 365);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(33, 16);
			this.label7.TabIndex = 140;
			this.label7.Text = "Notas";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.ForeColor = System.Drawing.Color.Black;
			this.label8.Location = new System.Drawing.Point(8, 43);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(61, 16);
			this.label8.TabIndex = 139;
			this.label8.Text = "Descripción";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.WindowsXPCommandButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(96, 104);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 9;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// cmbTipo
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance2;
			this.cmbTipo.CausesValidation = false;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 136;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipo.DisplayMember = "Tipo";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Enabled = false;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(64, 8);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(136, 21);
			this.cmbTipo.TabIndex = 1;
			this.cmbTipo.ValueMember = "idTipoAsiento";
			this.cmbTipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipo_KeyDown);
			this.cmbTipo.Validating += new System.ComponentModel.CancelEventHandler(this.cmbTipo_Validating);
			this.cmbTipo.ValueChanged += new System.EventHandler(this.cmbTipo_ValueChanged);
			this.cmbTipo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipo_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			this.ultraDataSource1.Band.Key = "AsientoTipo";
			// 
			// cmbCentroDeCosto
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCentroDeCosto.Appearance = appearance3;
			this.cmbCentroDeCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroDeCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroDeCosto.DataSource = this.ultraDataSource3;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 250;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCentroDeCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCentroDeCosto.DisplayMember = "Nombre";
			this.cmbCentroDeCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroDeCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroDeCosto.Enabled = false;
			this.cmbCentroDeCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroDeCosto.Location = new System.Drawing.Point(336, 104);
			this.cmbCentroDeCosto.Name = "cmbCentroDeCosto";
			this.cmbCentroDeCosto.Size = new System.Drawing.Size(250, 21);
			this.cmbCentroDeCosto.TabIndex = 10;
			this.cmbCentroDeCosto.ValueMember = "idProyecto";
			this.cmbCentroDeCosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCentroDeCosto_KeyDown);
			this.cmbCentroDeCosto.Validating += new System.ComponentModel.CancelEventHandler(this.cmbCentroDeCosto_Validating);
			this.cmbCentroDeCosto.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCentroDeCosto_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(26, 16);
			this.label1.TabIndex = 150;
			this.label1.Text = "Tipo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(224, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 16);
			this.label3.TabIndex = 155;
			this.label3.Text = "Origen";
			// 
			// cmbOrigen
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbOrigen.Appearance = appearance4;
			this.cmbOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbOrigen.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbOrigen.DataSource = this.ultraDataSource2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbOrigen.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbOrigen.DisplayMember = "Descripcion";
			this.cmbOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbOrigen.Enabled = false;
			this.cmbOrigen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbOrigen.Location = new System.Drawing.Point(288, 8);
			this.cmbOrigen.Name = "cmbOrigen";
			this.cmbOrigen.Size = new System.Drawing.Size(360, 21);
			this.cmbOrigen.TabIndex = 2;
			this.cmbOrigen.ValueMember = "idOrigenAsiento";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			this.ultraDataSource2.Band.Key = "AsientoTipo";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.ForeColor = System.Drawing.Color.Black;
			this.label4.Location = new System.Drawing.Point(224, 75);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 16);
			this.label4.TabIndex = 156;
			this.label4.Text = "Dcto de Origen";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(715, 70);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 241;
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource4;
			appearance5.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance5;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 70;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 61;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 57;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Width = 104;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Width = 176;
			ultraGridColumn12.Header.Caption = "Centro de Costo";
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn12.Width = 126;
			ultraGridColumn13.Header.Caption = "Forma de Pago";
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn13.Width = 91;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Width = 235;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance6;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Width = 86;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance7;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 9;
			ultraGridColumn16.Width = 105;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 10;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 11;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 12;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 13;
			ultraGridColumn20.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn20});
			ultraGridBand4.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance8;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance9;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand4.SummaryFooterCaption = "Totales";
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance11;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8.5F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 136);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(944, 192);
			this.uGridDetalle.TabIndex = 11;
			this.uGridDetalle.AfterRowsDeleted += new System.EventHandler(this.uGridDetalle_AfterRowsDeleted);
			this.uGridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalle_AfterRowInsert);
			this.uGridDetalle.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridDetalle_KeyPress);
			this.uGridDetalle.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDetalle_BeforeCellUpdate);
			this.uGridDetalle.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalle_BeforeRowsDeleted);
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 432);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 0;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(83, 432);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 14;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(164, 432);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 15;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(243, 432);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 16;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(323, 432);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 13;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(481, 432);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 18;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(402, 432);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 17;
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(556, 432);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 19;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox5
			// 
			this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox5.Location = new System.Drawing.Point(8, 416);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(944, 8);
			this.groupBox5.TabIndex = 420;
			this.groupBox5.TabStop = false;
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			this.cdsSeteoC.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoC_BeforeFill);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(928, 432);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 429;
			this.txtIdAsiento.Visible = false;
			// 
			// cmbCentroDeCostoL
			// 
			this.cmbCentroDeCostoL.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroDeCostoL.DataSource = this.ultraDataSource5;
			ultraGridBand5.AddButtonCaption = "Proyecto";
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Width = 250;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbCentroDeCostoL.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCentroDeCostoL.DisplayMember = "Nombre";
			this.cmbCentroDeCostoL.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroDeCostoL.Location = new System.Drawing.Point(48, 200);
			this.cmbCentroDeCostoL.Name = "cmbCentroDeCostoL";
			this.cmbCentroDeCostoL.Size = new System.Drawing.Size(272, 100);
			this.cmbCentroDeCostoL.TabIndex = 430;
			this.cmbCentroDeCostoL.ValueMember = "idProyecto";
			this.cmbCentroDeCostoL.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// txtHaber
			// 
			this.txtHaber.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtHaber.Appearance = appearance16;
			this.txtHaber.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHaber.Enabled = false;
			this.txtHaber.FormatString = "#,##0.00";
			this.txtHaber.Location = new System.Drawing.Point(864, 340);
			this.txtHaber.Name = "txtHaber";
			this.txtHaber.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtHaber.PromptChar = ' ';
			this.txtHaber.Size = new System.Drawing.Size(90, 22);
			this.txtHaber.TabIndex = 432;
			// 
			// txtDebe
			// 
			this.txtDebe.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDebe.Appearance = appearance17;
			this.txtDebe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDebe.Enabled = false;
			this.txtDebe.FormatString = "#,##0.00";
			this.txtDebe.Location = new System.Drawing.Point(768, 340);
			this.txtDebe.Name = "txtDebe";
			this.txtDebe.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDebe.PromptChar = ' ';
			this.txtDebe.Size = new System.Drawing.Size(90, 22);
			this.txtDebe.TabIndex = 431;
			// 
			// txtDescripcion
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance18;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(96, 40);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(352, 22);
			this.txtDescripcion.TabIndex = 4;
			this.txtDescripcion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescripcion_KeyDown);
			this.txtDescripcion.Validating += new System.ComponentModel.CancelEventHandler(this.txtDescripcion_Validating);
			this.txtDescripcion.ValueChanged += new System.EventHandler(this.txtDescripcion_ValueChanged);
			// 
			// txtNumero
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance19;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(840, 7);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 3;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// txtRUC
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRUC.Appearance = appearance20;
			this.txtRUC.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRUC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRUC.Enabled = false;
			this.txtRUC.Location = new System.Drawing.Point(536, 40);
			this.txtRUC.Name = "txtRUC";
			this.txtRUC.Size = new System.Drawing.Size(112, 22);
			this.txtRUC.TabIndex = 5;
			// 
			// txtChequeDeposito
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtChequeDeposito.Appearance = appearance21;
			this.txtChequeDeposito.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtChequeDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtChequeDeposito.Enabled = false;
			this.txtChequeDeposito.Location = new System.Drawing.Point(96, 72);
			this.txtChequeDeposito.Name = "txtChequeDeposito";
			this.txtChequeDeposito.Size = new System.Drawing.Size(104, 22);
			this.txtChequeDeposito.TabIndex = 6;
			this.txtChequeDeposito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtChequeDeposito_KeyDown);
			this.txtChequeDeposito.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtChequeDeposito_KeyPress);
			this.txtChequeDeposito.Validating += new System.ComponentModel.CancelEventHandler(this.txtChequeDeposito_Validating);
			this.txtChequeDeposito.ValueChanged += new System.EventHandler(this.txtChequeDeposito_ValueChanged);
			// 
			// txtDctoOrigen
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDctoOrigen.Appearance = appearance22;
			this.txtDctoOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDctoOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDctoOrigen.Enabled = false;
			this.txtDctoOrigen.Location = new System.Drawing.Point(336, 72);
			this.txtDctoOrigen.Name = "txtDctoOrigen";
			this.txtDctoOrigen.Size = new System.Drawing.Size(112, 22);
			this.txtDctoOrigen.TabIndex = 7;
			// 
			// txtNotas
			// 
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance23;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(72, 340);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(600, 66);
			this.txtNotas.TabIndex = 433;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// lblDescuadrado
			// 
			this.lblDescuadrado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuadrado.AutoSize = true;
			this.lblDescuadrado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuadrado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblDescuadrado.Location = new System.Drawing.Point(715, 103);
			this.lblDescuadrado.Name = "lblDescuadrado";
			this.lblDescuadrado.Size = new System.Drawing.Size(0, 22);
			this.lblDescuadrado.TabIndex = 434;
			// 
			// lblPeriodo
			// 
			this.lblPeriodo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblPeriodo.AutoSize = true;
			this.lblPeriodo.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPeriodo.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPeriodo.Location = new System.Drawing.Point(712, 380);
			this.lblPeriodo.Name = "lblPeriodo";
			this.lblPeriodo.Size = new System.Drawing.Size(0, 26);
			this.lblPeriodo.TabIndex = 435;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource6;
			ultraGridBand6.AddButtonCaption = "CompraFormaPago";
			ultraGridBand6.ColHeadersVisible = false;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.Caption = "";
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 186;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(386, 188);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(176, 86);
			this.cmbFormaPago.TabIndex = 436;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
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
			// btnDistribucion
			// 
			this.btnDistribucion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnDistribucion.CausesValidation = false;
			this.btnDistribucion.Enabled = false;
			this.btnDistribucion.Image = ((System.Drawing.Image)(resources.GetObject("btnDistribucion.Image")));
			this.btnDistribucion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDistribucion.Location = new System.Drawing.Point(840, 40);
			this.btnDistribucion.Name = "btnDistribucion";
			this.btnDistribucion.Size = new System.Drawing.Size(112, 24);
			this.btnDistribucion.TabIndex = 437;
			this.btnDistribucion.Text = "Distribucion";
			this.btnDistribucion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDistribucion.Click += new System.EventHandler(this.btnDistribucion_Click);
			// 
			// mnuDistribucion
			// 
			this.mnuDistribucion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																										this.mnuAgencias,
																																										this.mnuAdministrativos,
																																										this.mnuGeneral,
																																										this.mnuAgenciasIg,
																																										this.mnuAdministrativosIg,
																																										this.mnuGeneralIg});
			// 
			// mnuAgencias
			// 
			this.mnuAgencias.Index = 0;
			this.mnuAgencias.Text = "Agencias";
			this.mnuAgencias.Click += new System.EventHandler(this.mnuAgencias_Click);
			// 
			// mnuAdministrativos
			// 
			this.mnuAdministrativos.Index = 1;
			this.mnuAdministrativos.Text = "Administrativos";
			this.mnuAdministrativos.Click += new System.EventHandler(this.mnuAdministrativos_Click);
			// 
			// mnuGeneral
			// 
			this.mnuGeneral.Index = 2;
			this.mnuGeneral.Text = "General";
			this.mnuGeneral.Click += new System.EventHandler(this.mnuGeneral_Click);
			// 
			// mnuAgenciasIg
			// 
			this.mnuAgenciasIg.Index = 3;
			this.mnuAgenciasIg.Text = "Agencias - Igualdad";
			this.mnuAgenciasIg.Click += new System.EventHandler(this.mnuAgenciasIg_Click);
			// 
			// mnuAdministrativosIg
			// 
			this.mnuAdministrativosIg.Index = 4;
			this.mnuAdministrativosIg.Text = "Administrativos-Igualdad";
			this.mnuAdministrativosIg.Click += new System.EventHandler(this.mnuAdministrativosIg_Click);
			// 
			// mnuGeneralIg
			// 
			this.mnuGeneralIg.Index = 5;
			this.mnuGeneralIg.Text = "General - Igualdad";
			this.mnuGeneralIg.Click += new System.EventHandler(this.mnuGeneralIg_Click);
			// 
			// btnDuplicar
			// 
			this.btnDuplicar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnDuplicar.Enabled = false;
			this.btnDuplicar.Image = ((System.Drawing.Image)(resources.GetObject("btnDuplicar.Image")));
			this.btnDuplicar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDuplicar.Location = new System.Drawing.Point(648, 432);
			this.btnDuplicar.Name = "btnDuplicar";
			this.btnDuplicar.Size = new System.Drawing.Size(76, 23);
			this.btnDuplicar.TabIndex = 438;
			this.btnDuplicar.Text = "&Duplicar";
			this.btnDuplicar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDuplicar.Click += new System.EventHandler(this.btnDuplicar_Click);
			// 
			// frmAsientos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(959, 462);
			this.Controls.Add(this.btnDuplicar);
			this.Controls.Add(this.btnDistribucion);
			this.Controls.Add(this.lblPeriodo);
			this.Controls.Add(this.lblDescuadrado);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txtDctoOrigen);
			this.Controls.Add(this.txtChequeDeposito);
			this.Controls.Add(this.txtRUC);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.txtHaber);
			this.Controls.Add(this.txtDebe);
			this.Controls.Add(this.cmbCentroDeCostoL);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.lblCheque);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.cmbOrigen);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbCentroDeCosto);
			this.Controls.Add(this.cmbFormaPago);
			this.KeyPreview = true;
			this.Name = "frmAsientos";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Asientos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAsientos_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmAsientos_Closing);
			this.Load += new System.EventHandler(this.frmAsientos_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCostoL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHaber)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDebe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRUC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtChequeDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDctoOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		int iEstado = 0;
		bool bEliminaAlValidar = false;
		public Acceso miAcceso;
		private bool bActualiza = true;
		bool bDistribucion = false;
		private void CargaRegistro(int idAsiento)
		{
			try
			{	
				#region Botones
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;				
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;				
				this.btnAnular.Text = "Anular";				
				#endregion Botones
				
				#region Consulta
				string sSQL = string.Format("Exec AsientoBuscaIndividual {0}", idAsiento);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoC, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdAsiento.Value = idAsiento;
					this.cmbTipo.Value = dr.GetInt32(1);	
					this.cmbOrigen.Value = dr.GetInt32(2);
					if (dr.GetValue(3) != System.DBNull.Value) this.txtNumero.Text = dr.GetString(3);
					if (dr.GetValue(4) != System.DBNull.Value) this.txtDescripcion.Text = dr.GetString(4);
					this.txtRUC.Text = dr.GetString(5);
					this.txtChequeDeposito.Text = dr.GetString(6);
					this.txtDctoOrigen.Text = dr.GetString(7);
					this.dtFecha.Value = dr.GetDateTime(8);
					this.cmbCentroDeCosto.Value = dr.GetInt32(9); 
					this.txtNotas.Text = dr.GetString(10);
					this.txtDebe.Value = dr.GetDecimal(11);
					this.txtHaber.Value = dr.GetDecimal(12);
					bAnulado = dr.GetBoolean(13);
					iEstado = dr.GetInt32(14); 
				}
				dr.Close();
				
				this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoC, string.Format("Exec AsientoConsultaDetalle {0}", idAsiento));
				#endregion Consulta
				
				#region Controles
				if (!bAnulado)
				{
					if (iEstado == 1) this.lblEstado.Text = "PENDIENTE";
					if (iEstado == 2) this.lblEstado.Text = "CONTABILIZADO";       
				 
					if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoC, "Contabilidad")  )
					{
						if (iEstado == 1)
						{
							if (miAcceso.Editar) 
							{
								if (miAcceso.EditarDiaIngreso)
								{
									DateTime dtFechaAsi = Funcion.dtEscalarSQL(cdsSeteoC, "Select GetDate()");
									DateTime dtFechaIng = (DateTime) this.dtFecha.Value;
									if (dtFechaAsi.ToShortDateString() == dtFechaIng.ToShortDateString()) this.btnEditar.Enabled = true;
								}
								else this.btnEditar.Enabled = true;		
							}
						}
						if (miAcceso.Anular) this.btnAnular.Enabled = true;						
					}
					else this.lblPeriodo.Text = "PERIODO CERRADO";		

					if (miAcceso.Anular) this.btnAnular.Enabled = true;		

					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
				}
				else
				{
					this.lblEstado.Text = "ANULADO";
					this.btnAnular.Text = "Eliminar";
					if (miAcceso.Eliminar) this.btnAnular.Enabled = true;						
				}
					
				this.btnCancelar.Enabled = true;
				this.btnDuplicar.Enabled = true;
				#endregion Controles
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void Listas ()
		{
			this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoC, "Select idTipoAsiento, Tipo From AsientoTipo Where Activo = 1 Order By Tipo");
			this.cmbOrigen.DataSource = Funcion.dvProcedimiento(cdsSeteoC, "Select idOrigenAsiento, Descripcion From OrigenAsientos Order By Descripcion");
			this.cmbCentroDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoC, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order By Nombre");
			this.cmbCentroDeCostoL.DataSource = Funcion.dvProcedimiento(cdsSeteoC, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order By Nombre");
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoC, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 4, 7, 13, 16, 19)");
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmAsientos_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoC, "0402");

			if (!Funcion.Permiso("44",cdsSeteoC))
			{				
				MessageBox.Show("No puede ingresar a Asientos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				Listas();

				FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);

				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			
				if ((int)this.txtIdAsiento.Value > 0) this.CargaRegistro((int)this.txtIdAsiento.Value);				
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsSeteoC_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoC.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;
			
			this.txtIdAsiento.Value = 0;
			this.cmbTipo.Value = System.DBNull.Value;
			this.cmbOrigen.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.txtDescripcion.Text = "";
			this.txtRUC.Text = "";
			this.lblCheque.Text = "Cheque N.";
			this.txtChequeDeposito.Text = "";
			this.txtDctoOrigen.Text = "";
			this.dtFecha.Value = System.DBNull.Value;
			this.cmbCentroDeCosto.Value = System.DBNull.Value;
			this.lblEstado.Text = "";
			this.lblDescuadrado.Text = "";
			this.lblPeriodo.Text = "";
			this.txtNotas.Text = "";
			this.txtDebe.Value = 0;
			this.txtHaber.Value = 0;

			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoC, "Exec AsientoConsultaDetalle 0");

			this.cmbTipo.Enabled = false;			
			this.txtNumero.Enabled = false;
			this.txtDescripcion.Enabled = false;
			this.txtRUC.Enabled = false;
			this.txtChequeDeposito.Enabled = false;
			this.txtDctoOrigen.Enabled = false;
			this.dtFecha.Enabled = false;
			this.cmbCentroDeCosto.Enabled = false;			
			this.txtNotas.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;

			bAnulado = false;
			bEliminaAlValidar = false;

			this.btnAnular.Text = "Anular";
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.cmbTipo.Enabled = true;
			this.txtDescripcion.Enabled = true;			
			//this.dtFecha.Enabled = true;
			//if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			DateTime dtHoy = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);	
			if (miAcceso.BFecha3Días) 
			{
				DateTime dtFechaVal1 = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT DATEADD(DAY,3,EOMONTH('{0}',-1))",Convert.ToDateTime(dtHoy).ToString("yyyyMMdd")));
				DateTime dtFechaVal2 = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT EOMONTH('{0}',-1)",Convert.ToDateTime(dtHoy).ToString("yyyyMMdd")));	
				int iCan= Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDiasCompas '{0}','{1}'", Convert.ToDateTime(dtFechaVal2).ToString("yyyyMMdd"),Convert.ToDateTime(dtHoy).ToString("yyyyMMdd")));
				if(iCan<4)
				{
					this.dtFecha.Enabled = true;
					this.dtFecha.CalendarInfo.MinDate = dtFechaVal2;
					this.dtFecha.CalendarInfo.MaxDate = DateTime.Today.AddDays(iCan);
				}
			}

			/* si tengo el acceso a cambiar fecha pero en didiembre solo dieicmebre llega */
			if (miAcceso.CambiarFecha)
			{

				if (miAcceso.CambiarFechaDiciembre)
				{
					this.dtFecha.Enabled = true;
					this.dtFecha.CalendarInfo.MinDate = new DateTime(2024, 12, 1);
       
					this.dtFecha.CalendarInfo.MaxDate = dtHoy;
				}
				else
				{
					this.dtFecha.Enabled = true;
				}
			}

			this.cmbCentroDeCosto.Enabled = true;			
			this.txtNotas.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			//this.btnAnular.Enabled = true;
			this.btnCancelar.Enabled = true;
			if (miAcceso.BtnDistribucion) this.btnDistribucion.Enabled = true;

			this.cmbTipo.Value = 1;
			
			//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//this.dtFecha.Value = DateTime.Today;this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoC, false);
			//this.dtFecha.CalendarInfo.MinDate=DateTime.Parse("28/02/2022");

			this.cmbOrigen.Value = 0;
			this.cmbCentroDeCosto.Value = 19;

			this.txtDescripcion.Focus();

			bNuevo = true;
			bEdicion = true;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.CargaRegistro((int)this.txtIdAsiento.Value);

			if (bAnulado)
			{
				MessageBox.Show("No puede Editar, El Registro esta ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
				return;
			}

			if (iEstado == 2)
			{
				MessageBox.Show("No puede Editar, El Registro esta CONTABILIZADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
				return;
			}

			if (this.ValidaMovimientoConciliado("EDITAR")) return;
			if (!FuncionesContabilidad.bValidaEdicionAnulacionAsientos(this.cmbOrigen, 2, cdsSeteoC)) return;

			int IdOrigen = 0;
			if (this.cmbOrigen.ActiveRow != null) IdOrigen = (int)this.cmbOrigen.Value;
			
			this.txtDescripcion.Enabled = true;									
			this.txtNotas.Enabled = true;
			
			if (IdOrigen == 0) 
			{
				this.cmbTipo.Enabled = true;
				this.txtChequeDeposito.Enabled = true;
				if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;				
				this.cmbCentroDeCosto.Enabled = true;				

				FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);				
			}
		
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;

			this.txtDescripcion.Focus();

			bNuevo = false;
			bEdicion = true;			
		}

		private bool ValidaMovimientoConciliado (string sAccion)
		{
			bool bRes = false;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
			{	
				string sSQL = string.Format("Exec ValidaMovimientoBancarioConciliado {0}, {1}", (int)dr.Cells["idDetAsiento"].Value, (int)dr.Cells["idCuenta"].Value);
				string sNumero = Funcion.sEscalarSQL(cdsSeteoC, sSQL);

				if (sNumero.Length > 0)
				{
					MessageBox.Show(string.Format("No puede '{1}', El movimiento se encuentra registrado en la Conciliacion N. '{0}'", sNumero, sAccion), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					bRes = true;
				}
			}

			return bRes;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
		
			using (frmBusquedaDeAsientos Busqueda = new frmBusquedaDeAsientos())
				if (DialogResult.OK == Busqueda.ShowDialog())
					this.CargaRegistro((int)Busqueda.uGridAsientos.ActiveRow.Cells["idAsiento"].Value);
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			
			int IdTipo = (int)this.cmbTipo.Value;
			int iTipoImp = 0;
			string strReporte = "";			
			string strFiltro = "{Asiento.idAsiento} = " + this.txtIdAsiento.Value.ToString();

			switch (IdTipo)
			{
				case 1:					
				case 2:
				case 4:
				case 5:
					strReporte = "Asientos.rpt";					
					break;
				case 3:
					if (MenuLatinium.stDirFacturacion == "CROMATIKA")	strReporte = "ComprobanteDeEgresoCromatika.rpt";
					else if (MenuLatinium.stDirFacturacion == "ZOOMMATIKA")	strReporte = "ComprobanteDeEgresoCromatika.rpt";
					else if (MenuLatinium.stDirFacturacion == "LOTESERVICIO") strReporte = "ComprobanteDeEgreso.rpt";
					else if (MenuLatinium.stDirFacturacion == "GRETASANTORINI") strReporte = "ComprobanteDeEgresoA4.rpt";
					else if (MenuLatinium.stDirFacturacion == "IWANTYOU") strReporte = "ComprobanteDeEgreso.rpt";					
					else strReporte = "ComprobanteDeEgresoA4.rpt";
					break;					
			}

			if ((int) this.cmbTipo.Value == 3)
			{
				#region Egresos
				int idTipo = Funcion.iEscalarSQL(cdsSeteoC, string.Format("Exec TipoAutorizacionAsiento {0}", (int)this.txtIdAsiento.Value));

				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfidAsiento = new ParameterField ();
				ParameterDiscreteValue ValIdAsiento = new ParameterDiscreteValue ();
				pfidAsiento.ParameterFieldName = "@idAsiento";
				ValIdAsiento.Value =  (int)this.txtIdAsiento.Value;
				pfidAsiento.CurrentValues.Add (ValIdAsiento);
				paramFields.Add (pfidAsiento);
				
				ParameterField pfidTipo = new ParameterField ();
				ParameterDiscreteValue ValidTipo = new ParameterDiscreteValue ();
				pfidTipo.ParameterFieldName = "@idTipo";
				ValidTipo.Value = idTipo;
				pfidTipo.CurrentValues.Add (ValidTipo);
				paramFields.Add (pfidTipo);
				
				ParameterField pfidAsientoDet = new ParameterField ();
				ParameterDiscreteValue ValIdAsientoDet = new ParameterDiscreteValue ();
				pfidAsientoDet.ParameterFieldName = "@idAsientoDet";
				ValIdAsientoDet.Value = (int)this.txtIdAsiento.Value;
				pfidAsientoDet.CurrentValues.Add (ValIdAsientoDet);
				paramFields.Add (pfidAsientoDet);
											
				Reporte miRepor = new Reporte(strReporte, "", "ComprobanteDeEgreso", true);
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Egreso");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				miRepor.Show();	

				//				Reporte miRep = new Reporte("SolicitudVacaciones.rpt", "");
				//				miRep.MdiParent = this.MdiParent;
				//				miRep.Titulo("Solicitud Vacaciones");
				//				miRep.crVista.ParameterFieldInfo = paramFields;
				//				miRep.Show();




				Cursor = Cursors.Default;					
				#endregion Egresos				
			}
			else
			{
				Reporte miReporte = new Reporte(strReporte, strFiltro);

				if (iTipoImp == 0)
				{
					miReporte.MdiParent = MdiParent;
					miReporte.Show();
				}
				else miReporte.ImprimeRep(1);
			}

			Cursor = Cursors.Default;
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{			
			FuncionesProcedimientos.FormatoGrid(e, "Debe", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Haber", 2);
		}

		private void cmbTipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.txtNumero.Enabled) this.txtNumero.Focus(); else this.txtDescripcion.Focus();
		}

		private void txtDescripcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.txtChequeDeposito.Enabled) this.txtChequeDeposito.Focus(); else this.dtFecha.Focus();
		}

		private void txtChequeDeposito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFecha.Focus();
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDescripcion.Focus();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCentroDeCosto.Focus();
		}

		private void cmbCentroDeCosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				this.uGridDetalle.DisplayLayout.Bands[0].AddNew();
				this.uGridDetalle.ActiveCell = this.uGridDetalle.Rows[0].Cells["Codigo"];
			}
		}

		private void cmbTipo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbTipo.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un Tipo de Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTipo.Focus();
					e.Cancel = true;
				}
				else if (bNuevo && (int)this.cmbTipo.Value == 3 && MenuLatinium.IdUsuario != 86)
				{
					MessageBox.Show("No puede crear Egresos desde este Modulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTipo.Focus();
					e.Cancel = true;
				}
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtFecha.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione una Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;
				}
				else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoC, "Contabilidad"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;
				}
			}
		}

		private void txtDescripcion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			//			if (bEdicion)
			//			{
			//				if (this.txtDescripcion.Text.ToString().Length == 0)
			//				{
			//					MessageBox.Show("Ingrese la Descripción del Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtDescripcion.Focus();
			//					e.Cancel = true;
			//				}
			//			}
		}

		private void cmbCentroDeCosto_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbCentroDeCosto.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un Centro de Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbCentroDeCosto.Focus();
					e.Cancel = true;
				}
			}
		}

		private void cmbTipo_ValueChanged(object sender, System.EventArgs e)
		{
			if (cmbTipo.ActiveRow != null)
			{
				if ((int)cmbTipo.Value == 2) this.lblCheque.Text = "Deposito N.";
				else if ((int)cmbTipo.Value == 3) this.lblCheque.Text = "Cheque N.";
				else if ((int)cmbTipo.Value == 4) this.lblCheque.Text = "Debito N.";
				else if ((int)cmbTipo.Value == 5) this.lblCheque.Text = "Credito N.";
			}

			if (bEdicion)
			{
				if (cmbTipo.ActiveRow != null)
				{
					if ((int)cmbTipo.Value == 2) this.lblCheque.Text = "Deposito N.";
					else if ((int)cmbTipo.Value == 3) this.lblCheque.Text = "Cheque N.";
					else if ((int)cmbTipo.Value == 4) this.lblCheque.Text = "Debito N.";
					else if ((int)cmbTipo.Value == 5) this.lblCheque.Text = "Credito N.";
							
					if ((int)this.cmbTipo.Value > 1) this.txtChequeDeposito.Enabled = true;
					else this.txtChequeDeposito.Enabled = false;

					if (this.uGridDetalle.Rows.Count > 0)
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDetalle.Rows.All)
						{
							if (Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From CuentaCorriente Where idCuenta = {0}", (int)dr.Cells["idCuenta"].Value), true) > 0) 
							{
								if ((int)this.cmbTipo.Value == 1) dr.Cells["idFormaPago"].Value = System.DBNull.Value;
								if ((int)this.cmbTipo.Value == 2) dr.Cells["idFormaPago"].Value = 13;
								if ((int)this.cmbTipo.Value == 3) dr.Cells["idFormaPago"].Value = 2;
								if ((int)this.cmbTipo.Value == 4) dr.Cells["idFormaPago"].Value = 16;
								if ((int)this.cmbTipo.Value == 5) dr.Cells["idFormaPago"].Value = 4;
							}
						}
					}
				}
			}
		}

		private void txtChequeDeposito_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.cmbTipo.Value == 3)
				{
					if (this.txtChequeDeposito.Text.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el Numero de Cheque", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtChequeDeposito.Focus();
						e.Cancel = true;
					}					
				}
				else if ((int)this.cmbTipo.Value == 2)
				{
					if (this.txtChequeDeposito.Text.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el Numero del Deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtChequeDeposito.Focus();
						e.Cancel = true;
					}					
				}
			}
		}

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetAsiento"].Value = 0;
				e.Row.Cells["idAsiento"].Value = 0;
				e.Row.Cells["idProyecto"].Value = 19;
				e.Row.Cells["Debe"].Value = 0.00;
				e.Row.Cells["Haber"].Value = 0.00;
				e.Row.Cells["idAutorizacionPago"].Value = 0;
				e.Row.Cells["Origen"].Value = 0;
				e.Row.Cells["Tipo"].Value = 0;
				e.Row.Cells["idFactura"].Value = 0;				
			}
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{			
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalle);

			#region Cuentas
			if (e.KeyCode == Keys.F3)
			{	
				int idOrigen = 0;				
				if (this.cmbOrigen.ActiveRow != null) idOrigen = (int)this.cmbOrigen.Value;
				
				if (idOrigen == 0)
				{
					if ((this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO") && bEdicion)
					{																								
						using (BuscaCuentas Busqueda = new BuscaCuentas(this.uGridDetalle.ActiveRow.Cells["Codigo"].Value.ToString(), 1))
						{
							if (DialogResult.OK == Busqueda.ShowDialog())
							{
								this.uGridDetalle.ActiveRow.Cells["Codigo"].Value = Busqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
							}
						}
					}							
				}
			}			
			#endregion Cuentas
		}

		private void uGridDetalle_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if (!bEliminaAlValidar)
				{
					for (int i=0; i<e.Rows.Length; i++)
					{					
						if (this.cmbOrigen.ActiveRow != null)
						{
							MessageBox.Show("No Puede Eliminar Filas de Asientos Creados Automaticamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cancel = true;			
							return;
						}
						
						if (DialogResult.Yes ==	MessageBox.Show("¿Esta Seguro de Borrar las Lineas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{
							if ((int)e.Rows[i].Cells["idDetAsiento"].Value == 0) e.DisplayPromptMsg = false;
							else
							{
								string sSQL = string.Format("Delete From AsientoDetalle Where idDetAsiento = {0}", (int)e.Rows[i].Cells["idDetAsiento"].Value);
								Funcion.EjecutaSQL(cdsSeteoC, sSQL, true);
							}
						}
						else 	e.Cancel = true;						
					}
				}
				else e.DisplayPromptMsg = false;
			}
		}

		private void TotalDebeHaber()
		{
			decimal dDebe = 0;
			decimal dHaber = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDetalle.Rows.All)
			{
				if (dr.Cells["Debe"].Value != System.DBNull.Value) dDebe += (decimal) dr.Cells["Debe"].Value;
				if (dr.Cells["Haber"].Value != System.DBNull.Value) dHaber += (decimal) dr.Cells["Haber"].Value;
			}
			
			txtDebe.Value = Math.Round(dDebe, 2);
			txtHaber.Value = Math.Round(dHaber, 2);

			if (Math.Round(dDebe, 2) != Math.Round(dHaber, 2)) 
			{
				this.lblDescuadrado.Text = "DESCUADRADO";
				decimal dDiferencia = dDebe - dHaber;
				this.lblDescuadrado.Text += " " + dDiferencia.ToString("#,##0.00");
				return;
			}
			else this.lblDescuadrado.Text = "";
		}

		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
	
			#region Codigo
			if (e.Cell.Column.ToString() == "Codigo")
			{
				#region Busca Cuenta
				string sSQLContador = string.Format("Exec AsientoBuscaCuenta '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
				int iContador = Funcion.iEscalarSQL(cdsSeteoC, sSQLContador);

				if (iContador == 1)
				{
					string sSQLInfo = string.Format("Exec AsientoInfoCuenta '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoC, sSQLInfo, true);
					dr.Read();
					if (dr.HasRows)
					{
						e.Cell.Row.Cells["idCuenta"].Value = dr.GetInt32(0);
						e.Cell.Row.Cells["Descripcion"].Value = dr.GetString(2);
					}
					dr.Close();

					if (Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From CuentaCorriente Where idCuenta = {0}", (int)e.Cell.Row.Cells["idCuenta"].Value), true) > 0) 
					{
						if ((int)this.cmbTipo.Value == 2) e.Cell.Row.Cells["idFormaPago"].Value = 13;
						if ((int)this.cmbTipo.Value == 4) e.Cell.Row.Cells["idFormaPago"].Value = 16;
						if ((int)this.cmbTipo.Value == 5) e.Cell.Row.Cells["idFormaPago"].Value = 4;
					}
				}
				else
				{
					using (BuscaCuentas Busqueda = new BuscaCuentas(e.Cell.Row.Cells["Codigo"].Value.ToString(), 1))
					{
						if (DialogResult.OK == Busqueda.ShowDialog())
						{
							e.Cell.Row.Cells["Codigo"].Value = Busqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
						}
					}
				}
				#endregion Busca Cuenta	
			}
			#endregion Codigo
	
			
			#region Forma de Pago
			if (e.Cell.Column.ToString() == "idFormaPago")
			{
				if (e.Cell.Row.Cells["idFormaPago"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idFormaPago"].Value > 0)
					{
						#region Valida Forma de Pago con Cuenta Bancaria
						int iCuentaBancaria = Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From CuentaCorriente Where idCuenta = {0}", (int)e.Cell.Row.Cells["idCuenta"].Value), true);
						if (iCuentaBancaria == 0) 
						{
							e.Cell.Value = System.DBNull.Value;
							MessageBox.Show("La forma de pago es para cuentas bancarias", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
						#endregion Valida Forma de Pago con Cuenta Bancaria

						#region Valida Forma de Pago Seleccionada
						if (iCuentaBancaria > 0)
						{
							if ((int)this.cmbTipo.Value == 2 && ((int)e.Cell.Row.Cells["idFormaPago"].Value != 13 && (int)e.Cell.Row.Cells["idFormaPago"].Value != 7)) 
							{
								e.Cell.Row.Cells["idFormaPago"].Value = 13;
								MessageBox.Show("El Tipo de Movimiento Contable es Ingreso, la Forma de Pago debe ser DEPOSITO o TRANSFERENCIA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							}
							if ((int)this.cmbTipo.Value == 4 && (int)e.Cell.Row.Cells["idFormaPago"].Value != 16)
							{
								e.Cell.Row.Cells["idFormaPago"].Value = 16;
								MessageBox.Show("El Tipo de Movimiento Contable es Nota de Debito, la Forma de Pago debe ser NOTA DE DEBITO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							}
							if ((int)this.cmbTipo.Value == 5 && (int)e.Cell.Row.Cells["idFormaPago"].Value != 4)
							{
								e.Cell.Row.Cells["idFormaPago"].Value = 4;
								MessageBox.Show("El Tipo de Movimiento Contable es Nota de Credito, la Forma de Pago debe ser NOTA DE CREDITO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							}
						}
						#endregion Valida Forma de Pago Seleccionada
					}
				}
			}
			#endregion Forma de Pago

			#region Debe
			if (e.Cell.Column.ToString() == "Debe")
			{
				if (e.Cell.Row.Cells["idCuenta"].Value == System.DBNull.Value)
				{	
					if (e.Cell.Row.Cells["Debe"].Value != System.DBNull.Value)
					{	
						if ((decimal)e.Cell.Row.Cells["Debe"].Value > 0.00m)
						{
							e.Cell.Value = 0.00m;
							MessageBox.Show("Seleccione una Cuenta antes de ingresar el valor del Debe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}

				TotalDebeHaber();
			}
			#endregion Debe

			#region Haber
			if (e.Cell.Column.ToString() == "Haber")
			{
				if (e.Cell.Row.Cells["idCuenta"].Value == System.DBNull.Value)
				{							
					if (e.Cell.Row.Cells["Haber"].Value != System.DBNull.Value)
					{
						if ((decimal)e.Cell.Row.Cells["Haber"].Value > 0.00m)
						{
							e.Cell.Value = 0.00m;
							MessageBox.Show("Seleccione una Cuenta antes de ingresar el valor del Debe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}

				TotalDebeHaber();
			}
			#endregion Haber

			
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void uGridDetalle_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.uGridDetalle.ActiveRow.Cells["idCuenta"].Value == System.DBNull.Value && 
				(this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "DEBE" || this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "HABER" )) e.Handled = true;

			int iCuentaBancaria = 0;
			if (this.uGridDetalle.ActiveRow.Cells["idCuenta"].Value != System.DBNull.Value)
				iCuentaBancaria = Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From CuentaCorriente Where idCuenta = {0}", (int)this.uGridDetalle.ActiveRow.Cells["idCuenta"].Value), true);

			if (iCuentaBancaria > 0 && (int)this.cmbTipo.Value == 5 && this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "HABER") e.Handled = true;
			if (iCuentaBancaria > 0 && (int)this.cmbTipo.Value == 4 && this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "DEBE" ) e.Handled = true;

			int iDebitoCredito = 0;

			if (this.uGridDetalle.ActiveRow.Cells["idCuenta"].Value != System.DBNull.Value)
				if (iCuentaBancaria == 0) 
					iDebitoCredito = Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select ISNULL(DebitoCredito, 0) From Cuenta Where idCuenta = {0}", 
						(int)this.uGridDetalle.ActiveRow.Cells["idCuenta"].Value));
			
			if (iDebitoCredito > 0 && (int)this.cmbTipo.Value < 4 && this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "HABER") e.Handled = true;
			if (iDebitoCredito > 0 && (int)this.cmbTipo.Value < 4 && this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "DEBE") e.Handled = true;

			if (iCuentaBancaria == 0 && iDebitoCredito == 1 && (int)this.cmbTipo.Value == 5 && this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "HABER") e.Handled = true;
			if (iCuentaBancaria == 0 && iDebitoCredito == 2 && (int)this.cmbTipo.Value == 4 && this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "DEBE" ) e.Handled = true;

			if (this.uGridDetalle.ActiveRow.Cells["Debe"].Value != System.DBNull.Value)
				if ((decimal)this.uGridDetalle.ActiveRow.Cells["Debe"].Value > 0 && this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "HABER") e.Handled = true;

			if (this.uGridDetalle.ActiveRow.Cells["Haber"].Value != System.DBNull.Value)
				if ((decimal)this.uGridDetalle.ActiveRow.Cells["Haber"].Value > 0 && this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "DEBE") e.Handled = true;
		}

		private void uGridDetalle_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			TotalDebeHaber();
		}

		private void cmbTipo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Elimina filas vacias
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
			{
				if (dr.Cells["idCuenta"].Value == System.DBNull.Value || dr.Cells["Codigo"].Value == System.DBNull.Value || dr.Cells["Descripcion"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}
			}
			#endregion Elimina filas vacias

			bool bNumAutomatico = Funcion.bEscalarSQL(cdsSeteoC, string.Format("Select Automatico From AsientoTipo Where idTipoAsiento = {0}", (int)this.cmbTipo.Value), true);

			TotalDebeHaber();

			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoC, "Contabilidad"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede Grabar este Asiento, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);											
			}
			else if (this.cmbTipo.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Tipo de Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipo.Focus();			
			}
			else if (this.txtDescripcion.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese una descripcion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDescripcion.Focus();			
			}
			else if (bNuevo && (int)this.cmbTipo.Value == 3)
			{
				MessageBox.Show("No puede crear Egresos desde este Modulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipo.Focus();				
			}
			else if (!bNumAutomatico && Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From Asiento Where CodAsiento = '{0}' And idTipoAsiento = {1} And idAsiento <> {2} And LEN(CodAsiento) > 0", this.txtNumero.Text.ToString(), (int)this.cmbTipo.Value, (int)this.txtIdAsiento.Value)) > 0)
			{
				MessageBox.Show(string.Format("Numero de '{0}' ya Existe", this.cmbTipo.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();				
			}
			else if (this.txtDescripcion.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese la Descripción del Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDescripcion.Focus();			
			}
				//			else if ((int)this.cmbTipo.Value > 1 && this.txtChequeDeposito.Text.ToString().Length == 0)
				//			{
				//				MessageBox.Show("Ingrese el Numero de Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//				this.txtChequeDeposito.Focus();				
				//			}								
			else if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();			
			}
			else if (this.cmbCentroDeCosto.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Centro de Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCentroDeCosto.Focus();			
			}
			else if (this.uGridDetalle.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese el Detalle del Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridDetalle.Focus();			
			}
			else if ((decimal)this.txtDebe.Value == 0 && (decimal)this.txtHaber.Value == 0)
			{
				MessageBox.Show("No Puede Grabar en Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridDetalle.Focus();			
			}
			else if (Math.Round((decimal)this.txtDebe.Value, 2) != Math.Round((decimal)this.txtHaber.Value, 2))
			{
				decimal dDiferencia = Math.Round((decimal)this.txtDebe.Value, 2) - Math.Round((decimal)this.txtHaber.Value, 2);
				MessageBox.Show("ASIENTO DESCUADRADO DIFERENCIA DE " + dDiferencia.ToString("#,##0.00"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				this.uGridDetalle.Focus();			
			}
			else
			{
				#region Valida Detalle y Guarda Registro
				#region Cuenta Contable
				bool bCuenta = false;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDetalle.Rows.All)
				{
					if (dr.Cells["idCuenta"].Value == System.DBNull.Value && ((decimal)dr.Cells["Debe"].Value > 0 || (decimal)dr.Cells["Haber"].Value > 0))
					{
						bCuenta = true;
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						continue;
					}
				}
				#endregion Cuenta Contable

				if (bCuenta)
				{
					MessageBox.Show("Ingrese una Cuenta o Elimine la Linea", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
				else
				{
					#region Asiento Banco
					bool bAsientoBanco = false;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					{
						if (Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From CuentaCorriente Where idCuenta = {0}", (int)dr.Cells["idCuenta"].Value), true) > 0 && (int)this.cmbTipo.Value == 1) 
						{
							bAsientoBanco = true;
							this.uGridDetalle.ActiveRow = dr;
							this.uGridDetalle.ActiveRow.Selected = true;
							continue;
						}
					}
					#endregion Asiento Banco

					if (bAsientoBanco)
					{
						MessageBox.Show("No puede crear un ASIENTO con Cuenta Bancaria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					else
					{
						#region Forma de Pago
						bool bFormaDePago = false;
						
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
						{
							if (Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From CuentaCorriente Where idCuenta = {0}", (int)dr.Cells["idCuenta"].Value), true) > 0)
							{
								if (dr.Cells["idFormaPago"].Value == System.DBNull.Value) 
								{
									bFormaDePago = true;
									this.uGridDetalle.ActiveRow = dr;
									this.uGridDetalle.ActiveRow.Selected = true;
									continue;
								}
								else if ((int)dr.Cells["idFormaPago"].Value == 0) 
								{
									bFormaDePago = true;
									this.uGridDetalle.ActiveRow = dr;
									this.uGridDetalle.ActiveRow.Selected = true;
									continue;
								}								
							}
						}
						#endregion Forma de Pago

						if (bFormaDePago)
						{
							MessageBox.Show("Seleccione una forma de pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}
						else
						{
							#region Numero Documento
							bool bNumeroDocumento = false;

							if ((int)this.cmbTipo.Value > 1)
							{						
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
								{
									if (Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From CuentaCorriente Where idCuenta = {0}", (int)dr.Cells["idCuenta"].Value), true) > 0)
									{
										if (Funcion.bEscalarSQL(cdsSeteoC, string.Format("Exec ValidaNumeroDocumento {0}, '{1}', {2}, {3}", (int)dr.Cells["idCuenta"].Value, this.txtChequeDeposito.Text.ToString(), (int)this.cmbTipo.Value, (int)this.txtIdAsiento.Value), true))										
										{
											bNumeroDocumento = true;
											this.uGridDetalle.ActiveRow = dr;
											this.uGridDetalle.ActiveRow.Selected = true;
											continue;
										}								
									}
								}								
							}
							#endregion Numero Documento

							if (bNumeroDocumento)
							{
								MessageBox.Show("el Numero de Documento ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.txtNumero.Focus();
							}
							else
							{
								#region Valores
								bool bValores = false;

								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDetalle.Rows.All)
								{
									if ((decimal)dr.Cells["Debe"].Value == 0 && (decimal)dr.Cells["Haber"].Value == 0)
									{
										bValores = true;
										this.uGridDetalle.ActiveRow = dr;
										this.uGridDetalle.ActiveRow.Selected = true;
										continue;
									}
								}
								#endregion Valores

								if (bValores)
								{
									MessageBox.Show("Ingrese Valor al Debe o al Haber", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								}
								else
								{
									#region Centro de Costo
									bool bCentroDeCosto = false;

									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDetalle.Rows.All)
									{
										if (dr.Cells["idProyecto"].Value == System.DBNull.Value)
										{
											bCentroDeCosto = true;
											this.uGridDetalle.ActiveRow = dr;
											this.uGridDetalle.ActiveRow.Selected = true;
											continue;
										}
										else if ((int)dr.Cells["idProyecto"].Value == 0)
										{
											bCentroDeCosto = true;
											this.uGridDetalle.ActiveRow = dr;
											this.uGridDetalle.ActiveRow.Selected = true;
											continue;
										}
									}
									#endregion Centro de Costo

									if (bCentroDeCosto)
									{
										MessageBox.Show("Seleccione el Centro de Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									}
									else
									{
										#region Cuenta Bancaria Notas De Debito
										//										int iGastosBancarios = 0;
										//										if ((int)this.cmbTipo.Value == 4)
										//										{											
										//											foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDetalle.Rows.All)
										//											{
										//												if (Funcion.bEscalarSQL(cdsSeteoC, string.Format("Select Gasto_Bancario From Cuenta Where idCuenta = {0}", dr.Cells["idCuenta"].Value), true))
										//												{
										//													iGastosBancarios++;
										//													this.uGridDetalle.ActiveRow = dr;
										//													this.uGridDetalle.ActiveRow.Selected = true;
										//													continue;
										//												}
										//											}
										//
										//											if (iGastosBancarios == 0)
										//											{
										//												MessageBox.Show("Para Notas De Debito solo se permite registro de gastos bancarios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
										//												return;
										//											}										
										//										}
										#endregion Cuenta Bancaria Notas De Debito

										#region valida Presupuesto
										if (bNuevo)
										{
											if (MenuLatinium.stIdDB == 1)// solo para point 
											{
												int iCentroDeCosto = 0 ;
												foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
												{
													if (dr.Cells["idProyecto"].Value != System.DBNull.Value) iCentroDeCosto = (int)dr.Cells["idProyecto"].Value;
													if ( iCentroDeCosto > 0 )
													{
														int idCuentaExi = 0;
														idCuentaExi = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCuentaPresupuestoAsiento] {0}, {1}, '{2}'",
															(int)dr.Cells["idCuenta"].Value, iCentroDeCosto, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));	

														if (idCuentaExi > 0 )
														{
															int iProyectoS = 0;
															iProyectoS = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCCPresupuestoAsiento] {0}, '{1}'",
																iCentroDeCosto,Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd") ));

															if ( iProyectoS > 0 )
															{
																#region valida valores
																decimal Presupuesto = 0;
																decimal Utilizado = 0;
																decimal Disponible = 0;
																string Cuenta = "";
																if ( (decimal)dr.Cells["Debe"].Value  > 0 )
																{
																	string sSQL = string.Format("Exec [ValidaCCPresupuestoAsientoValores] '{0}', {1}, {2}", 
																		Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)dr.Cells["idCuenta"].Value, (int) iCentroDeCosto );
																	SqlDataReader dra = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
																	dra.Read();
																	if (dra.HasRows)
																	{
																		if (dra.GetValue(0) != System.DBNull.Value)Presupuesto = dra.GetDecimal (0);
																		if (dra.GetValue(1) != System.DBNull.Value)Utilizado = dra.GetDecimal (1);
																		if (dra.GetValue(2) != System.DBNull.Value)Disponible = dra.GetDecimal (2);
																		if (dra.GetValue(3) != System.DBNull.Value)Cuenta = dra.GetString (3);
											
																	}
																	dra.Close();
																}
																if ( (decimal)dr.Cells["Debe"].Value > Disponible )
																{
																	this.uGridDetalle.ActiveRow = dr;
																	this.uGridDetalle.ActiveRow.Selected = true;
																	MessageBox.Show(string.Format("Ha superado el presupuesto de '{0}' \n\n Presupuesto : {1}  \n\n Utilizado : {2} \n\n Solicitado : {3}", Cuenta, Presupuesto, Utilizado, Convert.ToDecimal (dr.Cells["Debe"].Value )), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);											
																	return;
																}
																#endregion valida valores
															}           
														}
													}
												}
											}
										}
										#endregion valida Presupuesto

										#region Conexion
										using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
										{
											oConexion.Open();

											SqlTransaction oTransaction = null;
											SqlCommand oCmdActualiza = oConexion.CreateCommand();
											oCmdActualiza.Connection = oConexion;
											oCmdActualiza.CommandType = CommandType.Text;
											oCmdActualiza.CommandTimeout = 0;
						
											#region Guarda Registro
											try
											{		
												TotalDebeHaber();

												oTransaction = oConexion.BeginTransaction();
												oCmdActualiza.Transaction = oTransaction;

												#region Numeracion
												if (bNuevo)
												{
													string sSQLNumero = string.Format("Exec AsientoNumeracion {0}", (int)this.cmbTipo.Value);
													oCmdActualiza.CommandText = sSQLNumero;
													this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
												}
												#endregion Numeracion

												#region Graba Encabezado
												DateTime dtFecha = (DateTime)this.dtFecha.Value;
												int idOrigenAsiento = 0;
												if (this.cmbOrigen.Value != System.DBNull.Value) idOrigenAsiento = (int)this.cmbOrigen.Value;

												string sSQLMaestro = string.Format("Exec AsientoGrabaMaestro {0}, {1}, {2}, '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', {9}, '{10}', {11}, {12}",
													(int)this.txtIdAsiento.Value, (int)this.cmbTipo.Value, idOrigenAsiento, this.txtNumero.Text.ToString(), this.txtDescripcion.Text.ToString(), this.txtRUC.Text.ToString(), this.txtChequeDeposito.Text.ToString(), this.txtDctoOrigen.Text.ToString(), dtFecha.ToString("yyyyMMdd"), (int)this.cmbCentroDeCosto.Value, this.txtNotas.Text.ToString(), (decimal)this.txtDebe.Value, (decimal)this.txtHaber.Value);
												oCmdActualiza.CommandText = sSQLMaestro;									
												this.txtIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar();									
												#endregion Graba Encabezado

												#region Graba Detalle
												foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
												{	
													int iCentroDeCosto = 0;
													if (dr.Cells["idProyecto"].Value != System.DBNull.Value) iCentroDeCosto = (int)dr.Cells["idProyecto"].Value;
													int idFormaPago = 0;
													if (dr.Cells["idFormaPago"].Value != System.DBNull.Value) idFormaPago = (int)dr.Cells["idFormaPago"].Value;
										
													string sSQLDetalle = string.Format("Exec AsientoGrabaDetalle {0}, {1}, {2}, {3}, '{4}', {5}, {6}, {7}",
														(int)dr.Cells["idDetAsiento"].Value, (int)this.txtIdAsiento.Value, (int)dr.Cells["idCuenta"].Value, iCentroDeCosto, dr.Cells["Detalle"].Value.ToString(), 
														(decimal)dr.Cells["Debe"].Value, (decimal)dr.Cells["Haber"].Value, idFormaPago);
													oCmdActualiza.CommandText = sSQLDetalle;
													oCmdActualiza.ExecuteNonQuery();
												}
												#endregion Graba Detalle				

												#region Actualiza Grid
												string sSQLActDetalle = string.Format("Exec AsientoConsultaDetalle {0}", (int)this.txtIdAsiento.Value);
												oCmdActualiza.CommandText = sSQLActDetalle;
												SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
												DataTable oDTDetalle = new DataTable();
												oDADetalle.Fill(oDTDetalle);
												this.uGridDetalle.DataSource = oDTDetalle;
												#endregion Actualiza Grid

												oTransaction.Commit();

												#region Controles
												this.cmbTipo.Enabled = false;			
												this.txtNumero.Enabled = false;
												this.txtDescripcion.Enabled = false;
												this.txtRUC.Enabled = false;
												this.txtChequeDeposito.Enabled = false;
												this.txtDctoOrigen.Enabled = false;
												this.dtFecha.Enabled = false;
												this.cmbCentroDeCosto.Enabled = false;			
												this.txtNotas.Enabled = false;

												FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);

												if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
												if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
												if (miAcceso.Editar) this.btnEditar.Enabled = true;
												if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
												this.btnGuardar.Enabled = false;
												if (!miAcceso.Anular) this.btnAnular.Enabled = false;																							
									
												bNuevo = false;
												bEdicion = false;

												this.CargaRegistro((int)this.txtIdAsiento.Value);

												#endregion Controles
											}
											catch (Exception ex)
											{
												oTransaction.Rollback();

												MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);												
											}
											finally
											{
												oConexion.Close();
											}
											#endregion Guarda Registro
										}
										#endregion Conexion
									}
								}
							}
						}
					}
				}
				#endregion Valida Detalle y Guarda Registro
			}
		}

		private void txtDescripcion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From Asiento Where Numero = '{0}' And idTipoAsiento = {1} And idAsiento <> {2} And LEN(Numero) > 0", this.txtNumero.Text.ToString(), (int)this.cmbTipo.Value, (int)this.txtIdAsiento.Value)) > 0)
				{
					MessageBox.Show(string.Format("Numero de '{0}' ya Existe", this.cmbTipo.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					e.Cancel = true;
				}
			}
		}

		private void txtChequeDeposito_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbCentroDeCosto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{			
			if (bAnulado && !miAcceso.Eliminar)
			{
				MessageBox.Show("El Registro ya esta ANULADO, No tiene acceso a Eliminar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
				return;
			}
			if ((int)this.cmbTipo.Value != 3) 
			{
				if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoC, "Contabilidad"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No Puede Anular/Borrar este Asiento, El Período '{0}' esta Cerrado, \n\nConsulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
					return;
				}	
			}

			#region Genera Nuevo Egreso Anulado
			if (bNuevo)
			{
				if ((int)	this.cmbTipo.Value != 3)
				{
					MessageBox.Show("Esta opcion es solo para cheques", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
					this.cmbTipo.Focus();
					return;
				}
				if (this.txtChequeDeposito.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el numero de Cheque", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
					this.txtChequeDeposito.Focus();
					return;
				}
				if (this.uGridDetalle.Rows.Count == 0)
				{
					MessageBox.Show("Ingrese la cuenta del Banco al que pertenece el Cheque", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
					this.txtChequeDeposito.Focus();
					return;
				}
				if ((int)	this.cmbTipo.Value == 3 && Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From Asiento Where CodAsiento = '{0}' And idTipoAsiento = {1} And idAsiento <> {2} And LEN(CodAsiento) > 0", this.txtNumero.Text.ToString(), (int)this.cmbTipo.Value, (int)this.txtIdAsiento.Value)) > 0)
				{
					MessageBox.Show(string.Format("Numero de '{0}' ya Existe", this.cmbTipo.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();				
					return;
				}
				
				int iCont = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				{
					if (Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From CuentaCorriente Where idCuenta = {0}", (int)dr.Cells["idCuenta"].Value), true) > 0) iCont++;	
				}

				if (iCont == 0)
				{
					MessageBox.Show("Ingrese la cuenta del Banco al que pertenece el Cheque", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
					this.txtChequeDeposito.Focus();
					return;
				}
				
				#region Crea Registro Anulado
				/*
				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 0;
						
					try
					{							
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;

						if (bNuevo)
						{
							string sSQLNumero = string.Format("Exec AsientoNumeracion {0}", (int)this.cmbTipo.Value);
							oCmdActualiza.CommandText = sSQLNumero;
							this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
						}
									
						DateTime dtFecha = (DateTime)this.dtFecha.Value;
						int idOrigenAsiento = 0;
						if (this.cmbOrigen.Value != System.DBNull.Value) idOrigenAsiento = (int)this.cmbOrigen.Value;

						string sSQLMaestro = string.Format("Exec AsientoGrabaMaestro {0}, {1}, {2}, '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', {9}, '{10}', {11}, {12}",
							(int)this.txtIdAsiento.Value, (int)this.cmbTipo.Value, idOrigenAsiento, this.txtNumero.Text.ToString(), this.txtDescripcion.Text.ToString(), this.txtRUC.Text.ToString(), this.txtChequeDeposito.Text.ToString(), this.txtDctoOrigen.Text.ToString(), dtFecha.ToString("yyyyMMdd"), (int)this.cmbCentroDeCosto.Value, this.txtNotas.Text.ToString(), (decimal)this.txtDebe.Value, (decimal)this.txtHaber.Value);
						oCmdActualiza.CommandText = sSQLMaestro;									
						this.txtIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar();									

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
						{	
							int iCentroDeCosto = 0;
							if (dr.Cells["idProyecto"].Value != System.DBNull.Value) iCentroDeCosto = (int)dr.Cells["idProyecto"].Value;
										
							string sSQLDetalle = string.Format("Exec AsientoGrabaDetalle {0}, {1}, {2}, {3}, '{4}', {5}, {6}, {7}",
								(int)dr.Cells["idDetAsiento"].Value, (int)this.txtIdAsiento.Value, (int)dr.Cells["idCuenta"].Value, iCentroDeCosto, dr.Cells["Detalle"].Value.ToString(), 
								(decimal)dr.Cells["Debe"].Value, (decimal)dr.Cells["Haber"].Value, (int)dr.Cells["idFormaPago"].Value);
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}

						string sSQLAnula = string.Format("Update Asiento Set Borrar = 1 Where idAsiento = {0}", (int)this.txtIdAsiento.Value);
						oCmdActualiza.CommandText = sSQLAnula;
						oCmdActualiza.ExecuteNonQuery();								

						string sSQLActDetalle = string.Format("Exec AsientoConsultaDetalle {0}", (int)this.txtIdAsiento.Value);
						oCmdActualiza.CommandText = sSQLActDetalle;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);
						this.uGridDetalle.DataSource = oDTDetalle;
	
						oTransaction.Commit();

						this.cmbTipo.Enabled = false;			
						this.txtNumero.Enabled = false;
						this.txtDescripcion.Enabled = false;
						this.txtRUC.Enabled = false;
						this.txtChequeDeposito.Enabled = false;
						this.txtDctoOrigen.Enabled = false;
						this.dtFecha.Enabled = false;
						this.cmbCentroDeCosto.Enabled = false;			
						this.txtNotas.Enabled = false;

						FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);

						if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
						if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
						if (miAcceso.Editar) this.btnEditar.Enabled = true;
						if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
						if (!miAcceso.Eliminar) this.btnAnular.Enabled = false;
								
						this.btnAnular.Text = "Eliminar";
								
						this.btnGuardar.Enabled = false;
						this.lblEstado.Text = "ANULADO";
						bAnulado = true;
									
						bNuevo = false;
						bEdicion = false;
					}
					catch (Exception ex)
					{
						try
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
						}
					}
					finally
					{
						oConexion.Close();
					}
				}
				*/
				#endregion Crea Registro Anulado								

				return;
			}
			#endregion Genera Nuevo Egreso Anulado

			if (!bAnulado)
			{				
				this.CargaRegistro((int)this.txtIdAsiento.Value);

				if (!FuncionesContabilidad.bValidaEdicionAnulacionAsientos(this.cmbOrigen, 3, cdsSeteoC)) return;
				
				if (this.ValidaMovimientoConciliado("ANULAR")) return;

				if ((int)this.cmbOrigen.Value == 23 && Funcion.iEscalarSQL(cdsSeteoC, string.Format("Select COUNT(*) From Asiento Where idOrigenAsiento = 41 And Borrar = 0 And CodAsiento = '{0}'", this.txtDctoOrigen.Text.ToString())) > 0)
				{
					MessageBox.Show("No puede Anular:\n\nEl Documento de Origen ya fue Cruzado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);										
					return;
				}
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro de Anular este Asiento?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)) 
				{
					using (frmJustAnulaVenta Nota = new frmJustAnulaVenta(20))
					{
						if (DialogResult.OK == Nota.ShowDialog())
						{
							if ((int)this.cmbTipo.Value == 3) 
							{
								bool bAsientoEgresoAnulado = Funcion.bEscalarSQL(cdsSeteoC, string.Format("select ISNULL (AsientoEgresoAnulado, 0)  from Seteo"), true);

								if(bAsientoEgresoAnulado)
								{
									#region Encabezado Anula
									string sNumEgresoA ="";
									int idAsientoAnulado =0;
									string sSQLNumero = string.Format("Exec AsientoNumeracion {0}", 1);
									sNumEgresoA = Funcion.sEscalarSQL(cdsSeteoF, sSQLNumero);					

									string sSQLEncEgreso = string.Format("Exec GeneraEncabezadoEgresoAnulado {0}, '{1}', '{2}', {3}, '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', {10}, {11}, {12}", 
										0, sNumEgresoA, this.txtChequeDeposito.Text.ToString(), 19, this.txtDescripcion.Text.ToString(), Nota.txtNotas.Text.ToString(),this.txtDctoOrigen.Text.ToString (), 
										DateTime.Today.ToString("yyyyMMdd"), 
										DateTime.Today.ToString("yyyyMMdd"),  this.txtRUC.Text.ToString (), (decimal)this.txtHaber.Value, (decimal)this.txtDebe.Value,  86);								
									idAsientoAnulado = Funcion.iEscalarSQL(cdsSeteoF,sSQLEncEgreso);

									string sDetalle = "ANULACIÓN EGRESO # " + this.txtNumero.Text + "  CON CHEQUE # " + this.txtChequeDeposito.Text;
									#endregion Encabezado Anula

									#region HistorialAutorizacionDePago
									string sSQLDetalleHis = string.Format("Exec GeneraEgresoHistorialAutorizacionDePago {0}, {1}",
										(int)this.txtIdAsiento.Value, idAsientoAnulado);
									Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalleHis);

									#endregion HistorialAutorizacionDePago

									#region Graba Detalle
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
									{	
										int iCentroDeCosto = 0;
										if (dr.Cells["idProyecto"].Value != System.DBNull.Value) iCentroDeCosto = (int)dr.Cells["idProyecto"].Value;
										int idFormaPago = 0;
										if (dr.Cells["idFormaPago"].Value != System.DBNull.Value) idFormaPago = (int)dr.Cells["idFormaPago"].Value;
										
										string sSQLDetalle = string.Format("Exec AsientoGrabaDetalle {0}, {1}, {2}, {3}, '{4}', {5}, {6}, {7}",
											0, idAsientoAnulado, (int)dr.Cells["idCuenta"].Value, iCentroDeCosto, sDetalle, 
											(decimal)dr.Cells["Haber"].Value, (decimal)dr.Cells["Debe"].Value, idFormaPago);
										Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
									}
									#endregion Graba Detalle		
                
									string strReporte = "";			
									string strFiltro = "{Asiento.idAsiento} = " + idAsientoAnulado.ToString();
								
									if (MenuLatinium.stDirFacturacion == "CROMATIKA")	strReporte = "ComprobanteDeEgresoCromatika.rpt";
									else if (MenuLatinium.stDirFacturacion == "ZOOMMATIKA")	strReporte = "ComprobanteDeEgresoCromatika.rpt";
									else if (MenuLatinium.stDirFacturacion == "COMCREA")	strReporte = "ComprobanteDeEgresoCromatika.rpt";
									else if (MenuLatinium.stDirFacturacion == "LOTESERVICIO") strReporte = "ComprobanteDeEgreso.rpt";
									else if (MenuLatinium.stDirFacturacion == "GRETASANTORINI") strReporte = "ComprobanteDeEgreso.rpt";
									else if (MenuLatinium.stDirFacturacion == "PROSERVICIOS") strReporte = "ComprobanteDeEgreso.rpt";
									else if (MenuLatinium.stDirFacturacion == "IWANTYOU") strReporte = "ComprobanteDeEgreso.rpt";					
									else strReporte = "ComprobanteDeEgresoA4.rpt";
								
									#region Egresos Impresion
									int idTipo = Funcion.iEscalarSQL(cdsSeteoC, string.Format("Exec TipoAutorizacionAsiento {0}", idAsientoAnulado));

									ParameterFields paramFields = new ParameterFields ();
				
									ParameterField pfidAsiento = new ParameterField ();
									ParameterDiscreteValue ValIdAsiento = new ParameterDiscreteValue ();
									pfidAsiento.ParameterFieldName = "@idAsiento";
									ValIdAsiento.Value =  idAsientoAnulado;
									pfidAsiento.CurrentValues.Add (ValIdAsiento);
									paramFields.Add (pfidAsiento);
				
									ParameterField pfidTipo = new ParameterField ();
									ParameterDiscreteValue ValidTipo = new ParameterDiscreteValue ();
									pfidTipo.ParameterFieldName = "@idTipo";
									ValidTipo.Value = idTipo;
									pfidTipo.CurrentValues.Add (ValidTipo);
									paramFields.Add (pfidTipo);
				
									ParameterField pfidAsientoDet = new ParameterField ();
									ParameterDiscreteValue ValIdAsientoDet = new ParameterDiscreteValue ();
									pfidAsientoDet.ParameterFieldName = "@idAsientoDet";
									ValIdAsientoDet.Value = idAsientoAnulado;
									pfidAsientoDet.CurrentValues.Add (ValIdAsientoDet);
									paramFields.Add (pfidAsientoDet);
											
									Reporte miRepor = new Reporte(strReporte, "", "ComprobanteDeEgreso", true);
									miRepor.MdiParent = this.MdiParent;
									miRepor.Titulo("Egreso");
									miRepor.crVista.ParameterFieldInfo = paramFields;
									miRepor.Show();	

									Cursor = Cursors.Default;					
									#endregion Egresos Impresion				
								}
								else
								{
									string sSQLAnula = string.Format("Update Asiento Set Borrar = 1, Estado = 3, Notas = ISNULL(Notas, '') + '//' + '{1}' Where idAsiento = {0}", 
										(int)this.txtIdAsiento.Value, this.txtNotas.Text.ToString());
									Funcion.EjecutaSQL(cdsSeteoC, sSQLAnula);
								}
							}	
							else
							{
								string sSQLAnula = string.Format("Update Asiento Set Borrar = 1, Estado = 3, Notas = ISNULL(Notas, '') + '//' + '{1}' Where idAsiento = {0}", 
									(int)this.txtIdAsiento.Value, this.txtNotas.Text.ToString());
								Funcion.EjecutaSQL(cdsSeteoC, sSQLAnula);
							}

							if (this.cmbOrigen.Value != System.DBNull.Value)
							{
								if ((int)this.cmbOrigen.Value == 18)
								{
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
									{
										if (dr.Cells["idFactura"].Value != System.DBNull.Value)
										{
											if ((int)dr.Cells["idFactura"].Value > 0)
											{
												Funcion.EjecutaSQL(cdsSeteoC, string.Format("Exec ActualizaSaldoFacturas {0}, 1, 9", (int)dr.Cells["idFactura"].Value));
												Funcion.EjecutaSQL(cdsSeteoC, string.Format("Update Compra Set Bloquear = 0, idAsientoLoteSolidario = 0 Where idCompra = {0}", (int)dr.Cells["idFactura"].Value));
											}
										}
									}
								}

								if ((int)this.cmbOrigen.Value == 41)
								{
									Funcion.EjecutaSQL(cdsSeteoC, string.Format("Exec AnulaPagosConAnticipos {0}", (int)this.txtIdAsiento.Value));										
								}

								if ((int)this.cmbOrigen.Value == 54)
								{
									Funcion.EjecutaSQL(cdsSeteoC, string.Format("Update RolPrestamo Set Estado = 0, idAsiento = 0 Where idAsiento = {0}", (int)this.txtIdAsiento.Value));										
								}
							}

							if ((int)this.cmbTipo.Value == 3) Funcion.EjecutaSQL(cdsSeteoC, string.Format("Exec AnulaAutorizacionPago {0}, 1", (int)this.txtIdAsiento.Value));
							if ((int)this.cmbTipo.Value != 3)
							{
								MessageBox.Show("Documento Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.lblEstado.Text = "ANULADO";
							}
							else
							{
								MessageBox.Show("Documento Generado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.lblEstado.Text = "CONTABILIZADO";
							}
							this.CargaRegistro((int)this.txtIdAsiento.Value);

							bAnulado = true;
							
							this.btnAnular.Text = "Eliminar";

							if (!miAcceso.Eliminar) this.btnAnular.Enabled = false;
						}
					}
				}
			}
			else
			{				
				if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro de Borrar este Asiento?\n\nEste proceso NO PUEDE SER REVERTIDO", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2))
				{
					string sSQL = string.Format("Delete Asiento Where idAsiento = {0} Delete AsientoDetalle Where idAsiento = {0}", (int)this.txtIdAsiento.Value);
					Funcion.EjecutaSQL(cdsSeteoC, sSQL, true);

					MessageBox.Show("Documento Eliminado del Sistema", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.btnCancelar_Click(sender, e);
				}
			}
		}

		private void frmAsientos_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void frmAsientos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdAsiento.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
					if (e.KeyCode == Keys.A) if (miAcceso.Anular) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoC, "0402");

						Listas();

						this.CargaRegistro((int)this.txtIdAsiento.Value);						
					}
				}

				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btnCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
				if (e.KeyCode == Keys.F5) Listas();
			}
		}

		private void uGridDetalle_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (bEdicion)
			{	
				if (e.Cell.Column.ToString() == "Codigo")
					if (this.cmbOrigen.ActiveRow != null)
						if ((int)this.cmbOrigen.Value > 0)
							e.Cancel = true;				
				if (e.Cell.Column.ToString() == "Descripcion")
					if (this.cmbOrigen.ActiveRow != null)
						if ((int)this.cmbOrigen.Value > 0)
							e.Cancel = true;	
				if (e.Cell.Column.ToString() == "idProyecto")
					if (this.cmbOrigen.ActiveRow != null)
						if ((int)this.cmbOrigen.Value > 0)
							e.Cancel = true;				
				if (e.Cell.Column.ToString() == "idFormaPago")
					if (this.cmbOrigen.ActiveRow != null)
						if ((int)this.cmbOrigen.Value > 0)
							e.Cancel = true;
				if (e.Cell.Column.ToString() == "Detalle")
					if (this.cmbOrigen.ActiveRow != null)
						if ((int)this.cmbOrigen.Value > 0)
							e.Cancel = true;				
				if (e.Cell.Column.ToString() == "Debe")
					if (this.cmbOrigen.ActiveRow != null)
						if ((int)this.cmbOrigen.Value > 0)
							e.Cancel = true;
				if (e.Cell.Column.ToString() == "Haber")
					if (this.cmbOrigen.ActiveRow != null)
						if ((int)this.cmbOrigen.Value > 0)
							e.Cancel = true;
			}
		}

		private void txtChequeDeposito_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnDistribucion_Click(object sender, System.EventArgs e)
		{
			this.mnuDistribucion.Show(this.btnDistribucion, new Point(1, 20));
		}
		public void  Distribucion (int sTipo)
		{
			DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
			#region  distribucion 
			using (frmAsientoDistribucion AsiDistribucion = new frmAsientoDistribucion())
			{
				if (DialogResult.OK == AsiDistribucion.ShowDialog())
				{
					if (AsiDistribucion.uGridDetalle.Rows.Count > 0)
					{	
						bDistribucion=true;
						#region comentado
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in AsiDistribucion.uGridDetalle.Rows.All)
						{
							int iCuenta = (int)dr.Cells["idCuenta"].Value ;	
							int iFormaPago= (int)dr.Cells["idFormaPago"].Value ;
							string Detalle = (string)dr.Cells["Detalle"].Value ;
							decimal Debe = (decimal)dr.Cells["Debe"].Value;
							decimal Haber = (decimal)dr.Cells["Haber"].Value;									

							try
							{													
								string sSQLInf = string.Format("Exec GeneraAsientoDistribucion {0},{1},'{2}',{3},{4},'{5}',{6}", iCuenta,iFormaPago,Detalle,Debe,Haber,dtFechaFac.ToString("yyyyMMdd"),sTipo);
								SqlDataReader drInf = Funcion.rEscalarSQL(cdsSeteoF, sSQLInf, true);
								while (drInf.Read())
								{	
									bActualiza=false;
									this.uGridDetalle.Rows.Band.AddNew();
									this.uGridDetalle.ActiveRow.Cells["Codigo"].Value = drInf.GetString(4);
									this.uGridDetalle.ActiveRow.Cells["Descripcion"].Value = drInf.GetString(5);
									this.uGridDetalle.ActiveRow.Cells["idProyecto"].Value = drInf.GetInt32(6);
									this.uGridDetalle.ActiveRow.Cells["idFormaPago"].Value = drInf.GetInt32(7);
									this.uGridDetalle.ActiveRow.Cells["Detalle"].Value = drInf.GetString(8);
									this.uGridDetalle.ActiveRow.Cells["Debe"].Value = drInf.GetDecimal(9);
									this.uGridDetalle.ActiveRow.Cells["Haber"].Value = drInf.GetDecimal(10);
									bActualiza=true;
								}
								drInf.Close();	
								TotalDebeHaber();

								bNuevo = true;
								bEdicion = true;

								#region Controles
					
								this.btnNuevo.Enabled = false;
								this.btnConsultar.Enabled = false;
								this.btnGuardar.Enabled = true;
								this.btnCancelar.Enabled = true;
								this.btnDistribucion.Enabled = false;
								#endregion Controles
						
								bNuevo = true;
								bEdicion = true;
								bDistribucion=false;
					
								this.uGridDetalle.Rows.Band.AddNew();
							}
							catch(Exception Exc)
							{
								if (this.uGridDetalle.ActiveRow.Cells["Codigo"].Value == System.DBNull.Value)
								{
									MessageBox.Show("Cuenta no asignada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
									return;
								}
								
								MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

							}
						}
						#endregion comentado
								
						bEliminaAlValidar = true;
						Validacion.vEliminaFilasVacias(this.uGridDetalle, "idCuenta");
						bEliminaAlValidar = false;

					}
				}
			}
			
			#endregion  distribucion 
		}

		private void mnuAgencias_Click(object sender, System.EventArgs e)
		{
			Distribucion(1);
		}

		private void mnuAdministrativos_Click(object sender, System.EventArgs e)
		{
			Distribucion(2);
		}

		private void mnuGeneral_Click(object sender, System.EventArgs e)
		{
			Distribucion(3);
		}

		private void btnDuplicar_Click(object sender, System.EventArgs e)
		{
			int idAsientos = 0;
			idAsientos = (int)this.txtIdAsiento.Value;
			if ((int)this.txtIdAsiento.Value > 0 )
			{
				if((int) this.cmbOrigen.Value == 0)
				{
					this.btnNuevo_Click(sender,e);
					#region Consulta
					string sSQL = string.Format("Exec AsientoBuscaIndividual {0}", idAsientos);
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoC, sSQL, true);
					dr.Read();
					if (dr.HasRows)
					{
						this.txtIdAsiento.Value = 0;
						this.cmbTipo.Value = dr.GetInt32(1);	
						this.cmbOrigen.Value = dr.GetInt32(2);
						//if (dr.GetValue(3) != System.DBNull.Value) this.txtNumero.Text = dr.GetString(3);
						if (dr.GetValue(4) != System.DBNull.Value) this.txtDescripcion.Text = dr.GetString(4);
						this.txtRUC.Text = dr.GetString(5);
						this.txtChequeDeposito.Text = dr.GetString(6);
						this.txtDctoOrigen.Text = dr.GetString(7);
						//	this.dtFecha.Value = dr.GetDateTime(8);
						this.cmbCentroDeCosto.Value = dr.GetInt32(9); 
						this.txtNotas.Text = dr.GetString(10);
						this.txtDebe.Value = dr.GetDecimal(11);
						this.txtHaber.Value = dr.GetDecimal(12);
						bAnulado = dr.GetBoolean(13);
						//iEstado = dr.GetInt32(14); 
					}
					dr.Close();
					iEstado = 1;
					this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoC, string.Format("Exec AsientoConsultaDetalleDuplicado {0}", idAsientos));
					#endregion Consulta
					this.btnDuplicar.Enabled = false;
				}
			}
		}

		private void mnuAgenciasIg_Click(object sender, System.EventArgs e)
		{
			Distribucion(4);
		}

		private void mnuAdministrativosIg_Click(object sender, System.EventArgs e)
		{
			Distribucion(5);
		}

		private void mnuGeneralIg_Click(object sender, System.EventArgs e)
		{
			Distribucion(6);
		}
	}
}

