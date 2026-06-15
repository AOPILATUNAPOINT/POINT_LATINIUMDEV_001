using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmConciliacionAutomatica.
	/// </summary>
	public class frmConciliacionAutomatica : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtHasta;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor dtDesde;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private System.Windows.Forms.Label lblComprobante;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSalida;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEntrada;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridECSalidas;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridECEntradas;
		private System.Windows.Forms.Button btnComparar;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiferencias;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoBancario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoConciliado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoNoConciliado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoEnLibros;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNoConciliadosSalida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNoConciliadosEntrada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtConciliadosSalidas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtConciliadosEntrada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSistemaSalida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSistemaEntrada;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdConciliacionBancaria;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Button btnAbrir;
		private System.Windows.Forms.Button btnImportar;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Button btnGuardar;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConciliacionAutomatica()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoDeCuentaBancaria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idConciliacionBancaria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 7, true);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoDeCuentaBancaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idConciliacionBancaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoDeCuentaBancaria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idConciliacionBancaria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 7, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoDeCuentaBancaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idConciliacionBancaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Movimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Confirmado");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 7, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Movimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Confirmado");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 7, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConciliacionAutomatica));
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			this.uGridECSalidas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnImportar = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.dtDesde = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.uGridECEntradas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSalida = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridEntrada = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnComparar = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtDiferencias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoBancario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoConciliado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoNoConciliado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoEnLibros = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNoConciliadosSalida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNoConciliadosEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConciliadosSalidas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConciliadosEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSistemaSalida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSistemaEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnAbrir = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.txtIdConciliacionBancaria = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnGuardar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridECSalidas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridECEntradas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiferencias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoBancario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoConciliado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoNoConciliado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoEnLibros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosSalidas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdConciliacionBancaria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridECSalidas
			// 
			this.uGridECSalidas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridECSalidas.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridECSalidas.DisplayLayout.Appearance = appearance1;
			this.uGridECSalidas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 67;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 56;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 7;
			ultraGridColumn3.Width = 33;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 26;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 72;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 77;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 77;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance2;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Width = 66;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand1.Header.Caption = "Movimientos De Salida Estado de Cuenta";
			ultraGridBand1.HeaderVisible = true;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance3;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridECSalidas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridECSalidas.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridECSalidas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridECSalidas.DisplayLayout.Override.CardAreaAppearance = appearance5;
			this.uGridECSalidas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8.5F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridECSalidas.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridECSalidas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridECSalidas.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridECSalidas.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridECSalidas.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridECSalidas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridECSalidas.Location = new System.Drawing.Point(368, 416);
			this.uGridECSalidas.Name = "uGridECSalidas";
			this.uGridECSalidas.Size = new System.Drawing.Size(346, 272);
			this.uGridECSalidas.TabIndex = 240;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// btnImportar
			// 
			this.btnImportar.CausesValidation = false;
			this.btnImportar.Enabled = false;
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(432, 88);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.TabIndex = 241;
			this.btnImportar.Text = "Importar";
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Enabled = false;
			this.label1.Location = new System.Drawing.Point(232, 88);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 16);
			this.label1.TabIndex = 245;
			this.label1.Text = "Hasta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Enabled = false;
			this.label2.Location = new System.Drawing.Point(16, 88);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(35, 16);
			this.label2.TabIndex = 244;
			this.label2.Text = "Desde";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			this.dtHasta.DateTime = new System.DateTime(1753, 1, 1, 0, 0, 0, 0);
			this.dtHasta.Enabled = false;
			this.dtHasta.FormatString = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(288, 88);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.Size = new System.Drawing.Size(120, 22);
			this.dtHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.dtHasta.TabIndex = 243;
			this.dtHasta.Value = null;
			// 
			// dtDesde
			// 
			this.dtDesde.DateTime = new System.DateTime(1753, 1, 1, 0, 0, 0, 0);
			this.dtDesde.Enabled = false;
			this.dtDesde.FormatString = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(80, 88);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.Size = new System.Drawing.Size(120, 22);
			this.dtDesde.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.dtDesde.TabIndex = 242;
			this.dtDesde.Value = null;
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.Caption = "Cuenta";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 360;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(80, 56);
			this.cmbCuenta.MaxDropDownItems = 30;
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(328, 21);
			this.cmbCuenta.TabIndex = 650;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.ValueChanged += new System.EventHandler(this.cmbCuenta_ValueChanged);
			this.cmbCuenta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuenta_InitializeLayout);
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Enabled = false;
			this.lblComprobante.Location = new System.Drawing.Point(16, 56);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(39, 16);
			this.lblComprobante.TabIndex = 651;
			this.lblComprobante.Text = "Cuenta";
			this.lblComprobante.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridECEntradas
			// 
			this.uGridECEntradas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridECEntradas.DataSource = this.ultraDataSource2;
			appearance9.BackColor = System.Drawing.Color.White;
			this.uGridECEntradas.DisplayLayout.Appearance = appearance9;
			this.uGridECEntradas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 67;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 56;
			ultraGridColumn13.Header.Caption = "...";
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Width = 33;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 26;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Width = 72;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Width = 77;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn17.Width = 77;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance10;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 6;
			ultraGridColumn18.Width = 66;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			ultraGridBand3.Header.Caption = "Movimientos De Entrada Estado de Cuenta";
			ultraGridBand3.HeaderVisible = true;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance11;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings2});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridECEntradas.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridECEntradas.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.uGridECEntradas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.uGridECEntradas.DisplayLayout.Override.CardAreaAppearance = appearance13;
			this.uGridECEntradas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 8.5F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridECEntradas.DisplayLayout.Override.HeaderAppearance = appearance14;
			this.uGridECEntradas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridECEntradas.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridECEntradas.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.uGridECEntradas.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.uGridECEntradas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridECEntradas.Location = new System.Drawing.Point(368, 136);
			this.uGridECEntradas.Name = "uGridECEntradas";
			this.uGridECEntradas.Size = new System.Drawing.Size(346, 272);
			this.uGridECEntradas.TabIndex = 654;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(bool);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.DateTime);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn25.DataType = typeof(bool);
			ultraDataColumn25.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25});
			// 
			// uGridSalida
			// 
			this.uGridSalida.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSalida.DataSource = this.ultraDataSource3;
			appearance17.BackColor = System.Drawing.Color.White;
			this.uGridSalida.DisplayLayout.Appearance = appearance17;
			this.uGridSalida.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 95;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 115;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Width = 55;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 49;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.Caption = "Tipo";
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.Width = 37;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.Caption = "Número";
			ultraGridColumn24.Header.VisiblePosition = 5;
			ultraGridColumn24.Width = 97;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.Caption = "Concepto";
			ultraGridColumn25.Header.VisiblePosition = 6;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 350;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance18;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 7;
			ultraGridColumn26.Width = 107;
			ultraGridColumn27.Header.Caption = "...";
			ultraGridColumn27.Header.VisiblePosition = 8;
			ultraGridColumn27.Width = 29;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			ultraGridBand4.Header.Caption = "Movimientos de Salida Sistema";
			ultraGridBand4.HeaderVisible = true;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance19;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings3});
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridSalida.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSalida.DisplayLayout.Override.ActiveRowAppearance = appearance20;
			this.uGridSalida.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSalida.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance21.BackColor = System.Drawing.Color.Transparent;
			this.uGridSalida.DisplayLayout.Override.CardAreaAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance22.FontData.BoldAsString = "True";
			appearance22.FontData.Name = "Arial";
			appearance22.FontData.SizeInPoints = 8.25F;
			appearance22.ForeColor = System.Drawing.Color.White;
			appearance22.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSalida.DisplayLayout.Override.HeaderAppearance = appearance22;
			this.uGridSalida.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance23.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance23.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalida.DisplayLayout.Override.RowAlternateAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalida.DisplayLayout.Override.RowSelectorAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalida.DisplayLayout.Override.SelectedRowAppearance = appearance25;
			this.uGridSalida.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSalida.Location = new System.Drawing.Point(8, 416);
			this.uGridSalida.Name = "uGridSalida";
			this.uGridSalida.Size = new System.Drawing.Size(346, 272);
			this.uGridSalida.TabIndex = 656;
			this.uGridSalida.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.uGridSalida.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSalida_CellChange);
			this.uGridSalida.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSalida_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(System.DateTime);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn33.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn34.DataType = typeof(bool);
			ultraDataColumn34.DefaultValue = false;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34});
			// 
			// uGridEntrada
			// 
			this.uGridEntrada.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEntrada.DataSource = this.ultraDataSource1;
			appearance26.BackColor = System.Drawing.Color.White;
			this.uGridEntrada.DisplayLayout.Appearance = appearance26;
			this.uGridEntrada.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 95;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 115;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 2;
			ultraGridColumn30.Width = 72;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 3;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 50;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.Caption = "Tipo";
			ultraGridColumn32.Header.VisiblePosition = 4;
			ultraGridColumn32.Width = 49;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.Caption = "Número";
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn33.Width = 88;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn34.Header.Caption = "Concepto";
			ultraGridColumn34.Header.VisiblePosition = 6;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 350;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance27;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 7;
			ultraGridColumn35.Width = 88;
			ultraGridColumn36.Header.Caption = "...";
			ultraGridColumn36.Header.VisiblePosition = 8;
			ultraGridColumn36.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn36.Width = 28;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36});
			ultraGridBand5.Header.Caption = "Movimientos de Entrada Sistema";
			ultraGridBand5.HeaderVisible = true;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance28;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4});
			ultraGridBand5.SummaryFooterCaption = "";
			this.uGridEntrada.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEntrada.DisplayLayout.Override.ActiveRowAppearance = appearance29;
			this.uGridEntrada.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridEntrada.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance30.BackColor = System.Drawing.Color.Transparent;
			this.uGridEntrada.DisplayLayout.Override.CardAreaAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance31.FontData.BoldAsString = "True";
			appearance31.FontData.Name = "Arial";
			appearance31.FontData.SizeInPoints = 8.25F;
			appearance31.ForeColor = System.Drawing.Color.White;
			appearance31.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEntrada.DisplayLayout.Override.HeaderAppearance = appearance31;
			this.uGridEntrada.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance32.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance32.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.RowAlternateAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.RowSelectorAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.SelectedRowAppearance = appearance34;
			this.uGridEntrada.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEntrada.Location = new System.Drawing.Point(8, 136);
			this.uGridEntrada.Name = "uGridEntrada";
			this.uGridEntrada.Size = new System.Drawing.Size(346, 272);
			this.uGridEntrada.TabIndex = 655;
			this.uGridEntrada.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.uGridEntrada.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridEntrada_CellChange);
			this.uGridEntrada.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridEntrada_InitializeLayout);
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
			// btnComparar
			// 
			this.btnComparar.CausesValidation = false;
			this.btnComparar.Enabled = false;
			this.btnComparar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComparar.Location = new System.Drawing.Point(528, 88);
			this.btnComparar.Name = "btnComparar";
			this.btnComparar.TabIndex = 658;
			this.btnComparar.Text = "Comparar";
			this.btnComparar.Click += new System.EventHandler(this.btnComparar_Click);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label13.Location = new System.Drawing.Point(728, 608);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(59, 16);
			this.label13.TabIndex = 700;
			this.label13.Text = "Diferencias";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label14.Location = new System.Drawing.Point(728, 584);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(78, 16);
			this.label14.TabIndex = 699;
			this.label14.Text = "Saldo Bancario";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label15.Location = new System.Drawing.Point(728, 560);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(86, 16);
			this.label15.TabIndex = 698;
			this.label15.Text = "Saldo Conciliado";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label16.Location = new System.Drawing.Point(728, 536);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(103, 16);
			this.label16.TabIndex = 697;
			this.label16.Text = "Saldo No Conciliado";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label18.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label18.Location = new System.Drawing.Point(728, 480);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(92, 17);
			this.label18.TabIndex = 696;
			this.label18.Text = "CONCILIACIÓN";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label19.Location = new System.Drawing.Point(728, 512);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(81, 16);
			this.label19.TabIndex = 691;
			this.label19.Text = "Saldo en Libros";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label8.Location = new System.Drawing.Point(728, 416);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(39, 16);
			this.label8.TabIndex = 689;
			this.label8.Text = "Salidas";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label9.Location = new System.Drawing.Point(728, 392);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(48, 16);
			this.label9.TabIndex = 688;
			this.label9.Text = "Entradas";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label5.Location = new System.Drawing.Point(728, 312);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(39, 16);
			this.label5.TabIndex = 687;
			this.label5.Text = "Salidas";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label7.Location = new System.Drawing.Point(728, 288);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(48, 16);
			this.label7.TabIndex = 686;
			this.label7.Text = "Entradas";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label12.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label12.Location = new System.Drawing.Point(728, 360);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(107, 17);
			this.label12.TabIndex = 685;
			this.label12.Text = "NO CONCILIADOS";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label11.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label11.Location = new System.Drawing.Point(728, 256);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(86, 17);
			this.label11.TabIndex = 684;
			this.label11.Text = "CONCILIADOS";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label10.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label10.Location = new System.Drawing.Point(728, 144);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(56, 17);
			this.label10.TabIndex = 683;
			this.label10.Text = "SISTEMA";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label4.Location = new System.Drawing.Point(728, 200);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 678;
			this.label4.Text = "Salidas";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Location = new System.Drawing.Point(728, 176);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(48, 16);
			this.label3.TabIndex = 676;
			this.label3.Text = "Entradas";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiferencias
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiferencias.Appearance = appearance35;
			this.txtDiferencias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiferencias.Enabled = false;
			this.txtDiferencias.FormatString = "#,##0.00";
			this.txtDiferencias.Location = new System.Drawing.Point(848, 608);
			this.txtDiferencias.Name = "txtDiferencias";
			this.txtDiferencias.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDiferencias.PromptChar = ' ';
			this.txtDiferencias.Size = new System.Drawing.Size(112, 22);
			this.txtDiferencias.TabIndex = 695;
			// 
			// txtSaldoBancario
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoBancario.Appearance = appearance36;
			this.txtSaldoBancario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoBancario.Enabled = false;
			this.txtSaldoBancario.FormatString = "#,##0.00";
			this.txtSaldoBancario.Location = new System.Drawing.Point(848, 584);
			this.txtSaldoBancario.Name = "txtSaldoBancario";
			this.txtSaldoBancario.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoBancario.PromptChar = ' ';
			this.txtSaldoBancario.Size = new System.Drawing.Size(112, 22);
			this.txtSaldoBancario.TabIndex = 694;
			this.txtSaldoBancario.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSaldoBancario_KeyDown);
			this.txtSaldoBancario.Validated += new System.EventHandler(this.txtSaldoBancario_Validated);
			// 
			// txtSaldoConciliado
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoConciliado.Appearance = appearance37;
			this.txtSaldoConciliado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoConciliado.Enabled = false;
			this.txtSaldoConciliado.FormatString = "#,##0.00";
			this.txtSaldoConciliado.Location = new System.Drawing.Point(848, 560);
			this.txtSaldoConciliado.Name = "txtSaldoConciliado";
			this.txtSaldoConciliado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoConciliado.PromptChar = ' ';
			this.txtSaldoConciliado.Size = new System.Drawing.Size(112, 22);
			this.txtSaldoConciliado.TabIndex = 693;
			// 
			// txtSaldoNoConciliado
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoNoConciliado.Appearance = appearance38;
			this.txtSaldoNoConciliado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoNoConciliado.Enabled = false;
			this.txtSaldoNoConciliado.FormatString = "#,##0.00";
			this.txtSaldoNoConciliado.Location = new System.Drawing.Point(848, 536);
			this.txtSaldoNoConciliado.Name = "txtSaldoNoConciliado";
			this.txtSaldoNoConciliado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoNoConciliado.PromptChar = ' ';
			this.txtSaldoNoConciliado.Size = new System.Drawing.Size(112, 22);
			this.txtSaldoNoConciliado.TabIndex = 692;
			// 
			// txtSaldoEnLibros
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoEnLibros.Appearance = appearance39;
			this.txtSaldoEnLibros.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoEnLibros.Enabled = false;
			this.txtSaldoEnLibros.FormatString = "#,##0.00";
			this.txtSaldoEnLibros.Location = new System.Drawing.Point(848, 512);
			this.txtSaldoEnLibros.Name = "txtSaldoEnLibros";
			this.txtSaldoEnLibros.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoEnLibros.PromptChar = ' ';
			this.txtSaldoEnLibros.Size = new System.Drawing.Size(112, 22);
			this.txtSaldoEnLibros.TabIndex = 690;
			// 
			// txtNoConciliadosSalida
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNoConciliadosSalida.Appearance = appearance40;
			this.txtNoConciliadosSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNoConciliadosSalida.Enabled = false;
			this.txtNoConciliadosSalida.FormatString = "#,##0.00";
			this.txtNoConciliadosSalida.Location = new System.Drawing.Point(848, 416);
			this.txtNoConciliadosSalida.Name = "txtNoConciliadosSalida";
			this.txtNoConciliadosSalida.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNoConciliadosSalida.PromptChar = ' ';
			this.txtNoConciliadosSalida.Size = new System.Drawing.Size(112, 22);
			this.txtNoConciliadosSalida.TabIndex = 682;
			// 
			// txtNoConciliadosEntrada
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNoConciliadosEntrada.Appearance = appearance41;
			this.txtNoConciliadosEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNoConciliadosEntrada.Enabled = false;
			this.txtNoConciliadosEntrada.FormatString = "#,##0.00";
			this.txtNoConciliadosEntrada.Location = new System.Drawing.Point(848, 392);
			this.txtNoConciliadosEntrada.Name = "txtNoConciliadosEntrada";
			this.txtNoConciliadosEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNoConciliadosEntrada.PromptChar = ' ';
			this.txtNoConciliadosEntrada.Size = new System.Drawing.Size(112, 22);
			this.txtNoConciliadosEntrada.TabIndex = 681;
			// 
			// txtConciliadosSalidas
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtConciliadosSalidas.Appearance = appearance42;
			this.txtConciliadosSalidas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConciliadosSalidas.Enabled = false;
			this.txtConciliadosSalidas.FormatString = "#,##0.00";
			this.txtConciliadosSalidas.Location = new System.Drawing.Point(848, 312);
			this.txtConciliadosSalidas.Name = "txtConciliadosSalidas";
			this.txtConciliadosSalidas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtConciliadosSalidas.PromptChar = ' ';
			this.txtConciliadosSalidas.Size = new System.Drawing.Size(112, 22);
			this.txtConciliadosSalidas.TabIndex = 680;
			// 
			// txtConciliadosEntrada
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtConciliadosEntrada.Appearance = appearance43;
			this.txtConciliadosEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConciliadosEntrada.Enabled = false;
			this.txtConciliadosEntrada.FormatString = "#,##0.00";
			this.txtConciliadosEntrada.Location = new System.Drawing.Point(848, 288);
			this.txtConciliadosEntrada.Name = "txtConciliadosEntrada";
			this.txtConciliadosEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtConciliadosEntrada.PromptChar = ' ';
			this.txtConciliadosEntrada.Size = new System.Drawing.Size(112, 22);
			this.txtConciliadosEntrada.TabIndex = 679;
			// 
			// txtSistemaSalida
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSistemaSalida.Appearance = appearance44;
			this.txtSistemaSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSistemaSalida.Enabled = false;
			this.txtSistemaSalida.FormatString = "#,##0.00";
			this.txtSistemaSalida.Location = new System.Drawing.Point(848, 200);
			this.txtSistemaSalida.Name = "txtSistemaSalida";
			this.txtSistemaSalida.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSistemaSalida.PromptChar = ' ';
			this.txtSistemaSalida.Size = new System.Drawing.Size(112, 22);
			this.txtSistemaSalida.TabIndex = 677;
			// 
			// txtSistemaEntrada
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSistemaEntrada.Appearance = appearance45;
			this.txtSistemaEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSistemaEntrada.Enabled = false;
			this.txtSistemaEntrada.FormatString = "#,##0.00";
			this.txtSistemaEntrada.Location = new System.Drawing.Point(848, 176);
			this.txtSistemaEntrada.Name = "txtSistemaEntrada";
			this.txtSistemaEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSistemaEntrada.PromptChar = ' ';
			this.txtSistemaEntrada.Size = new System.Drawing.Size(112, 22);
			this.txtSistemaEntrada.TabIndex = 675;
			// 
			// btnEditar
			// 
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(170, 8);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 709;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnAbrir
			// 
			this.btnAbrir.CausesValidation = false;
			this.btnAbrir.Enabled = false;
			this.btnAbrir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAbrir.Location = new System.Drawing.Point(499, 8);
			this.btnAbrir.Name = "btnAbrir";
			this.btnAbrir.Size = new System.Drawing.Size(112, 23);
			this.btnAbrir.TabIndex = 707;
			this.btnAbrir.Text = "Cerrar Conciliación";
			this.btnAbrir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAbrir.Click += new System.EventHandler(this.btnAbrir_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(86, 8);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 706;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 8);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 701;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(252, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 702;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(417, 8);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 703;
			this.btnAnular.Text = "Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(617, 8);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 704;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(695, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 705;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(968, 8);
			this.groupBox1.TabIndex = 710;
			this.groupBox1.TabStop = false;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(784, 56);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 712;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance46;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(848, 56);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 711;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(784, 88);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 713;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(728, 656);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 17);
			this.lblUsuario.TabIndex = 714;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtIdConciliacionBancaria
			// 
			this.txtIdConciliacionBancaria.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdConciliacionBancaria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdConciliacionBancaria.Enabled = false;
			this.txtIdConciliacionBancaria.Location = new System.Drawing.Point(928, 664);
			this.txtIdConciliacionBancaria.Name = "txtIdConciliacionBancaria";
			this.txtIdConciliacionBancaria.PromptChar = ' ';
			this.txtIdConciliacionBancaria.Size = new System.Drawing.Size(16, 22);
			this.txtIdConciliacionBancaria.TabIndex = 716;
			this.txtIdConciliacionBancaria.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(944, 664);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 715;
			this.txtEstado.Visible = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 120);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(968, 8);
			this.groupBox2.TabIndex = 717;
			this.groupBox2.TabStop = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(335, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 718;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// frmConciliacionAutomatica
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(968, 694);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.txtIdConciliacionBancaria);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnAbrir);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtDiferencias);
			this.Controls.Add(this.txtSaldoBancario);
			this.Controls.Add(this.txtSaldoConciliado);
			this.Controls.Add(this.txtSaldoNoConciliado);
			this.Controls.Add(this.txtSaldoEnLibros);
			this.Controls.Add(this.txtNoConciliadosSalida);
			this.Controls.Add(this.txtNoConciliadosEntrada);
			this.Controls.Add(this.txtConciliadosSalidas);
			this.Controls.Add(this.txtConciliadosEntrada);
			this.Controls.Add(this.txtSistemaSalida);
			this.Controls.Add(this.txtSistemaEntrada);
			this.Controls.Add(this.btnComparar);
			this.Controls.Add(this.uGridECEntradas);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.uGridECSalidas);
			this.Controls.Add(this.uGridSalida);
			this.Controls.Add(this.uGridEntrada);
			this.MaximizeBox = false;
			this.Name = "frmConciliacionAutomatica";
			this.Text = "Conciliación Bancaria Automatica";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmConciliacionAutomatica_Closing);
			this.Load += new System.EventHandler(this.frmConciliacionAutomatica_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridECSalidas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridECEntradas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiferencias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoBancario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoConciliado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoNoConciliado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoEnLibros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosSalidas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdConciliacionBancaria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		string path = "";

		public bool CargaMovimientosEntrada(string stArchivo)
		{ 
			path = stArchivo;
 
			if(File.Exists(path))
			{
				OleDbConnection oConn = new OleDbConnection();
 
				OleDbCommand oCmd = new OleDbCommand();
 
				OleDbDataAdapter oDa = new OleDbDataAdapter();
 
				DataSet oDs = new DataSet();

				oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "; Extended Properties= Excel 8.0;";
 
				oConn.Open();
 
				oCmd.CommandText = "SELECT * FROM [ENERO$] WHERE REFERENCIA IN ('DP', 'NC')";
 
				oCmd.Connection = oConn;
 
				oDa.SelectCommand = oCmd;
 
				oDa.Fill(oDs, "EstadoDeCuenta");
 
				this.uGridECEntradas.DataSource = oDs.Tables[0].DefaultView;
 
				this.uGridECEntradas.DataBind();
 
				oConn.Close();
 
				return true;
 
			}
			else
			{
				return false; 
			}
		}

		public bool CargaMovimientosSalida(string stArchivo) 
		{ 
			path = stArchivo;
 
			if(File.Exists(path))
			{
				OleDbConnection oConn = new OleDbConnection();
 
				OleDbCommand oCmd = new OleDbCommand();
 
				OleDbDataAdapter oDa = new OleDbDataAdapter();
 
				DataSet oDs = new DataSet();

				oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "; Extended Properties= Excel 8.0;";
 
				oConn.Open();
 
				oCmd.CommandText = "SELECT * FROM [ENERO$] WHERE REFERENCIA IN ('CH', 'ND')";
 
				oCmd.Connection = oConn;
 
				oDa.SelectCommand = oCmd;
 
				oDa.Fill(oDs, "EstadoDeCuenta");
 
				foreach (DataRow row in oDs.Tables[0].Rows)
				{
					string sSQLInsertaEC = string.Format("Exec CBInsertaEstadoDeCuenta '{0}', '{1}', '{2}', {3}", (DateTime)row["FECHA"], row["REFERENCIA"].ToString(), row["NUMERO"].ToString(), Convert.ToDecimal(row["VALOR"]));
					Funcion.EjecutaSQL(cdsSeteoF, sSQLInsertaEC);
				}
				//				this.uGridECSalidas.DataSource = oDs.Tables[0].DefaultView;
				// 
				//				this.uGridECSalidas.DataBind();
 
				oConn.Close();
 
				return true;
 
			}
			else
			{
				return false; 
			}
		}

		public bool InsertaEstadoDeCuenta(string stArchivo, int idConciliacionBancaria, int idCuenta) 
		{ 
			path = stArchivo;
 
			if(File.Exists(path))
			{
				OleDbConnection oConn = new OleDbConnection();
 
				OleDbCommand oCmd = new OleDbCommand();
 
				OleDbDataAdapter oDa = new OleDbDataAdapter();
 
				DataSet oDs = new DataSet();

				oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "; Extended Properties= Excel 8.0;";
 
				oConn.Open();
 
				oCmd.CommandText = "SELECT * FROM [ESTADODECUENTA$]";
 
				oCmd.Connection = oConn;
 
				oDa.SelectCommand = oCmd;
 
				oDa.Fill(oDs, "EstadoDeCuenta");
 
				foreach (DataRow row in oDs.Tables[0].Rows)
				{
					DateTime dtFecha = DateTime.Today;
					string sReferencia = "";
					string sNumero = "";
					decimal dValor = 0.00m;
					
					if (row["FECHA"] != System.DBNull.Value) dtFecha = (DateTime)row["FECHA"];
					if (row["REFERENCIA"] != System.DBNull.Value) sReferencia = row["REFERENCIA"].ToString();
					if (row["NUMERO"] != System.DBNull.Value) sNumero = row["NUMERO"].ToString();
					if (row["VALOR"] != System.DBNull.Value) dValor = decimal.Parse(row["VALOR"].ToString()); //Convert.ToDecimal(row["VALOR"]);

					string sSQLInsertaEC = string.Format("Exec CBInsertaEstadoDeCuenta {0}, {1}, '{2}', '{3}', '{4}', {5}", 
						idConciliacionBancaria, idCuenta, dtFecha.ToString("yyyyMMdd"), sReferencia, sNumero, dValor);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLInsertaEC);
				}
				
				oConn.Close();
 
				return true;
 
			}
			else
			{
				return false; 
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;

		private void HabiltaGrids(bool bModifica)
		{
			this.uGridEntrada.Enabled = true;
			this.uGridSalida.Enabled = true;
			this.uGridECEntradas.Enabled = true;
			this.uGridECSalidas.Enabled = true;

			if (bModifica)
			{
				this.uGridEntrada.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridSalida.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridECEntradas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridECSalidas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				this.uGridEntrada.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridSalida.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridECEntradas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridECSalidas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			}
		}
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmConciliacionAutomatica_Load(object sender, System.EventArgs e)
		{

			miAcceso = new Acceso(cdsSeteoF, "040810");

			if (!Funcion.Permiso("919", cdsSeteoF))
			{				
				MessageBox.Show("No puede tiene Acceso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			
			
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			this.HabiltaGrids(false);
			
			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select c.idCuenta, c.Descripcion From Cuenta c Inner Join CuentaCorriente cc On cc.idCuenta = c.idCuenta Order By c.Descripcion");
		}

		private void btnComparar_Click(object sender, System.EventArgs e)
		{
			using (frmComparaEstadosDeCuenta Compara = new frmComparaEstadosDeCuenta(this.uGridEntrada, this.uGridSalida, this.uGridECEntradas, this.uGridECSalidas))
			{				
				if (DialogResult.OK == Compara.ShowDialog())
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drEntradas in Compara.uGridEntradaConciliados.Rows.All)
					{
						string sSQLEntradas = string.Format("Exec GuardaDetalleConciliacionBancariaAutomatica {0}, {1}", 
							(int)this.txtIdConciliacionBancaria.Value, (int)drEntradas.Cells["idDetAsiento"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLEntradas, true);
					}

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drSalidas in Compara.uGridSalidaConciliados.Rows.All)
					{
						string sSQLSalidas = string.Format("Exec GuardaDetalleConciliacionBancariaAutomatica {0}, {1}", 
							(int)this.txtIdConciliacionBancaria.Value, (int)drSalidas.Cells["idDetAsiento"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLSalidas, true);
					}
					
					this.EntradasSalidas();

					string sSQLEntradasA = string.Format("Exec CBConsultaEstadoDeCuenta {0}, {1}, 0", (int)this.txtIdConciliacionBancaria.Value, (int)this.cmbCuenta.Value);
					this.uGridECEntradas.DataSource = Funcion.dvProcedimiento (cdsSeteoF, sSQLEntradasA);
					string sSQLSalidasA = string.Format("Exec CBConsultaEstadoDeCuenta {0}, {1}, 1", (int)this.txtIdConciliacionBancaria.Value, (int)this.cmbCuenta.Value);
					this.uGridECSalidas.DataSource = Funcion.dvProcedimiento (cdsSeteoF, sSQLSalidasA);

					this.btnGuardar.Enabled = true;
				}
			}
		}

		private void EntradasSalidas()
		{
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;
      			
			this.uGridEntrada.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConciliacionBancariaMovimientos {0}, '{1}', '{2}', '{3}', {4}, 0", (int)this.cmbCuenta.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), "", (int)this.txtIdConciliacionBancaria.Value));
			this.uGridSalida.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConciliacionBancariaMovimientos {0}, '{1}', '{2}', '{3}', {4}, 1", (int)this.cmbCuenta.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), "", (int)this.txtIdConciliacionBancaria.Value));      

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Exec SumaValoresConciliacionBancaria {0}, '{1}', '{2}'", (int)this.cmbCuenta.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd")), true);
			dr.Read();
			if (dr.HasRows)
			{				
				this.txtSistemaEntrada.Value = Convert.ToDecimal(dr.GetValue(0));
				this.txtSistemaSalida.Value = Convert.ToDecimal(dr.GetValue(1));
			}
			dr.Close();

			if ((int)this.txtEstado.Value != 2) Resumen();
		}

		private void Resumen()
		{
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			string sSQLResumen = string.Format("Exec ResumenConciliacionBancaria {0}, '{1}', '{2}', {3}", 
				(int)this.cmbCuenta.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int)this.txtIdConciliacionBancaria.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLResumen, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtConciliadosEntrada.Value = Convert.ToDecimal(dr.GetValue(0));
				this.txtConciliadosSalidas.Value = Convert.ToDecimal(dr.GetValue(1));
				this.txtNoConciliadosEntrada.Value = Convert.ToDecimal(dr.GetValue(2));
				this.txtNoConciliadosSalida.Value = Convert.ToDecimal(dr.GetValue(3));
				this.txtSaldoEnLibros.Value = Convert.ToDecimal(dr.GetValue(4));
			}
			dr.Close();
			
			this.txtSaldoNoConciliado.Value = Convert.ToDecimal(this.txtNoConciliadosEntrada.Value) - Convert.ToDecimal(this.txtNoConciliadosSalida.Value);
			this.txtSaldoConciliado.Value = Convert.ToDecimal(this.txtSaldoEnLibros.Value) + Math.Abs(Convert.ToDecimal(this.txtSaldoNoConciliado.Value));
			this.txtDiferencias.Value = Convert.ToDecimal(this.txtSaldoConciliado.Value) - Convert.ToDecimal(this.txtSaldoBancario.Value);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cmbCuenta_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.dtDesde.Value = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("Exec RetornaFechaConciliacion {0}", (int)this.cmbCuenta.Value)); 
				DateTime dtDesde = (DateTime)this.dtDesde.Value;
				this.dtHasta.Value = DateTime.DaysInMonth(dtDesde.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString();			
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bNuevo)
			{
				string sSQL = string.Format("Exec CBCancela {0}", (int)this.txtIdConciliacionBancaria.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}

			bNuevo = false;
			bEdicion = false;

			this.txtIdConciliacionBancaria.Value = 0;
			this.txtEstado.Value = 0;
			this.txtNumero.Text = "";
			this.cmbCuenta.Value = System.DBNull.Value;
			this.dtDesde.Value = System.DBNull.Value;
			this.dtHasta.Value = System.DBNull.Value;
			this.txtSistemaEntrada.Value = 0;
			this.txtSistemaSalida.Value = 0;
			this.txtConciliadosEntrada.Value = 0;
			this.txtConciliadosSalidas.Value = 0;
			this.txtNoConciliadosEntrada.Value = 0;
			this.txtNoConciliadosSalida.Value = 0;
			this.txtSaldoEnLibros.Value = 0;
			this.txtSaldoNoConciliado.Value = 0;
			this.txtSaldoConciliado.Value = 0;
			this.txtSaldoBancario.Value = 0;
			this.txtDiferencias.Value = 0;
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
 
			this.cmbCuenta.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			this.txtSaldoBancario.Enabled = false;

			this.uGridEntrada.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConciliacionBancariaMovimientos 0, '', '', '', 0, 0");
			this.uGridSalida.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConciliacionBancariaMovimientos 0, '', '', '', 0, 1");

			this.uGridECEntradas.DataSource = Funcion.dvProcedimiento (cdsSeteoF, "Exec CBConsultaEstadoDeCuenta 0, 0, 0");			
			this.uGridECSalidas.DataSource = Funcion.dvProcedimiento (cdsSeteoF, "Exec CBConsultaEstadoDeCuenta 0, 0, 1");

			this.HabiltaGrids(false);
			
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;			
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;			
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnAbrir.Enabled = false;	
			this.btnCancelar.Enabled = false;
			
			this.btnImportar.Enabled = false;
			this.btnComparar.Enabled = false;			
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.cmbCuenta.Enabled = true;
			this.dtDesde.Enabled = true;
			this.dtHasta.Enabled = true;
			this.txtSaldoBancario.Enabled = true;

			this.btnImportar.Enabled = true;
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;			
			this.btnAnular.Enabled = false;
			this.btnAbrir.Enabled = false;
			this.btnCancelar.Enabled = true;

			this.HabiltaGrids(true);			

			this.cmbCuenta.Focus();

			bNuevo = true;
			bEdicion = true;
		}

		private void Consulta(int idConciliacionBancaria)
		{			
			try
			{
				string sSQL = string.Format("Exec CBConsultaIndividual {0}", idConciliacionBancaria);
				SqlDataReader drBusca = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drBusca.Read();
				if (drBusca.HasRows)
				{					
					this.txtIdConciliacionBancaria.Value = drBusca.GetInt32(0);
					this.cmbCuenta.Value = drBusca.GetInt32(1);
					this.dtDesde.Value = drBusca.GetDateTime(4);
					this.dtHasta.Value = drBusca.GetDateTime(5);
					this.txtNumero.Value = drBusca.GetString(6);					
					this.txtSaldoBancario.Value = drBusca.GetDecimal(7);
					this.txtEstado.Value = drBusca.GetInt32(8);
					this.lblEstado.Text = drBusca.GetString(9);
					this.lblUsuario.Text = drBusca.GetString(10);
				}
				drBusca.Close();

				EntradasSalidas();
				
				string sSQLEntradas = string.Format("Exec CBConsultaEstadoDeCuenta {0}, {1}, 0", (int)this.txtIdConciliacionBancaria.Value, (int)this.cmbCuenta.Value);
				this.uGridECEntradas.DataSource = Funcion.dvProcedimiento (cdsSeteoF, sSQLEntradas);
				
				string sSQLSalidas = string.Format("Exec CBConsultaEstadoDeCuenta {0}, {1}, 1", (int)this.txtIdConciliacionBancaria.Value, (int)this.cmbCuenta.Value);
				this.uGridECSalidas.DataSource = Funcion.dvProcedimiento (cdsSeteoF, sSQLSalidas);
						
				if ((int)this.txtEstado.Value == 0) 
				{
					if (miAcceso.BBloquearCB) this.btnAbrir.Enabled = true;
					this.btnAbrir.Text = "Cerrar Conciliación";
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.Anular) this.btnAnular.Enabled = true;					
				}
				else if ((int)this.txtEstado.Value == 1) 
				{
					if (miAcceso.BDesbloquearCB) this.btnAbrir.Enabled = true;
					this.btnAbrir.Text = "Abrir Conciliación";				
				}						
				
				if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;				
				
				this.btnCancelar.Enabled = true;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaConciliaciones Busqueda = new frmBuscaConciliaciones())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridConciliaciones.ActiveRow.Cells["idConciliacionBancaria"].Value);
				}
			}
		}

		private void Guardar()
		{
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			if (bNuevo) this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, "Exec NumeroConciliacion");

			string sSQL = string.Format("Exec GuardaConciliacionBancaria {0}, {1}, '{2}', '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}",
				(int)this.txtIdConciliacionBancaria.Value, (int)this.cmbCuenta.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtNumero.Text.ToString(), (int)this.txtEstado.Value,
				Convert.ToDecimal(this.txtSaldoEnLibros.Value), Convert.ToDecimal(this.txtSaldoConciliado.Value), Convert.ToDecimal(this.txtSaldoNoConciliado.Value), Convert.ToDecimal(this.txtSaldoBancario.Value), Convert.ToDecimal(this.txtDiferencias.Value));
			this.txtIdConciliacionBancaria.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
		}


		private void btnImportar_Click(object sender, System.EventArgs e)
		{
			if (this.cmbCuenta.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCuenta.Focus();
			}
			else
			{
				this.openFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

				if(openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName.Length > 0) 
				{
					this.Guardar();

					this.EntradasSalidas();
					
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From EstadosDeCuentasBancarias Where idConciliacionBancaria = {0} And idCuenta = {1}", (int)this.txtIdConciliacionBancaria.Value, (int)this.cmbCuenta.Value));

					this.InsertaEstadoDeCuenta(this.openFileDialog1.FileName, (int)this.txtIdConciliacionBancaria.Value, (int)this.cmbCuenta.Value);

					string sSQLEntradas = string.Format("Exec CBConsultaEstadoDeCuenta {0}, {1}, 0", (int)this.txtIdConciliacionBancaria.Value, (int)this.cmbCuenta.Value);
					this.uGridECEntradas.DataSource = Funcion.dvProcedimiento (cdsSeteoF, sSQLEntradas);
					string sSQLSalidas = string.Format("Exec CBConsultaEstadoDeCuenta {0}, {1}, 1", (int)this.txtIdConciliacionBancaria.Value, (int)this.cmbCuenta.Value);
					this.uGridECSalidas.DataSource = Funcion.dvProcedimiento (cdsSeteoF, sSQLSalidas);
						
					this.btnComparar.Enabled = true;
				}				
			}
		}

		private void btnAbrir_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value == 0)
			{
				if (DialogResult.Yes == MessageBox.Show("Esta Seguro de Cerrar esta Conciliación", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					string sSQL = string.Format("Update ConciliacionBancaria Set Estado = 1 Where idConciliacionBancaria = {0}", (int)this.txtIdConciliacionBancaria.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
					this.lblEstado.Text = "CERRADO";

					this.btnAbrir.Text = "Abrir Conciliación";

					this.txtEstado.Value = 1;

					if (!miAcceso.BDesbloquearCB) this.btnAbrir.Enabled = false;	
		
					this.btnEditar.Enabled = false;
					this.btnAnular.Enabled = false;
				}
			}
			else if ((int)this.txtEstado.Value == 1)
			{
				if (DialogResult.Yes == MessageBox.Show("Esta Seguro de Abrir esta Conciliación", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					string sSQL = string.Format("Update ConciliacionBancaria Set Estado = 0 Where idConciliacionBancaria = {0}", (int)this.txtIdConciliacionBancaria.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
					this.lblEstado.Text = "ABIERTO";

					this.btnAbrir.Text = "Cerrar Conciliación";

					this.txtEstado.Value = 0;

					if (!miAcceso.BBloquearCB) this.btnAbrir.Enabled = false;

					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.Anular) this.btnAnular.Enabled = true;
				}
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show("Esta Seguro de Anular esta Conciliación", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
			{
				string sSQLAnula = string.Format("Exec AnulaConciliacionBancaria {0}", (int)this.txtIdConciliacionBancaria.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAnula, true);

				this.lblEstado.Text = "ANULADO";
			
				this.txtEstado.Value = 2;

				this.btnAbrir.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;				
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldCompra = new ParameterField ();
			ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
			paramFieldCompra.ParameterFieldName = "@idConciliacion";
			ValIdCompra.Value = int.Parse(this.txtIdConciliacionBancaria.Value.ToString());
			paramFieldCompra.CurrentValues.Add (ValIdCompra);
			paramFields.Add (paramFieldCompra);
				
			Reporte miRep = new Reporte("ResumenConciliacionBancaria.rpt", "");
			miRep.MdiParent = this.MdiParent;			
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
				
			ParameterFields pfsDetalle = new ParameterFields ();

			ParameterField pFIdConciliacion = new ParameterField ();
			ParameterDiscreteValue pdvIdConciliacion = new ParameterDiscreteValue ();
			pFIdConciliacion.ParameterFieldName = "@idConciliacion";
			pdvIdConciliacion.Value = int.Parse(this.txtIdConciliacionBancaria.Value.ToString());
			pFIdConciliacion.CurrentValues.Add (pdvIdConciliacion);
			pfsDetalle.Add (pFIdConciliacion);
			
			ParameterField pfIdCuenta = new ParameterField ();
			ParameterDiscreteValue pdvIdCuenta = new ParameterDiscreteValue ();
			pfIdCuenta.ParameterFieldName = "@idCuenta";
			pdvIdCuenta.Value = int.Parse(this.cmbCuenta.Value.ToString());
			pfIdCuenta.CurrentValues.Add (pdvIdCuenta);
			pfsDetalle.Add (pfIdCuenta);

			ParameterField pfDesde = new ParameterField ();
			ParameterDiscreteValue pdvDesde = new ParameterDiscreteValue ();
			pfDesde.ParameterFieldName = "@Desde";
			pdvDesde.Value = (DateTime)this.dtDesde.Value;
			pfDesde.CurrentValues.Add (pdvDesde);
			pfsDetalle.Add (pfDesde);
			
			ParameterField pfHasta = new ParameterField ();
			ParameterDiscreteValue pdvHasta = new ParameterDiscreteValue ();
			pfHasta.ParameterFieldName = "@Hasta";
			pdvHasta.Value = (DateTime)this.dtHasta.Value;
			pfHasta.CurrentValues.Add (pdvHasta);
			pfsDetalle.Add (pfHasta);

			Reporte Rep = new Reporte("ListaDetalleConciliacionBancaria.rpt", "");
			Rep.MdiParent = this.MdiParent;			
			Rep.crVista.ParameterFieldInfo = pfsDetalle;
			Rep.Show();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnAbrir.Enabled = false;
			this.btnComparar.Enabled = true;

			if (miAcceso.BModificarSaldoCB) this.txtSaldoBancario.Enabled = true;

			this.HabiltaGrids(true);

			bNuevo = false;
			bEdicion = true;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			this.Guardar();

			bNuevo = false;
			bEdicion = false;

			this.cmbCuenta.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			this.txtSaldoBancario.Enabled = false;

			this.btnGuardar.Enabled = false;
			this.btnImportar.Enabled = false;
			this.btnComparar.Enabled = false;
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			if (miAcceso.Imprimir)this.btnImprimir.Enabled = true; else this.btnImprimir.Enabled = false;
			if (miAcceso.Anular)this.btnAnular.Enabled = true;
			if (miAcceso.BBloquearCB) this.btnAbrir.Enabled = true;
			this.btnCancelar.Enabled = true;

			this.HabiltaGrids(true);					
		}

		private void txtSaldoBancario_Validated(object sender, System.EventArgs e)
		{
			Resumen();
		}

		private void txtSaldoBancario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.uGridEntrada.Focus();
		}

		private void uGridEntrada_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridEntrada_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idAsiento"].Value == System.DBNull.Value) return;

			if (e.Cell.Column.ToString() == "Confirmado")
			{
				this.uGridEntrada.UpdateData();

				string sSQL = string.Format("Exec GuardaDetalleConciliacionBancaria {0}, {1}, {2}", 
					(int)e.Cell.Row.Cells["idDetAsiento"].Value, (int)this.txtIdConciliacionBancaria.Value, (bool)e.Cell.Row.Cells["Confirmado"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

				DateTime dtDesde = (DateTime)this.dtDesde.Value;

				Resumen();
			}
		}

		private void uGridSalida_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idAsiento"].Value == System.DBNull.Value) return;

			if (e.Cell.Column.ToString() == "Confirmado")
			{
				this.uGridSalida.UpdateData();

				string sSQL = string.Format("Exec GuardaDetalleConciliacionBancaria {0}, {1}, {2}", 
					(int)e.Cell.Row.Cells["idDetAsiento"].Value, (int)this.txtIdConciliacionBancaria.Value, (bool)e.Cell.Row.Cells["Confirmado"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

				Resumen();
			}
		}

		private void uGridSalida_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmConciliacionAutomatica_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras Edita una Conciliación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
			}
		}

		private void cmbCuenta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}

