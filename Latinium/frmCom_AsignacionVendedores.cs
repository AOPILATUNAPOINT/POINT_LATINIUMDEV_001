using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_AsignacionVendedores.
	/// </summary>
	public class frmCom_AsignacionVendedores : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbLocales;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbVendedores;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRangos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstados;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.ComboBox cmbAño;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnImpresion;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnNuevoEmpleado;
		private System.Windows.Forms.Button btnFaltas;
		private System.Windows.Forms.DateTimePicker dtpPeriodo;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnRecalcular;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_AsignacionVendedores()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_AsignacionDeVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SumaPuerteo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango", -1, "cmbRangos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Estado", -1, "cmbEstados");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaMensual");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaMensualCredito");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaMensualAccesorios");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaMensualGEx");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaMensualContado");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diario");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensual");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiarioGE");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MensualGE");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiarioCarpa");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MensualCarpa");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMovimiento", 0);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_AsignacionVendedores));
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCalcula", 1);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_AsignacionDeVendedores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SumaPuerteo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaMensual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaMensualCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaMensualAccesorios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaMensualGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaMensualContado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mensual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiarioGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MensualGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiarioCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MensualCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rango");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			this.label4 = new System.Windows.Forms.Label();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbLocales = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbVendedores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbRangos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbAño = new System.Windows.Forms.ComboBox();
			this.btnVer = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnNuevoEmpleado = new System.Windows.Forms.Button();
			this.btnImpresion = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnFaltas = new System.Windows.Forms.Button();
			this.dtpPeriodo = new System.Windows.Forms.DateTimePicker();
			this.label7 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnRecalcular = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRangos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(11, 12);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(52, 20);
			this.label4.TabIndex = 676;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargos.DataSource = this.ultraDataSource3;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCargos.DisplayMember = "Cargo";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(22, 233);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(280, 136);
			this.cmbCargos.TabIndex = 682;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance1;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.Caption = "Local";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn6.Width = 180;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.Caption = "Grupo";
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 25;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn8.Width = 220;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 250;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.Caption = "Cargo";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 140;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.Caption = "Categoria";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Format = "";
			ultraGridColumn12.Header.Caption = "Estado";
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn12.Width = 75;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Format = "";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 70;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Format = "";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 70;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance2;
			ultraGridColumn15.Header.Caption = "N. Dias";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 50;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance3;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Meta Mensual";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 100;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance4;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Meta Crédito";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 100;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance5;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Meta Accesorios";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 120;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance6;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Meta G Ex";
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 100;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance7;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "Meta Contado";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 100;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance8;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "P Diario";
			ultraGridColumn21.Header.VisiblePosition = 22;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 85;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance9;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "P Mensual";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 85;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance10;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "Diario GEx";
			ultraGridColumn23.Header.VisiblePosition = 24;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 85;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance11;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "Mensual GEx";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 85;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance12;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.Caption = "Diario Carpa";
			ultraGridColumn25.Header.VisiblePosition = 25;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 85;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance13;
			ultraGridColumn26.Format = "#,##0";
			ultraGridColumn26.Header.Caption = "Meta Carpa";
			ultraGridColumn26.Header.VisiblePosition = 20;
			ultraGridColumn26.Width = 100;
			ultraGridColumn27.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance14.Image = ((object)(resources.GetObject("appearance14.Image")));
			appearance14.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance14.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn27.CellButtonAppearance = appearance14;
			ultraGridColumn27.Header.Caption = "...";
			ultraGridColumn27.Header.VisiblePosition = 3;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn27.Width = 30;
			ultraGridColumn28.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance15.Image = ((object)(resources.GetObject("appearance15.Image")));
			appearance15.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance15.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn28.CellButtonAppearance = appearance15;
			ultraGridColumn28.Header.Caption = "...";
			ultraGridColumn28.Header.VisiblePosition = 5;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn28.Width = 30;
			ultraGridBand2.Columns.AddRange(new object[] {
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
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridDetalle.DisplayLayout.GroupByBox.Hidden = true;
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance18;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 8F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance19;
			this.uGridDetalle.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance20.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance20.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.uGridDetalle.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(14, 107);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(1529, 330);
			this.uGridDetalle.TabIndex = 681;
			this.uGridDetalle.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_CellChange);
			this.uGridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalle_AfterRowInsert);
			this.uGridDetalle.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalle_BeforeRowsDeleted);
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_ClickCellButton);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn7.DefaultValue = false;
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn10.DefaultValue = 0;
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn11.DefaultValue = 0;
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn13.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn14.DataType = typeof(System.DateTime);
			ultraDataColumn14.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn15.DefaultValue = 0;
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn17.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn18.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn19.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn20.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn21.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn22.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn23.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn25.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn26.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn27.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// ultraDataSource4
			// 
			ultraDataColumn30.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn30,
																																 ultraDataColumn31});
			// 
			// cmbLocales
			// 
			this.cmbLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLocales.DataSource = this.ultraDataSource4;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 160;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.cmbLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbLocales.DisplayMember = "Nombre";
			this.cmbLocales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbLocales.Location = new System.Drawing.Point(582, 233);
			this.cmbLocales.Name = "cmbLocales";
			this.cmbLocales.Size = new System.Drawing.Size(280, 136);
			this.cmbLocales.TabIndex = 684;
			this.cmbLocales.ValueMember = "Bodega";
			this.cmbLocales.Visible = false;
			// 
			// cmbVendedores
			// 
			this.cmbVendedores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedores.DataSource = this.ultraDataSource2;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 57;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Width = 100;
			ultraGridColumn33.Header.VisiblePosition = 2;
			ultraGridColumn33.Width = 180;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33});
			this.cmbVendedores.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbVendedores.DisplayMember = "Nombre";
			this.cmbVendedores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbVendedores.Location = new System.Drawing.Point(302, 233);
			this.cmbVendedores.Name = "cmbVendedores";
			this.cmbVendedores.Size = new System.Drawing.Size(280, 136);
			this.cmbVendedores.TabIndex = 683;
			this.cmbVendedores.ValueMember = "idPersonal";
			this.cmbVendedores.Visible = false;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(11, 52);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 20);
			this.lblBodega.TabIndex = 686;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource5;
			ultraGridColumn34.Header.VisiblePosition = 2;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Width = 180;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(90, 50);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(280, 25);
			this.cmbBodega.TabIndex = 685;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn32.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34});
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
			// cmbRangos
			// 
			this.cmbRangos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRangos.DataSource = this.ultraDataSource6;
			ultraGridColumn37.Header.VisiblePosition = 0;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 1;
			ultraGridColumn38.Width = 160;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn37,
																										 ultraGridColumn38});
			this.cmbRangos.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbRangos.DisplayMember = "Rango";
			this.cmbRangos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbRangos.Location = new System.Drawing.Point(862, 233);
			this.cmbRangos.Name = "cmbRangos";
			this.cmbRangos.Size = new System.Drawing.Size(280, 136);
			this.cmbRangos.TabIndex = 687;
			this.cmbRangos.ValueMember = "idCom_Rango";
			this.cmbRangos.Visible = false;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn35.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn35,
																																 ultraDataColumn36});
			// 
			// cmbEstados
			// 
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource7;
			ultraGridColumn39.Header.VisiblePosition = 0;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 1;
			ultraGridColumn40.Width = 200;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn39,
																										 ultraGridColumn40});
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbEstados.DisplayMember = "Estado";
			this.cmbEstados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbEstados.Location = new System.Drawing.Point(1142, 233);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(280, 136);
			this.cmbEstados.TabIndex = 688;
			this.cmbEstados.ValueMember = "idCom_Estado";
			this.cmbEstados.Visible = false;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn37.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn37,
																																 ultraDataColumn38});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 87);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1557, 10);
			this.groupBox1.TabIndex = 689;
			this.groupBox1.TabStop = false;
			// 
			// cmbAño
			// 
			this.cmbAño.Location = new System.Drawing.Point(64, 8);
			this.cmbAño.MaxDropDownItems = 6;
			this.cmbAño.Name = "cmbAño";
			this.cmbAño.Size = new System.Drawing.Size(96, 21);
			this.cmbAño.TabIndex = 724;
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(818, 49);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(100, 28);
			this.btnVer.TabIndex = 725;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(258, 10);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 24);
			this.lblContador.TabIndex = 726;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnNuevoEmpleado
			// 
			this.btnNuevoEmpleado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevoEmpleado.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevoEmpleado.CausesValidation = false;
			this.btnNuevoEmpleado.Enabled = false;
			this.btnNuevoEmpleado.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevoEmpleado.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevoEmpleado.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevoEmpleado.Image")));
			this.btnNuevoEmpleado.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevoEmpleado.Location = new System.Drawing.Point(1053, 9);
			this.btnNuevoEmpleado.Name = "btnNuevoEmpleado";
			this.btnNuevoEmpleado.Size = new System.Drawing.Size(157, 27);
			this.btnNuevoEmpleado.TabIndex = 729;
			this.btnNuevoEmpleado.Text = "Nuevo Empleado";
			this.btnNuevoEmpleado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevoEmpleado.Click += new System.EventHandler(this.btnNuevoEmpleado_Click);
			// 
			// btnImpresion
			// 
			this.btnImpresion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImpresion.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnImpresion.CausesValidation = false;
			this.btnImpresion.Enabled = false;
			this.btnImpresion.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImpresion.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImpresion.Image = ((System.Drawing.Image)(resources.GetObject("btnImpresion.Image")));
			this.btnImpresion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImpresion.Location = new System.Drawing.Point(1333, 9);
			this.btnImpresion.Name = "btnImpresion";
			this.btnImpresion.Size = new System.Drawing.Size(101, 27);
			this.btnImpresion.TabIndex = 730;
			this.btnImpresion.Text = "Imprimir";
			this.btnImpresion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImpresion.Click += new System.EventHandler(this.btnImpresion_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1445, 9);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(101, 27);
			this.btnExportar.TabIndex = 731;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnFaltas
			// 
			this.btnFaltas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnFaltas.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnFaltas.CausesValidation = false;
			this.btnFaltas.Enabled = false;
			this.btnFaltas.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnFaltas.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnFaltas.Image = ((System.Drawing.Image)(resources.GetObject("btnFaltas.Image")));
			this.btnFaltas.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFaltas.Location = new System.Drawing.Point(1221, 9);
			this.btnFaltas.Name = "btnFaltas";
			this.btnFaltas.Size = new System.Drawing.Size(101, 27);
			this.btnFaltas.TabIndex = 732;
			this.btnFaltas.Text = "Faltas";
			this.btnFaltas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnFaltas.Click += new System.EventHandler(this.btnFaltas_Click);
			// 
			// dtpPeriodo
			// 
			this.dtpPeriodo.CustomFormat = "MM/yyyy";
			this.dtpPeriodo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpPeriodo.Location = new System.Drawing.Point(90, 10);
			this.dtpPeriodo.MinDate = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			this.dtpPeriodo.Name = "dtpPeriodo";
			this.dtpPeriodo.ShowUpDown = true;
			this.dtpPeriodo.Size = new System.Drawing.Size(134, 24);
			this.dtpPeriodo.TabIndex = 733;
			this.dtpPeriodo.Value = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			this.dtpPeriodo.ValueChanged += new System.EventHandler(this.dtpPeriodo_ValueChanged);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(403, 52);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(65, 20);
			this.label7.TabIndex = 735;
			this.label7.Text = "Vendedor";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(504, 50);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(291, 24);
			this.txtBuscar.TabIndex = 734;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// btnRecalcular
			// 
			this.btnRecalcular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnRecalcular.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnRecalcular.CausesValidation = false;
			this.btnRecalcular.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnRecalcular.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnRecalcular.Image = ((System.Drawing.Image)(resources.GetObject("btnRecalcular.Image")));
			this.btnRecalcular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRecalcular.Location = new System.Drawing.Point(1056, 48);
			this.btnRecalcular.Name = "btnRecalcular";
			this.btnRecalcular.Size = new System.Drawing.Size(152, 27);
			this.btnRecalcular.TabIndex = 736;
			this.btnRecalcular.Text = "Recalcular";
			this.btnRecalcular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRecalcular.Click += new System.EventHandler(this.button1_Click);
			// 
			// frmCom_AsignacionVendedores
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(7, 17);
			this.ClientSize = new System.Drawing.Size(1556, 444);
			this.Controls.Add(this.btnRecalcular);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnFaltas);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnImpresion);
			this.Controls.Add(this.btnNuevoEmpleado);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtpPeriodo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbRangos);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbCargos);
			this.Controls.Add(this.cmbLocales);
			this.Controls.Add(this.cmbVendedores);
			this.Controls.Add(this.uGridDetalle);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmCom_AsignacionVendedores";
			this.Text = "Asignación Mensual De Vendedores";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCom_AsignacionVendedores_KeyDown);
			this.Load += new System.EventHandler(this.frmCom_AsignacionVendedores_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRangos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private Acceso miAcceso;
		bool bValidaRango = false;
		bool bValidaMovPer = false;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCom_AsignacionVendedores_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "052105");
			
			if (!Funcion.Permiso("869", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Asignacion de vendedores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			
			this.dtpPeriodo.Value = DateTime.Parse("01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 7, 0");			
			this.cmbLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 7, 0");
			this.cmbCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 59, -1, 0");			
			this.cmbRangos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 60, 0");
			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 61, 0");
			
			if (miAcceso.BExportar) this.btnExportar.Enabled = true;
			if (miAcceso.Imprimir) this.btnImpresion.Enabled = true;
			if (miAcceso.Nuevo) this.btnNuevoEmpleado.Enabled = true;
			if (miAcceso.ReImprimir) this.btnFaltas.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle);
			if (miAcceso.Editar) FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);		
			if (miAcceso.CargoPersonal)
			{
				this.uGridDetalle.DisplayLayout.Bands[0].Columns["idCargo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}
			else
			{			
				this.uGridDetalle.DisplayLayout.Bands[0].Columns["idCargo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			}
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDetalle);
		}

		private void Actualiza(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int iBodega = 0;
			int idCargo = 0;
			int idPersonal = 0;
			int idCom_Rango = 0;
			int idCom_Estado = 0;

			if (e.Cell.Row.Cells["Bodega"].Value != System.DBNull.Value) iBodega = (int)e.Cell.Row.Cells["Bodega"].Value;
			if (e.Cell.Row.Cells["idCargo"].Value != System.DBNull.Value) idCargo = (int)e.Cell.Row.Cells["idCargo"].Value;
			if (e.Cell.Row.Cells["idPersonal"].Value != System.DBNull.Value) idPersonal = (int)e.Cell.Row.Cells["idPersonal"].Value;
			if (e.Cell.Row.Cells["idCom_Rango"].Value != System.DBNull.Value) idCom_Rango = (int)e.Cell.Row.Cells["idCom_Rango"].Value;
			if (e.Cell.Row.Cells["idCom_Estado"].Value != System.DBNull.Value) idCom_Estado = (int)e.Cell.Row.Cells["idCom_Estado"].Value;

			string sSQL = string.Format("Exec Com_GuardaAsignacionVendedores {0}, '{1}', {2}, {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, '{18}', {19}, {20}, {21}, {22}, {23}",
				(int)e.Cell.Row.Cells["idCom_AsignacionDeVendedores"].Value, this.dtpPeriodo.Text.ToString(), iBodega, idCargo, idPersonal, idCom_Rango, //5
				Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd"), //7
				(int)e.Cell.Row.Cells["NumeroDias"].Value, Convert.ToDecimal(e.Cell.Row.Cells["Diario"].Value), Convert.ToDecimal(e.Cell.Row.Cells["Mensual"].Value), //11
				Convert.ToDecimal(e.Cell.Row.Cells["DiarioGE"].Value), Convert.ToDecimal(e.Cell.Row.Cells["MensualGE"].Value),
				Convert.ToDecimal(e.Cell.Row.Cells["DiarioCarpa"].Value), Convert.ToDecimal(e.Cell.Row.Cells["MensualCarpa"].Value),
				idCom_Estado, 0, false, "20010101", (bool)e.Cell.Row.Cells["SumaPuerteo"].Value, 
				Convert.ToDecimal(e.Cell.Row.Cells["MetaMensual"].Value), Convert.ToDecimal(e.Cell.Row.Cells["MetaMensualContado"].Value),
				Convert.ToDecimal(e.Cell.Row.Cells["MetaMensualCredito"].Value), Convert.ToDecimal(e.Cell.Row.Cells["MetaMensualGEx"].Value));
			e.Cell.Row.Cells["idCom_AsignacionDeVendedores"].Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			sSQL = string.Format("Exec Com_ActualizaDatosNominaAV {0}, {1}, {2}, {3}, {4}", 
				iBodega, idCargo, idCom_Estado, idCom_Rango, idPersonal);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}
		
		private void Metas(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCom_Estado"].Value != System.DBNull.Value)
			{
				if (e.Cell.Row.Cells["idCargo"].Value != System.DBNull.Value)
				{
					e.Cell.Row.Cells["NumeroDias"].Value = FuncionesProcedimientos.CalculaDias((DateTime)e.Cell.Row.Cells["Desde"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, cdsSeteoF);

					int iDiasTrabaja = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(DAY, dbo.NominaRetornaFechaEntradaPersonal({0}, 0, 0), '{1}')", 
						(int)e.Cell.Row.Cells["idPersonal"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd")));

					e.Cell.Row.Cells["MetaMensual"].Value = FuncionesProcedimientos.CalculaMeta((int)e.Cell.Row.Cells["Bodega"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Month, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Year, (int)e.Cell.Row.Cells["idCargo"].Value, (int)e.Cell.Row.Cells["NumeroDias"].Value, 1, cdsSeteoF);
					e.Cell.Row.Cells["MetaMensualContado"].Value = 0.00m;//FuncionesProcedimientos.CalculaMeta((int)e.Cell.Row.Cells["Bodega"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Month, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Year, (int)e.Cell.Row.Cells["idCargo"].Value, (int)e.Cell.Row.Cells["NumeroDias"].Value, 2, cdsSeteoF);
					e.Cell.Row.Cells["MetaMensualCredito"].Value = FuncionesProcedimientos.CalculaMeta((int)e.Cell.Row.Cells["Bodega"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Month, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Year, (int)e.Cell.Row.Cells["idCargo"].Value, (int)e.Cell.Row.Cells["NumeroDias"].Value, 3, cdsSeteoF);
					e.Cell.Row.Cells["MetaMensualGEx"].Value = FuncionesProcedimientos.CalculaMeta((int)e.Cell.Row.Cells["Bodega"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Month, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Year, (int)e.Cell.Row.Cells["idCargo"].Value, (int)e.Cell.Row.Cells["NumeroDias"].Value, 4, cdsSeteoF);
					e.Cell.Row.Cells["MensualCarpa"].Value = FuncionesProcedimientos.CalculaMeta((int)e.Cell.Row.Cells["Bodega"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Month, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Year, (int)e.Cell.Row.Cells["idCargo"].Value, (int)e.Cell.Row.Cells["NumeroDias"].Value, 5, cdsSeteoF);
					e.Cell.Row.Cells["MetaMensualAccesorios"].Value = FuncionesProcedimientos.CalculaMeta((int)e.Cell.Row.Cells["Bodega"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Month, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).Year, (int)e.Cell.Row.Cells["idCargo"].Value, (int)e.Cell.Row.Cells["NumeroDias"].Value, 6, cdsSeteoF);
				}
			}
		}
		
		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{	
			if (bValidaRango) return;

			if (bValidaMovPer) return;
			
			if (e.Cell.Column.ToString() == "Bodega")
			{
				Metas(e);
				this.Actualiza(e);	
			}

			if (e.Cell.Column.ToString() == "idCargo")
			{
				Metas(e);
				this.Actualiza(e);	
			}

			if (e.Cell.Column.ToString() == "idCom_Estado")
			{
				this.Actualiza(e);	
			}

			if (e.Cell.Column.ToString() == "idPersonal")
			{
				this.Actualiza(e);	
			}

			if (e.Cell.Column.ToString() == "idCom_Rango")
			{
				this.Actualiza(e);					
			}

			if (e.Cell.Column.ToString() == "Desde")
			{
				Metas(e);
				this.Actualiza(e);											
			}

			if (e.Cell.Column.ToString() == "Hasta")
			{
				Metas(e);
				this.Actualiza(e);
			}

			if (e.Cell.Column.ToString() == "NumeroDias")
			{
				this.Actualiza(e);	
			}
			
			bValidaRango = false;
		}
				
		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCom_AsignacionDeVendedores"].Value = 0;	
			e.Row.Cells["idCom_Estado"].Value = 1;
		}

		private void Consulta()
		{
			if (this.dtpPeriodo.Text.ToString().Length == 0) return;			

			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec Com_ConsultaAsginacionVendedores '{0}', {1}, '{2}'", 
				this.dtpPeriodo.Text.ToString(), iBodega, this.txtBuscar.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDetalle);
			
			this.lblContador.Text = this.uGridDetalle.Rows.Count + " REGISTROS ENCONTRADOS";

//			this.uGridDetalle.DisplayLayout.Bands[0].SortedColumns.Add("Bodega", false, true);	
//			this.uGridDetalle.Rows.ExpandAll(true);

			if (FuncionesProcedimientos.ValidaPeriodoCerrado(this.dtpPeriodo.Text.ToString(), cdsSeteoF)) 
			{
				FuncionesProcedimientos.EstadoGrids(false, this.uGridDetalle); 
				this.btnNuevoEmpleado.Enabled = false;
				this.btnFaltas.Enabled = false;
			}
			else 
			{
				if (miAcceso.Editar) FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
				if (miAcceso.Nuevo) this.btnNuevoEmpleado.Enabled = true;
				if (miAcceso.ReImprimir) this.btnFaltas.Enabled = true;
			}

			if (miAcceso.Editar) FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}

			Consulta();
		}

		private void frmCom_AsignacionVendedores_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
			if (e.KeyCode == Keys.F5) this.ListaVendedores();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void uGridDetalle_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg = false;
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta Seguro de Borrar este registro?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				for (int i=0; i<e.Rows.Length; i++)
				{
					string sSQL = string.Format("Delete From Com_AsignacionDeVendedores Where idCom_AsignacionDeVendedores = {0}", (int)e.Rows[i].Cells["idCom_AsignacionDeVendedores"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				}
			}
			else e.Cancel = true;
		}

		private void btnImpresion_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}

			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldPeriodo = new ParameterField ();
			ParameterDiscreteValue ValIdPeriodo = new ParameterDiscreteValue ();
			paramFieldPeriodo.ParameterFieldName = "@Periodo";
			ValIdPeriodo.Value = this.dtpPeriodo.Text.ToString();
			paramFieldPeriodo.CurrentValues.Add (ValIdPeriodo);
			paramFields.Add (paramFieldPeriodo);

			ParameterField paramFieldBodega = new ParameterField ();
			ParameterDiscreteValue ValIdBodega = new ParameterDiscreteValue ();
			paramFieldBodega.ParameterFieldName = "@Bodega";
			ValIdBodega.Value = iBodega;
			paramFieldBodega.CurrentValues.Add (ValIdBodega);
			paramFields.Add (paramFieldBodega);
			
			ParameterField paramFieldUsuario = new ParameterField ();
			ParameterDiscreteValue ValIdUsuario = new ParameterDiscreteValue ();
			paramFieldUsuario.ParameterFieldName = "@Usuario";
			ValIdUsuario.Value = MenuLatinium.stUsuario;
			paramFieldUsuario.CurrentValues.Add (ValIdUsuario);
			paramFields.Add (paramFieldUsuario);

			Reporte Reporte = new Reporte("Com_AsignacionVendedores.rpt", "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}

			FuncionesProcedimientos.ExportaExcel(this.uGridDetalle);
		}	

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			this.Consulta();

			this.uGridDetalle.DisplayLayout.Bands[0].Columns["SumaPuerteo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;

			if (this.cmbBodega.ActiveRow != null)
				if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Puerteo From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true))
				{
					this.uGridDetalle.DisplayLayout.Bands[0].Columns["SumaPuerteo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				}
		}

		private void btnNuevoEmpleado_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}
			
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
				return;
			}

			DateTime dtDesde = DateTime.Parse("01/" + this.dtpPeriodo.Text.ToString());
			DateTime dtHasta = DateTime.Parse(DateTime.DaysInMonth(Convert.ToInt32(this.dtpPeriodo.Text.Substring(3, 4)), Convert.ToInt32(this.dtpPeriodo.Text.Substring(0, 2))) + "/" + this.dtpPeriodo.Text.ToString());

			int iBodega = 0;
			
			using (frmCom_CambioVendedor CV = new frmCom_CambioVendedor(0, 0, true, this.dtpPeriodo.Text.ToString(), 0, 0, 0, dtDesde, dtHasta, dtHasta, (int)this.cmbBodega.Value,0))
			{				
				if (DialogResult.OK == CV.ShowDialog())
				{
					this.btnVer_Click(sender, e);
				}
			}
		}

		private void btnFaltas_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}
			
			if (this.uGridDetalle.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un vendedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			try
			{
				DateTime dtUltimoDiaMes = DateTime.Parse(FuncionesProcedimientos.iUltimoDiaDelMes(int.Parse(this.dtpPeriodo.Text.ToString().Substring(0, 2)), int.Parse(this.dtpPeriodo.Text.ToString().Substring(3, 4))) + "/" + this.dtpPeriodo.Text.ToString());

				using (frmNominaPermisosFaltas NPF = new frmNominaPermisosFaltas((int)this.uGridDetalle.ActiveRow.Cells["idPersonal"].Value, (int)this.uGridDetalle.ActiveRow.Cells["idCom_AsignacionDeVendedores"].Value, 
								 (int)this.uGridDetalle.ActiveRow.Cells["Bodega"].Value, this.dtpPeriodo.Text.ToString(), 1, (DateTime)this.uGridDetalle.ActiveRow.Cells["Desde"].Value, (DateTime)this.uGridDetalle.ActiveRow.Cells["Hasta"].Value, dtUltimoDiaMes, 0, 0, 0))
				{				
					if (DialogResult.OK == NPF.ShowDialog())
					{
						
					}
				}
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void ListaVendedores()
		{
			string sMes = this.dtpPeriodo.Text.ToString().Substring(0, 2);
			string sAño = this.dtpPeriodo.Text.ToString().Substring(3, 4);
			
			DateTime dtFecha = DateTime.Parse("01/" + sMes + "/" + sAño);

			this.cmbVendedores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Com_ListaVendedores '{0}'", dtFecha.ToString("yyyyMMdd")));
		}

		private void dtpPeriodo_ValueChanged(object sender, System.EventArgs e)
		{
			this.Consulta();

			//this.ListaVendedores();
		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridDetalle_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnMovimiento")
			{	
				if ((DateTime)e.Cell.Row.Cells["Hasta"].Value < DateTime.Today)
				{
					MessageBox.Show("No puede generar movimientos de periodos pasados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)e.Cell.Row.Cells["idCom_Estado"].Value == 2)
				{
					MessageBox.Show("No puede generar movimientos de personal inactivo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				#region Valida Movimientos
				string sSQLValida = string.Format("Exec Com_ValidaMovimientos {0}, {1}, '{2}'", 
					(int)e.Cell.Row.Cells["idCom_AsignacionDeVendedores"].Value, (int)e.Cell.Row.Cells["idPersonal"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd"));
				if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValida) > 0)
				{
					bValidaMovPer = true;
					e.Cell.Value = e.Cell.OriginalValue;
					MessageBox.Show(string.Format("Ya existen movimientos registrados en fechas posteriores a la seleccionada", Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					bValidaMovPer = false;
					return;
				}
				#endregion Valida Movimientos
			
				#region Valida Permisos
				sSQLValida = string.Format("Exec Com_ValidaMovimientoPermisos {0}, '{1}', 0", 
					(int)e.Cell.Row.Cells["idPersonal"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd"));

				if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValida) > 0)
				{
					bValidaMovPer = true;
					e.Cell.Value = e.Cell.OriginalValue;
					MessageBox.Show(string.Format("Existen Permisos registrados en fechas posteriores a la seleccionada", Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					bValidaMovPer = false;
					return;
				}
				#endregion Valida Permisos

				#region Valida Ventas
				sSQLValida = string.Format("Exec Com_ValidaMovimientosVentas {0}, {1}, '{2}', '{3}', 0", 
					(int)e.Cell.Row.Cells["idPersonal"].Value, (int)e.Cell.Row.Cells["Bodega"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd"));

				if (Funcion.iEscalarSQL(cdsSeteoF, sSQLValida) > 0)
				{
					bValidaMovPer = true;
					e.Cell.Value = e.Cell.OriginalValue;
					MessageBox.Show(string.Format("Ya existen ventas registradas en fechas posteriores a la seleccionada", Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					bValidaMovPer = false;
					return;
				}
				#endregion Valida Ventas
				
				using (frmCom_ValidaCambioFechaAV VC = new frmCom_ValidaCambioFechaAV())
				{				
					if (DialogResult.OK == VC.ShowDialog())
					{	
						DateTime dtUltimoDiaMes = DateTime.Parse(FuncionesProcedimientos.iUltimoDiaDelMes(int.Parse(this.dtpPeriodo.Text.ToString().Substring(0, 2)), int.Parse(this.dtpPeriodo.Text.ToString().Substring(3, 4))) + "/" + this.dtpPeriodo.Text.ToString());

						#region Movimientos
						if (VC.rdbNuevoMovimiento.Checked) 
						{
							using (frmCom_CambioVendedor CV = new frmCom_CambioVendedor(
											 (int)e.Cell.Row.Cells["idPersonal"].Value, (int)e.Cell.Row.Cells["idCom_AsignacionDeVendedores"].Value,
											 false, this.dtpPeriodo.Text.ToString(), (int)e.Cell.Row.Cells["idCom_Estado"].Value, (int)e.Cell.Row.Cells["idCargo"].Value, (int)e.Cell.Row.Cells["idCom_Rango"].Value, 
											 (DateTime)e.Cell.Row.Cells["Desde"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, dtUltimoDiaMes, (int)e.Cell.Row.Cells["Bodega"].Value,(int)e.Cell.Row.Cells["idUsuario"].Value))
							{				
								if (DialogResult.OK == CV.ShowDialog())
								{
									if ((int)CV.txtIdCom_AsignacionDeVendedoresN.Value != (int)e.Cell.Row.Cells["idCom_AsignacionDeVendedores"].Value)
									{
										//Metas(e);
										//this.Actualiza(e);
										this.btnVer_Click(sender, e);
									}
								}					
							}	
						}
						#endregion Movimientos

						#region Permisos
						if (VC.rdbPFV.Checked) 
						{							
							int idCom_Rango = 0;
							int idCargo = 0;
							if (e.Cell.Row.Cells["idCom_Rango"].Value != System.DBNull.Value) idCom_Rango = (int)e.Cell.Row.Cells["idCom_Rango"].Value;
							if (e.Cell.Row.Cells["idCargo"].Value != System.DBNull.Value) idCargo = (int)e.Cell.Row.Cells["idCargo"].Value;

							using (frmNominaPermisosFaltas NPF = new frmNominaPermisosFaltas((int)e.Cell.Row.Cells["idPersonal"].Value, (int)e.Cell.Row.Cells["idCom_AsignacionDeVendedores"].Value, 
											 (int)e.Cell.Row.Cells["Bodega"].Value, this.dtpPeriodo.Text.ToString(), 0, (DateTime)e.Cell.Row.Cells["Desde"].Value, (DateTime)e.Cell.Row.Cells["Hasta"].Value, dtUltimoDiaMes, 
											 (int)e.Cell.Row.Cells["idCargo"].Value, idCom_Rango, (int)e.Cell.Row.Cells["idCom_Estado"].Value))
							{				
								if (DialogResult.OK == NPF.ShowDialog())
								{									
//									Metas(e);										
//									this.Actualiza(e);
									
									this.btnVer_Click(sender, e);
								}								
							}
						}						
						#endregion Permisos
					}
					else 
					{						
						bValidaMovPer = true;
						e.Cell.Value = e.Cell.OriginalValue;
						bValidaMovPer = false;
					}
				}				
			}

			if (e.Cell.Column.ToString() == "btnCalcula")
			{
				this.Metas(e);
				this.Actualiza(e);
			}
		}

		private void uGridDetalle_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "SumaPuerteo")
			{
				this.uGridDetalle.UpdateData();

				this.Actualiza(e);

				this.uGridDetalle.DisplayLayout.Bands[0].SortedColumns.Add("Bodega", false, false);
								
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				{
					if ((int) dr.Cells["idCargo"].Value == 10) 
					{
						dr.Cells["Desde"].Value = Convert.ToDateTime(dr.Cells["Desde"].Value).AddDays(1);
						dr.Cells["Desde"].Value = Convert.ToDateTime(dr.Cells["Desde"].Value).AddDays(-1);
						continue;
					}
				}

				this.uGridDetalle.DisplayLayout.Bands[0].SortedColumns.Add("Bodega", false, true);	
				this.uGridDetalle.Rows.ExpandAll(true);
			}
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if (this.uGridDetalle.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Recalcular",
					"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			// Mostrar el modal de "Procesando..."
			MenuLatinium.MostrarModalProcesando();

			try
			{
				Infragistics.Win.UltraWinGrid.UltraGridRow row;

				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in this.uGridDetalle.Rows)
				{
					row = r;

					this.Metas_Row(row);
					this.Actualiza_Row(row);
				}
			}
			finally
			{
				// Asegura que el modal SIEMPRE se cierre aunque haya un error
				MenuLatinium.OcultarModalProcesando();
			}

			// Mensaje de confirmación
			MessageBox.Show("Recalculo completado correctamente.",
				"Point Technology",
				MessageBoxButtons.OK,
				MessageBoxIcon.Information);
		}

		private void Metas_Row(Infragistics.Win.UltraWinGrid.UltraGridRow row)
		{
			if (row.Cells["idCom_Estado"].Value == DBNull.Value)
				return;

			if (row.Cells["idCargo"].Value == DBNull.Value)
				return;

			// Calcular días
			row.Cells["NumeroDias"].Value =
				FuncionesProcedimientos.CalculaDias(
				(DateTime)row.Cells["Desde"].Value,
				(DateTime)row.Cells["Hasta"].Value,
				cdsSeteoF);

			// Calcular días trabajados
			int iDiasTrabaja =
				Funcion.iEscalarSQL(
				cdsSeteoF,
				string.Format(
				"Select DATEDIFF(DAY, dbo.NominaRetornaFechaEntradaPersonal({0}, 0, 0), '{1}')",
				Convert.ToInt32(row.Cells["idPersonal"].Value),
				Convert.ToDateTime(row.Cells["Hasta"].Value).ToString("yyyyMMdd")
				)
				);

			// Datos base
			int bodega = Convert.ToInt32(row.Cells["Bodega"].Value);
			DateTime dtDesde = Convert.ToDateTime(row.Cells["Desde"].Value);
			int mes = dtDesde.Month;
			int anio = dtDesde.Year;
			int cargo = Convert.ToInt32(row.Cells["idCargo"].Value);
			int dias = Convert.ToInt32(row.Cells["NumeroDias"].Value);

			// Cálculos de metas
			row.Cells["MetaMensual"].Value =
				FuncionesProcedimientos.CalculaMeta(bodega, mes, anio, cargo, dias, 1, cdsSeteoF);

			row.Cells["MetaMensualContado"].Value = 0m;  // así lo tienes tú

			row.Cells["MetaMensualCredito"].Value =
				FuncionesProcedimientos.CalculaMeta(bodega, mes, anio, cargo, dias, 3, cdsSeteoF);

			row.Cells["MetaMensualGEx"].Value =
				FuncionesProcedimientos.CalculaMeta(bodega, mes, anio, cargo, dias, 4, cdsSeteoF);

			row.Cells["MensualCarpa"].Value =
				FuncionesProcedimientos.CalculaMeta(bodega, mes, anio, cargo, dias, 5, cdsSeteoF);

			row.Cells["MetaMensualAccesorios"].Value =
				FuncionesProcedimientos.CalculaMeta(bodega, mes, anio, cargo, dias, 6, cdsSeteoF);
		}

		private void Actualiza_Row(Infragistics.Win.UltraWinGrid.UltraGridRow row)
		{
			int iBodega = row.Cells["Bodega"].Value != DBNull.Value ?
				Convert.ToInt32(row.Cells["Bodega"].Value) : 0;

			int idCargo = row.Cells["idCargo"].Value != DBNull.Value ?
				Convert.ToInt32(row.Cells["idCargo"].Value) : 0;

			int idPersonal = row.Cells["idPersonal"].Value != DBNull.Value ?
				Convert.ToInt32(row.Cells["idPersonal"].Value) : 0;

			int idCom_Rango = row.Cells["idCom_Rango"].Value != DBNull.Value ?
				Convert.ToInt32(row.Cells["idCom_Rango"].Value) : 0;

			int idCom_Estado = row.Cells["idCom_Estado"].Value != DBNull.Value ?
				Convert.ToInt32(row.Cells["idCom_Estado"].Value) : 0;

			string sSQL =
				string.Format(
				"Exec Com_GuardaAsignacionVendedores {0}, '{1}', {2}, {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, '{18}', {19}, {20}, {21}, {22}, {23}",
				Convert.ToInt32(row.Cells["idCom_AsignacionDeVendedores"].Value),
				this.dtpPeriodo.Text,
				iBodega,
				idCargo,
				idPersonal,
				idCom_Rango,
				Convert.ToDateTime(row.Cells["Desde"].Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(row.Cells["Hasta"].Value).ToString("yyyyMMdd"),
				Convert.ToInt32(row.Cells["NumeroDias"].Value),
				Convert.ToDecimal(row.Cells["Diario"].Value),
				Convert.ToDecimal(row.Cells["Mensual"].Value),
				Convert.ToDecimal(row.Cells["DiarioGE"].Value),
				Convert.ToDecimal(row.Cells["MensualGE"].Value),
				Convert.ToDecimal(row.Cells["DiarioCarpa"].Value),
				Convert.ToDecimal(row.Cells["MensualCarpa"].Value),
				idCom_Estado,
				0,
				false,
				"20010101",
				Convert.ToBoolean(row.Cells["SumaPuerteo"].Value),
				Convert.ToDecimal(row.Cells["MetaMensual"].Value),
				Convert.ToDecimal(row.Cells["MetaMensualContado"].Value),
				Convert.ToDecimal(row.Cells["MetaMensualCredito"].Value),
				Convert.ToDecimal(row.Cells["MetaMensualGEx"].Value)
				);

			// Guardar ID retornado
			row.Cells["idCom_AsignacionDeVendedores"].Value =
				Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			// Actualizar nomina
			sSQL = string.Format(
				"Exec Com_ActualizaDatosNominaAV {0}, {1}, {2}, {3}, {4}",
				iBodega, idCargo, idCom_Estado, idCom_Rango, idPersonal);

			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}
	}
}
