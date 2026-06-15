using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmCre_EnvioCedulasVendedores.
	/// </summary>
	public class frmCre_EnvioCedulasVendedores : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label lblVendedor;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProducto;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnEnviar;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidoPaterno;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidoMaterno;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSegundoNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPrimerNombre;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private System.Windows.Forms.Timer timer1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCedulas;
		private System.Windows.Forms.Timer timer2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAutomatico;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.GroupBox gbCarpa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optAplica;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCarpa;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaNacimiento;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaVivienda;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVivienda;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaTrabajo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTrabajo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipoFuenteIngresos;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmpresaNegocio;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtdIngreso;
		private System.Windows.Forms.Label label20;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbEntidadFinanciera;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesCopiarUltimoMes;
		private System.Windows.Forms.MenuItem mnuAccionesAgregarNuevoLocal;
		private System.Windows.Forms.MenuItem mnuAccionesActualizar;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.ComponentModel.IContainer components;

		public frmCre_EnvioCedulasVendedores()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_EnvioCedulasVendedores));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioRespuesta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Respuesta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iEstado");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioRespuesta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Respuesta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iEstado");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoVivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActEconomica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActEconomica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.label4 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.txtProducto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnEnviar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.txtApellidoPaterno = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtApellidoMaterno = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtSegundoNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtPrimerNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.uGridCedulas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.timer2 = new System.Windows.Forms.Timer(this.components);
			this.chkAutomatico = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.gbCarpa = new System.Windows.Forms.GroupBox();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.optAplica = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.chkCarpa = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.dtFechaNacimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label14 = new System.Windows.Forms.Label();
			this.dtFechaVivienda = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label15 = new System.Windows.Forms.Label();
			this.cmbVivienda = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label16 = new System.Windows.Forms.Label();
			this.dtFechaTrabajo = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label17 = new System.Windows.Forms.Label();
			this.cmbTrabajo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optTipoFuenteIngresos = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label18 = new System.Windows.Forms.Label();
			this.txtEmpresaNegocio = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.txtdIngreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.cmbEntidadFinanciera = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnComandos = new System.Windows.Forms.Button();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesCopiarUltimoMes = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAgregarNuevoLocal = new System.Windows.Forms.MenuItem();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.mnuAccionesActualizar = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProducto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoPaterno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoMaterno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSegundoNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrimerNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCedulas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.gbCarpa.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optAplica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaTrabajo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTrabajo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoFuenteIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresaNegocio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 224);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 16);
			this.label4.TabIndex = 308;
			this.label4.Text = "Notas";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(768, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 307;
			this.label2.Text = "Número";
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(520, 8);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(38, 16);
			this.lblIdentificacion.TabIndex = 305;
			this.lblIdentificacion.Text = "Cédula";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 304;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(256, 168);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 16);
			this.lblVendedor.TabIndex = 303;
			this.lblVendedor.Text = "Vendedor";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(256, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 302;
			this.label6.Text = "Fecha";
			// 
			// txtNumero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(872, 7);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(152, 22);
			this.txtNumero.TabIndex = 1;
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance2;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(56, 200);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(528, 66);
			this.txtNotas.TabIndex = 11;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			this.txtNotas.ValueChanged += new System.EventHandler(this.txtNotas_ValueChanged);
			// 
			// txtRuc
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance3;
			this.txtRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Location = new System.Drawing.Point(616, 8);
			this.txtRuc.MaxLength = 10;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(128, 22);
			this.txtRuc.TabIndex = 2;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(56, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(168, 21);
			this.cmbBodega.TabIndex = 0;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// dtFecha
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance4;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFecha.Location = new System.Drawing.Point(360, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(144, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 3;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// cmbVendedor
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance5;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 335;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(360, 169);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(224, 21);
			this.cmbVendedor.TabIndex = 9;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVendedor_KeyDown);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(520, 136);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(66, 16);
			this.label7.TabIndex = 310;
			this.label7.Text = "Artículo (F3)";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtProducto
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProducto.Appearance = appearance6;
			this.txtProducto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProducto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProducto.Location = new System.Drawing.Point(616, 136);
			this.txtProducto.Name = "txtProducto";
			this.txtProducto.Size = new System.Drawing.Size(240, 22);
			this.txtProducto.TabIndex = 10;
			this.txtProducto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtProducto_KeyDown);
			this.txtProducto.Validated += new System.EventHandler(this.txtProducto_Validated);
			this.txtProducto.ValueChanged += new System.EventHandler(this.txtProducto_ValueChanged);
			// 
			// btnEnviar
			// 
			this.btnEnviar.Image = ((System.Drawing.Image)(resources.GetObject("btnEnviar.Image")));
			this.btnEnviar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEnviar.Location = new System.Drawing.Point(8, 312);
			this.btnEnviar.Name = "btnEnviar";
			this.btnEnviar.Size = new System.Drawing.Size(76, 23);
			this.btnEnviar.TabIndex = 12;
			this.btnEnviar.Text = "&Enviar";
			this.btnEnviar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEnviar.Click += new System.EventHandler(this.btnEnviar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(88, 312);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(76, 23);
			this.btnSalir.TabIndex = 13;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 296);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1040, 8);
			this.groupBox1.TabIndex = 315;
			this.groupBox1.TabStop = false;
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
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(85, 16);
			this.label1.TabIndex = 317;
			this.label1.Text = "Apellido Paterno";
			// 
			// txtApellidoPaterno
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoPaterno.Appearance = appearance7;
			this.txtApellidoPaterno.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoPaterno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoPaterno.Location = new System.Drawing.Point(104, 40);
			this.txtApellidoPaterno.MaxLength = 10;
			this.txtApellidoPaterno.Name = "txtApellidoPaterno";
			this.txtApellidoPaterno.Size = new System.Drawing.Size(120, 22);
			this.txtApellidoPaterno.TabIndex = 4;
			this.txtApellidoPaterno.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtApellidoPaterno_KeyDown);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(256, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 16);
			this.label3.TabIndex = 319;
			this.label3.Text = "Apellido Materno";
			// 
			// txtApellidoMaterno
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoMaterno.Appearance = appearance8;
			this.txtApellidoMaterno.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoMaterno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoMaterno.Location = new System.Drawing.Point(360, 40);
			this.txtApellidoMaterno.MaxLength = 10;
			this.txtApellidoMaterno.Name = "txtApellidoMaterno";
			this.txtApellidoMaterno.Size = new System.Drawing.Size(144, 22);
			this.txtApellidoMaterno.TabIndex = 5;
			this.txtApellidoMaterno.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtApellidoMaterno_KeyDown);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(768, 43);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(91, 16);
			this.label5.TabIndex = 323;
			this.label5.Text = "Segundo Nombre";
			// 
			// txtSegundoNombre
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSegundoNombre.Appearance = appearance9;
			this.txtSegundoNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSegundoNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSegundoNombre.Location = new System.Drawing.Point(872, 40);
			this.txtSegundoNombre.MaxLength = 10;
			this.txtSegundoNombre.Name = "txtSegundoNombre";
			this.txtSegundoNombre.Size = new System.Drawing.Size(150, 22);
			this.txtSegundoNombre.TabIndex = 7;
			this.txtSegundoNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSegundoNombre_KeyDown);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(520, 40);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 321;
			this.label8.Text = "Primer Nombre";
			// 
			// txtPrimerNombre
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrimerNombre.Appearance = appearance10;
			this.txtPrimerNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPrimerNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrimerNombre.Location = new System.Drawing.Point(616, 40);
			this.txtPrimerNombre.MaxLength = 10;
			this.txtPrimerNombre.Name = "txtPrimerNombre";
			this.txtPrimerNombre.Size = new System.Drawing.Size(128, 22);
			this.txtPrimerNombre.TabIndex = 6;
			this.txtPrimerNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPrimerNombre_KeyDown);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(8, 171);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(48, 16);
			this.label9.TabIndex = 325;
			this.label9.Text = "Telefono";
			// 
			// txtTelefono
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance11;
			this.txtTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Location = new System.Drawing.Point(104, 168);
			this.txtTelefono.MaxLength = 10;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(120, 22);
			this.txtTelefono.TabIndex = 8;
			this.txtTelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTelefono_KeyDown);
			this.txtTelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTelefono_KeyPress);
			// 
			// timer1
			// 
			this.timer1.Enabled = true;
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// uGridCedulas
			// 
			this.uGridCedulas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCedulas.CausesValidation = false;
			this.uGridCedulas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCedulas.DataSource = this.ultraDataSource1;
			appearance12.BackColor = System.Drawing.Color.White;
			this.uGridCedulas.DisplayLayout.Appearance = appearance12;
			this.uGridCedulas.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "SeteoF";
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Width = 75;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Width = 250;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Width = 120;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "Respuesta De";
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Width = 90;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Notas";
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Width = 225;
			ultraGridColumn16.Header.VisiblePosition = 9;
			ultraGridColumn16.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			ultraGridBand3.Header.Caption = "";
			this.uGridCedulas.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCedulas.DisplayLayout.Override.ActiveRowAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.uGridCedulas.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.White;
			appearance15.BackColor2 = System.Drawing.Color.White;
			this.uGridCedulas.DisplayLayout.Override.CellAppearance = appearance15;
			this.uGridCedulas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCedulas.DisplayLayout.Override.HeaderAppearance = appearance16;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCedulas.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCedulas.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCedulas.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.uGridCedulas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCedulas.Location = new System.Drawing.Point(8, 344);
			this.uGridCedulas.Name = "uGridCedulas";
			this.uGridCedulas.Size = new System.Drawing.Size(1016, 184);
			this.uGridCedulas.TabIndex = 326;
			this.uGridCedulas.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridCedulas_InitializeRow);
			this.uGridCedulas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridCedulas_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn2.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
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
																																 ultraDataColumn10});
			// 
			// timer2
			// 
			this.timer2.Interval = 5000;
			this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
			// 
			// chkAutomatico
			// 
			this.chkAutomatico.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAutomatico.Appearance = appearance20;
			this.chkAutomatico.BackColor = System.Drawing.Color.Transparent;
			this.chkAutomatico.CausesValidation = false;
			this.chkAutomatico.Enabled = false;
			this.chkAutomatico.Location = new System.Drawing.Point(928, 312);
			this.chkAutomatico.Name = "chkAutomatico";
			this.chkAutomatico.Size = new System.Drawing.Size(96, 22);
			this.chkAutomatico.TabIndex = 627;
			this.chkAutomatico.Text = "Automático";
			this.chkAutomatico.CheckedChanged += new System.EventHandler(this.chkAutomatico_CheckedChanged);
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Enabled = false;
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnActualizar.Location = new System.Drawing.Point(832, 312);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(76, 23);
			this.btnActualizar.TabIndex = 628;
			this.btnActualizar.Text = "&Actualizar";
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// label10
			// 
			this.label10.BackColor = System.Drawing.Color.Green;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.ForeColor = System.Drawing.Color.White;
			this.label10.Location = new System.Drawing.Point(480, 312);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(105, 25);
			this.label10.TabIndex = 629;
			this.label10.Text = "PENDIENTE";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label11
			// 
			this.label11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.ForeColor = System.Drawing.Color.White;
			this.label11.Location = new System.Drawing.Point(592, 312);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(105, 25);
			this.label11.TabIndex = 630;
			this.label11.Text = "APROBADO";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label12
			// 
			this.label12.BackColor = System.Drawing.Color.Red;
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.ForeColor = System.Drawing.Color.White;
			this.label12.Location = new System.Drawing.Point(704, 312);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(105, 25);
			this.label12.TabIndex = 631;
			this.label12.Text = "NEGADO";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// gbCarpa
			// 
			this.gbCarpa.Controls.Add(this.txtObservaciones);
			this.gbCarpa.Controls.Add(this.optAplica);
			this.gbCarpa.Controls.Add(this.chkCarpa);
			this.gbCarpa.Location = new System.Drawing.Point(632, 200);
			this.gbCarpa.Name = "gbCarpa";
			this.gbCarpa.Size = new System.Drawing.Size(392, 92);
			this.gbCarpa.TabIndex = 636;
			this.gbCarpa.TabStop = false;
			this.gbCarpa.Visible = false;
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.AcceptsReturn = true;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance21;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(112, 16);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(268, 66);
			this.txtObservaciones.TabIndex = 638;
			this.txtObservaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtObservaciones_KeyDown);
			// 
			// optAplica
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.optAplica.Appearance = appearance22;
			this.optAplica.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optAplica.CheckedIndex = 0;
			this.optAplica.Enabled = false;
			this.optAplica.ItemAppearance = appearance23;
			this.optAplica.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Si Aplica";
			valueListItem2.DataValue = 3;
			valueListItem2.DisplayText = "No Aplica";
			this.optAplica.Items.Add(valueListItem1);
			this.optAplica.Items.Add(valueListItem2);
			this.optAplica.ItemSpacingVertical = 10;
			this.optAplica.Location = new System.Drawing.Point(8, 24);
			this.optAplica.Name = "optAplica";
			this.optAplica.Size = new System.Drawing.Size(88, 56);
			this.optAplica.TabIndex = 637;
			this.optAplica.Text = "Si Aplica";
			// 
			// chkCarpa
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCarpa.Appearance = appearance24;
			this.chkCarpa.BackColor = System.Drawing.Color.Transparent;
			this.chkCarpa.CausesValidation = false;
			this.chkCarpa.Location = new System.Drawing.Point(0, 0);
			this.chkCarpa.Name = "chkCarpa";
			this.chkCarpa.Size = new System.Drawing.Size(56, 22);
			this.chkCarpa.TabIndex = 636;
			this.chkCarpa.Text = "Carpa";
			this.chkCarpa.CheckedChanged += new System.EventHandler(this.chkCarpa_CheckedChanged);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(8, 74);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(93, 16);
			this.label13.TabIndex = 638;
			this.label13.Text = "Fecha Nacimiento";
			// 
			// dtFechaNacimiento
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaNacimiento.Appearance = appearance25;
			this.dtFechaNacimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaNacimiento.DateButtons.Add(dateButton2);
			this.dtFechaNacimiento.Format = "dd/MM/yyyy";
			this.dtFechaNacimiento.Location = new System.Drawing.Point(104, 72);
			this.dtFechaNacimiento.Name = "dtFechaNacimiento";
			this.dtFechaNacimiento.NonAutoSizeHeight = 23;
			this.dtFechaNacimiento.Size = new System.Drawing.Size(120, 21);
			this.dtFechaNacimiento.SpinButtonsVisible = true;
			this.dtFechaNacimiento.TabIndex = 637;
			this.dtFechaNacimiento.Value = ((object)(resources.GetObject("dtFechaNacimiento.Value")));
			this.dtFechaNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaNacimiento_KeyDown);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(256, 72);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(79, 16);
			this.label14.TabIndex = 640;
			this.label14.Text = "Fecha Vivienda";
			// 
			// dtFechaVivienda
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaVivienda.Appearance = appearance26;
			this.dtFechaVivienda.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaVivienda.DateButtons.Add(dateButton3);
			this.dtFechaVivienda.Format = "dd/MM/yyyy";
			this.dtFechaVivienda.Location = new System.Drawing.Point(360, 72);
			this.dtFechaVivienda.Name = "dtFechaVivienda";
			this.dtFechaVivienda.NonAutoSizeHeight = 23;
			this.dtFechaVivienda.Size = new System.Drawing.Size(144, 21);
			this.dtFechaVivienda.SpinButtonsVisible = true;
			this.dtFechaVivienda.TabIndex = 639;
			this.dtFechaVivienda.Value = ((object)(resources.GetObject("dtFechaVivienda.Value")));
			this.dtFechaVivienda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaVivienda_KeyDown);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(520, 74);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(87, 16);
			this.label15.TabIndex = 642;
			this.label15.Text = "Tipo de Vivienda";
			// 
			// cmbVivienda
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVivienda.Appearance = appearance27;
			this.cmbVivienda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbVivienda.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVivienda.DataSource = this.ultraDataSource3;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 400;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbVivienda.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbVivienda.DisplayMember = "Nombre";
			this.cmbVivienda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVivienda.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVivienda.Location = new System.Drawing.Point(616, 72);
			this.cmbVivienda.Name = "cmbVivienda";
			this.cmbVivienda.Size = new System.Drawing.Size(128, 21);
			this.cmbVivienda.TabIndex = 641;
			this.cmbVivienda.ValueMember = "idTipoVivienda";
			this.cmbVivienda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVivienda_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(8, 106);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(75, 16);
			this.label16.TabIndex = 644;
			this.label16.Text = "Fecha Trabajo";
			// 
			// dtFechaTrabajo
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaTrabajo.Appearance = appearance28;
			this.dtFechaTrabajo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtFechaTrabajo.DateButtons.Add(dateButton4);
			this.dtFechaTrabajo.Format = "dd/MM/yyyy";
			this.dtFechaTrabajo.Location = new System.Drawing.Point(104, 104);
			this.dtFechaTrabajo.Name = "dtFechaTrabajo";
			this.dtFechaTrabajo.NonAutoSizeHeight = 23;
			this.dtFechaTrabajo.Size = new System.Drawing.Size(120, 21);
			this.dtFechaTrabajo.SpinButtonsVisible = true;
			this.dtFechaTrabajo.TabIndex = 643;
			this.dtFechaTrabajo.Value = ((object)(resources.GetObject("dtFechaTrabajo.Value")));
			this.dtFechaTrabajo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaTrabajo_KeyDown);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(256, 106);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(83, 16);
			this.label17.TabIndex = 646;
			this.label17.Text = "Tipo de Trabajo";
			// 
			// cmbTrabajo
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTrabajo.Appearance = appearance29;
			this.cmbTrabajo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTrabajo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTrabajo.DataSource = this.ultraDataSource2;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 300;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbTrabajo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbTrabajo.DisplayMember = "Nombre";
			this.cmbTrabajo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTrabajo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTrabajo.Location = new System.Drawing.Point(360, 104);
			this.cmbTrabajo.Name = "cmbTrabajo";
			this.cmbTrabajo.Size = new System.Drawing.Size(144, 21);
			this.cmbTrabajo.TabIndex = 645;
			this.cmbTrabajo.ValueMember = "idActEconomica";
			this.cmbTrabajo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTrabajo_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// optTipoFuenteIngresos
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipoFuenteIngresos.Appearance = appearance30;
			this.optTipoFuenteIngresos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTipoFuenteIngresos.ItemAppearance = appearance31;
			this.optTipoFuenteIngresos.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = false;
			valueListItem3.DisplayText = "Dependiente";
			valueListItem4.DataValue = true;
			valueListItem4.DisplayText = "Independiente";
			this.optTipoFuenteIngresos.Items.Add(valueListItem3);
			this.optTipoFuenteIngresos.Items.Add(valueListItem4);
			this.optTipoFuenteIngresos.ItemSpacingVertical = 10;
			this.optTipoFuenteIngresos.Location = new System.Drawing.Point(16, 135);
			this.optTipoFuenteIngresos.Name = "optTipoFuenteIngresos";
			this.optTipoFuenteIngresos.Size = new System.Drawing.Size(208, 24);
			this.optTipoFuenteIngresos.TabIndex = 639;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(256, 139);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(102, 16);
			this.label18.TabIndex = 647;
			this.label18.Text = "Fuente De Ingresos";
			this.label18.Click += new System.EventHandler(this.label18_Click);
			// 
			// txtEmpresaNegocio
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEmpresaNegocio.Appearance = appearance32;
			this.txtEmpresaNegocio.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtEmpresaNegocio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmpresaNegocio.Location = new System.Drawing.Point(360, 136);
			this.txtEmpresaNegocio.Name = "txtEmpresaNegocio";
			this.txtEmpresaNegocio.Size = new System.Drawing.Size(144, 22);
			this.txtEmpresaNegocio.TabIndex = 648;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(520, 104);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(86, 16);
			this.label19.TabIndex = 1013;
			this.label19.Text = "Valor Acreditado";
			// 
			// txtdIngreso
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdIngreso.Appearance = appearance33;
			this.txtdIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdIngreso.FormatString = "#,##0.00";
			this.txtdIngreso.Location = new System.Drawing.Point(616, 104);
			this.txtdIngreso.MinValue = 0;
			this.txtdIngreso.Name = "txtdIngreso";
			this.txtdIngreso.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtdIngreso.PromptChar = ' ';
			this.txtdIngreso.Size = new System.Drawing.Size(128, 22);
			this.txtdIngreso.TabIndex = 1012;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label20.Location = new System.Drawing.Point(768, 72);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(96, 17);
			this.label20.TabIndex = 1014;
			this.label20.Text = "Entidad Financiera";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEntidadFinanciera
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEntidadFinanciera.Appearance = appearance34;
			this.cmbEntidadFinanciera.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntidadFinanciera.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 208;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbEntidadFinanciera.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbEntidadFinanciera.DisplayMember = "Nombre";
			this.cmbEntidadFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntidadFinanciera.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntidadFinanciera.Location = new System.Drawing.Point(872, 72);
			this.cmbEntidadFinanciera.Name = "cmbEntidadFinanciera";
			this.cmbEntidadFinanciera.Size = new System.Drawing.Size(144, 21);
			this.cmbEntidadFinanciera.TabIndex = 1015;
			this.cmbEntidadFinanciera.ValueMember = "idEntidadFinanciera";
			this.cmbEntidadFinanciera.ValueChanged += new System.EventHandler(this.cmbEntidadFinanciera_ValueChanged);
			// 
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(936, 104);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.Size = new System.Drawing.Size(80, 23);
			this.btnComandos.TabIndex = 1049;
			this.btnComandos.Text = "Link Nuovo";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Visible = false;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click);
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesCopiarUltimoMes,
																																								this.mnuAccionesAgregarNuevoLocal,
																																								this.menuItem1,
																																								this.mnuAccionesActualizar});
			// 
			// mnuAccionesCopiarUltimoMes
			// 
			this.mnuAccionesCopiarUltimoMes.Index = 0;
			this.mnuAccionesCopiarUltimoMes.Text = "Código QR Instalar APP colateral";
			this.mnuAccionesCopiarUltimoMes.Click += new System.EventHandler(this.mnuAccionesCopiarUltimoMes_Click);
			// 
			// mnuAccionesAgregarNuevoLocal
			// 
			this.mnuAccionesAgregarNuevoLocal.Index = 1;
			this.mnuAccionesAgregarNuevoLocal.Text = "Código QR Enrolar";
			this.mnuAccionesAgregarNuevoLocal.Click += new System.EventHandler(this.mnuAccionesAgregarNuevoLocal_Click);
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 2;
			this.menuItem1.Text = "Código QR Huawei Enrolar";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// mnuAccionesActualizar
			// 
			this.mnuAccionesActualizar.Index = 3;
			this.mnuAccionesActualizar.Text = "Dispositivos Secured by Knox";
			this.mnuAccionesActualizar.Click += new System.EventHandler(this.mnuAccionesActualizar_Click);
			// 
			// frmCre_EnvioCedulasVendedores
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1034, 536);
			this.Controls.Add(this.btnComandos);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbEntidadFinanciera);
			this.Controls.Add(this.txtdIngreso);
			this.Controls.Add(this.txtEmpresaNegocio);
			this.Controls.Add(this.dtFechaTrabajo);
			this.Controls.Add(this.cmbVivienda);
			this.Controls.Add(this.dtFechaVivienda);
			this.Controls.Add(this.dtFechaNacimiento);
			this.Controls.Add(this.gbCarpa);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.chkAutomatico);
			this.Controls.Add(this.uGridCedulas);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.txtSegundoNombre);
			this.Controls.Add(this.txtPrimerNombre);
			this.Controls.Add(this.txtApellidoMaterno);
			this.Controls.Add(this.txtApellidoPaterno);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnEnviar);
			this.Controls.Add(this.txtProducto);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.optTipoFuenteIngresos);
			this.Controls.Add(this.cmbTrabajo);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCre_EnvioCedulasVendedores";
			this.Text = "Envio de Cedulas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_EnvioCedulasVendedores_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_EnvioCedulasVendedores_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProducto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoPaterno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoMaterno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSegundoNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrimerNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCedulas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.gbCarpa.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optAplica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaTrabajo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTrabajo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoFuenteIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresaNegocio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		int idCre_Solicitud = 0;
		int idEstado = 0;
		string sEstado = "";
		string sNegado = "";
		string sCodigo = "";
		int idNegadoPendiente = 0;
		int idArticulo = 0;
		int idEntidad = 0;
		int idArtGrup =0;
		int idValEntidad =0;
		private Acceso miAcceso;
		#endregion Variables

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		int iBodega = 0;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCre_EnvioCedulasVendedores_Load(object sender, System.EventArgs e)
		{
			string stTipo = "0825";
			string stPermiso = "635";
							
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			
			if (!Funcion.Permiso(stPermiso, cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Envio De Cedulas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
				iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

		
				sSQL = string.Format("Exec ListaBodegas {0}, 43", MenuLatinium.IdUsuario);
				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				//this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalVendedor {0}", iBodega));

				if (iBodega == 0) 
				{
					this.cmbBodega.Value = 12; 
					this.cmbBodega.Enabled = true;
				}
				else this.cmbBodega.Value = iBodega;

				this.dtFecha.Value = DateTime.Now;

				this.dtFechaNacimiento.CalendarInfo.MinDate = DateTime.Today.AddYears(-75);
				this.dtFechaNacimiento.CalendarInfo.MaxDate = DateTime.Today.AddYears(-18);
				this.dtFechaVivienda.CalendarInfo.MaxDate = DateTime.Today.AddYears(-1);
				this.dtFechaTrabajo.CalendarInfo.MaxDate = DateTime.Today.AddYears(-1);

				if (miAcceso.BCarpa) this.gbCarpa.Visible = true;

				this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
				this.cmbVivienda.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 88, 0");
				this.cmbTrabajo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 87, 2");
				bool PointMovil = Funcion.bEscalarSQL(cdsSeteoF, string.Format("select PointMovil from  Bodega where Bodega = {0}", (int)this.cmbBodega.Value), true);
				if (!PointMovil)
				{
					this.cmbEntidadFinanciera.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera  where idEntidadFinanciera in (1) Order By idEntidadFinanciera");
				}
				else
				{
					this.cmbEntidadFinanciera.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera  where idEntidadFinanciera in (1,14) Order By idEntidadFinanciera");
				}

				this.chkAutomatico.Checked = true;
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void txtProducto_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void frmCre_EnvioCedulasVendedores_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtRuc.Focus();
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtApellidoPaterno.Focus();
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtRuc.Text.ToString().Trim().Length > 0)
			{
				string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', 5)", this.txtRuc.Text.ToString());
				string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
				if (VlCed.StartsWith("Error"))
				{
					MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtRuc.Focus();
					e.Cancel = true;				
				}
			}
		}

		private void Controles()
		{
			#region Inicializa Variables - Controles
			idCre_Solicitud = 0;
			idEstado = 0;
			sEstado = "";
			idNegadoPendiente = 0;			
			sNegado = "";
			idArticulo = 0;

			this.txtApellidoPaterno.Text = "";
			this.txtApellidoMaterno.Text = "";
			this.txtPrimerNombre.Text = "";
			this.txtSegundoNombre.Text = "";
			this.txtEmpresaNegocio.Text = "";
			this.txtdIngreso.Value=0.00;
			this.txtNumero.Text = "";
			
			this.dtFechaNacimiento.Value = System.DBNull.Value;
			this.dtFechaNacimiento.Enabled = true;
			this.dtFechaVivienda.Value = System.DBNull.Value;
			this.cmbVivienda.Value = System.DBNull.Value;
			this.cmbEntidadFinanciera.Value = System.DBNull.Value;
			this.dtFechaTrabajo.Value = System.DBNull.Value;
			this.cmbTrabajo.Value = System.DBNull.Value;
			//this.optTipoFuenteIngresos.Value = System.DBNull.Value;
			this.txtTelefono.Text = "";

			this.cmbVendedor.Value = System.DBNull.Value;
			this.txtProducto.Text = "";
			this.txtNotas.Text = "";

			this.chkCarpa.Checked = false;
			this.optAplica.Enabled = false;
			this.txtObservaciones.Enabled = false;			
			this.txtObservaciones.Text = "";
			#endregion Inicializa Variables - Controles
		}

		private void Consulta()
		{
			this.uGridCedulas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_RespuestaEnvioCedulas {0}, '{1}'", Convert.ToInt32(this.cmbBodega.Value), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			Controles();

			#region Busca Cliente
			string sSQL = string.Format("Exec Cre_ECRetornaCliente '{0}'", this.txtRuc.Text.ToString());
			SqlDataReader drCliente = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drCliente.Read();
			if (drCliente.HasRows)
			{				
				idCre_Solicitud = drCliente.GetInt32(0);
				this.txtApellidoPaterno.Text = drCliente.GetString(1);
				this.txtApellidoMaterno.Text = drCliente.GetString(2);
				this.txtPrimerNombre.Text = drCliente.GetString(3);
				this.txtSegundoNombre.Text = drCliente.GetString(4);
				idEstado = drCliente.GetInt32(5);
				sEstado = drCliente.GetString(6);
				idNegadoPendiente = drCliente.GetInt32(7);
				sNegado = drCliente.GetString(8);
				this.optTipoFuenteIngresos.Value = drCliente.GetBoolean(9);
				this.txtEmpresaNegocio.Text = drCliente.GetString(10);
				if (drCliente.GetValue(11) != System.DBNull.Value) this.dtFechaTrabajo.Value = drCliente.GetDateTime(11);
//				if (drCliente.GetValue(12) != System.DBNull.Value) 
//				{
//					this.dtFechaNacimiento.Value = drCliente.GetDateTime(12);
//					this.dtFechaNacimiento.Enabled = false;
//				}
				if (drCliente.GetValue(13) != System.DBNull.Value) this.dtFechaVivienda.Value = drCliente.GetDateTime(13);
				if (drCliente.GetValue(14) != System.DBNull.Value) this.cmbVivienda.Value = drCliente.GetInt32(14);
				//if (drCliente.GetValue(15) != System.DBNull.Value) this.cmbTrabajo.Value = drCliente.GetInt32(15);	
				if (drCliente.GetValue(16) != System.DBNull.Value) this.txtdIngreso.Value = drCliente.GetInt32(16);
			}			
			drCliente.Close();			
			#endregion Busca Cliente
		}

		private void btnEnviar_Click(object sender, System.EventArgs e)
		{
			#region Valida Campos Obligatorios
			#region Valida identificación
			if (!Validacion.vbIdentificacion(this.txtRuc, 5, cdsSeteoF)) return;
			#endregion Valida identificación

			#region Nombres
			if (this.txtApellidoPaterno.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese el Apellido Paterno", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtApellidoPaterno.Focus();
				return;
			}

			if (this.txtPrimerNombre.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese el Primer Nombre", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtPrimerNombre.Focus();
				return;
			}
			#endregion Nombres

			#region Bloqueos
/*
			#region Datos Credito
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaNacimiento, "Seleccione la fecha de nacimiento", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaVivienda, "Seleccione la fecha inicio de vivienda", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbVivienda, "Seleccione el tipo de vivienda")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaTrabajo, "Seleccione la fecha de inicio de trabajo", false, "", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbTrabajo, "Seleccione el tipo de trabajo")) return;
			if (!Validacion.vbComboVacio(this.cmbEntidadFinanciera, "Seleccione Entidad Financiera")) return;
			if (this.optTipoFuenteIngresos.CheckedIndex == -1)
			{
				MessageBox.Show("Seleccione el tipo de fuente de ingresos del cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.optTipoFuenteIngresos.Focus();
				return;
			}
			if (!Validacion.vbTexto(this.txtEmpresaNegocio, 5, 100, "Fuente De Ingresos")) return;
			#endregion Datos Credito

			#region Telefono
			if (this.txtTelefono.Text.ToString().Trim().Length <= 9) if (!Validacion.vbValidaTelefono(this.txtTelefono, 1, true)) return;
			if (this.txtTelefono.Text.ToString().Trim().Length > 9) if (!Validacion.vbValidaTelefono(this.txtTelefono, 2, true)) return;
			#endregion Telefono

			#region Vendedor
			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un vendedor")) return;			
			#endregion Vendedor

			#region Articulo
			if (idArticulo == 0)
			{
				MessageBox.Show("Ingrese un Producto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtProducto.Focus();
				return;
			}

			string sSQLValArticulo = string.Format("Exec Cre_ValidaValorProductoEnvioCedulas {0}", idArticulo);
			string sMensValArticulo = Funcion.sEscalarSQL(cdsSeteoF, sSQLValArticulo);

			if (sMensValArticulo.Length > 0)
			{
				MessageBox.Show(sMensValArticulo + "\n\nINGRESE UN PRODUCTO CORRECTO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtProducto.Focus();
				return;
			}
			#endregion Articulo	
		
			if (!Validacion.vbCampoDecimalVacio (this.txtdIngreso,"Ingrese el Monto de Ingresos del Cliente ",1,10000)) return;	
			
*/
			#endregion Bloqueos
#endregion Valida Campos Obligatorios

			this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionEnviosCedula {0}", Convert.ToInt32(this.cmbBodega.Value)));

			DateTime dtFN = DateTime.Parse("01/01/1900");
			DateTime dtFV = DateTime.Parse("01/01/1900");
			DateTime dtFT = DateTime.Parse("01/01/1900");
			int idVendedor = 0;

			#region Guarda
			string sSQL = string.Format("Exec Cre_GuardaEnvioRevisionCedulas '{0}', '{1}', {2}, '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', {9}, {10}, {11}, {12}, '{13}', '{14}', '{15}', {16}, '{17}', {18}, {19}, '{20}', {21}, 0, {22}",
				this.txtNumero.Text.ToString(), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm:ss"), idCre_Solicitud, 
				this.txtRuc.Text.ToString(), this.txtApellidoPaterno.Text.ToString().Trim(), this.txtApellidoMaterno.Text.ToString().Trim(), this.txtPrimerNombre.Text.ToString().Trim(), this.txtSegundoNombre.Text.ToString().Trim(),
				this.txtTelefono.Text.ToString(), Convert.ToInt32(this.cmbBodega.Value), 0, idArticulo, 1, this.txtNotas.Text.ToString().Trim(),
				dtFN.ToString("yyyyMMdd"), dtFV.ToString("yyyyMMdd"), 0,
				dtFT.ToString("yyyyMMdd"), 0, (bool)this.optTipoFuenteIngresos.Value, this.txtEmpresaNegocio.Text.ToString(), this.txtdIngreso.Value, idEntidad);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			#endregion Guarda

			string sMensaje = "Cedula Enviada";
			
			#region Respuesta del Sistema
			if (idEstado == 3) 
			{
				sMensaje = string.Format("Solicitud Rechazada : {0}", sNegado);
				
				string sRespuesta = string.Format("Mensaje automatico generado por el Administrador Solicitud Rechazada : {0}", sNegado);

				string sSQLRespuesta = string.Format("Exec Cre_RespuestaHistorialRevisionCedulas {0}, {1}, '{2}', {3}, {4}, 'Administrador', '{5}', 2, {6}", 
					idCre_Solicitud, Convert.ToInt32(this.cmbBodega.Value), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
					idEstado, idNegadoPendiente, sRespuesta, this.chkCarpa.Checked);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLRespuesta);
			}
			#endregion Respuesta del Sistema

			#region Carpa
			if (this.chkCarpa.Checked)
			{
				if (idEstado == 3) 
				{
					this.optAplica.Value = 3;
					this.txtObservaciones.Text = this.txtObservaciones.Text.ToString() + " " + sMensaje;
				}			

				string sSQLCarpa = string.Format("Exec Cre_RespuestaHistorialRevisionCedulas {0}, {1}, '{2}', {3}, {4}, '{5}', '{6}', 2, {7}", 
					idCre_Solicitud, Convert.ToInt32(this.cmbBodega.Value), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
					(int)this.optAplica.Value, 0, MenuLatinium.stUsuario, this.txtObservaciones.Text.ToString(), this.chkCarpa.Checked);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLCarpa);
			}
			#endregion Carpa
	
			MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			Controles();

			this.txtRuc.Text = "";
			this.txtRuc.Focus();
		}

		private void txtApellidoPaterno_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtApellidoMaterno.Focus();
		}

		private void txtApellidoMaterno_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPrimerNombre.Focus();
		}

		private void txtPrimerNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtSegundoNombre.Focus();
		}

		private void txtSegundoNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFechaNacimiento.Focus();
		}

		private void txtTelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbVendedor.Focus();
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtProducto.Focus();
		}

		private void txtProducto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNotas.Focus();

			if (e.KeyCode == Keys.F3)
			{
				int iBodega = 0;
				if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

				using (BuscaArticulos miBusqueda = new BuscaArticulos(this.txtProducto.Text.ToString(), iBodega, 43, 9, DateTime.Today, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.txtProducto.Text = miBusqueda.sCodigo;
						idArticulo = (int)miBusqueda.idArticulo;
					}
				}
			}
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnEnviar.Focus();
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtTelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtProducto_Validated(object sender, System.EventArgs e)
		{
			
			idEntidad=Convert.ToInt32(this.cmbEntidadFinanciera.Value);
			if(idEntidad==0)
			{
				MessageBox.Show("Seleccione Entidad Financiera", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				idArticulo = 0;
				this.txtProducto.Text = "";	
			
			}
			else
			{
				if (this.txtProducto.Text.ToString().Length > 0)
				{
					string sSQL = string.Format("Select idArticulo, Codigo From Articulo Where (Codigo = '{0}' Or Articulo = '{0}')", this.txtProducto.Text.ToString().Trim()); 
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						idArticulo = drArticulo.GetInt32(0);
						sCodigo = drArticulo.GetString(1);	
						drArticulo.Close();
						idArtGrup = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select idGrupoArticulo from Articulo  where idArticulo= {0}",idArticulo));
						if (idEntidad==1)
						{
							if (idArtGrup == 8)
							{
								MessageBox.Show("Para celulares la entidad  financiera  debe ser  POINT MOVIL", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								idArticulo = 0;
								this.txtProducto.Text = "";	
							}
							else 
							{
								this.txtProducto.Text = sCodigo;	
							}
						}
						else if (idEntidad==14)
						{
							if (idArtGrup!= 8)
							{
								MessageBox.Show("POINT MOVIL no permite estos productos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								idArticulo = 0;
								this.txtProducto.Text = "";	
							}
							else  if (idArtGrup == 8)
							{
								int sSCodigo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select  COUNT(*) from  Articulo  where  Preciomovil =1 and idArticulo ={0}",idArticulo));
								if(sSCodigo ==0)
								{
									MessageBox.Show("Producto no disponible para Point Movil", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									idArticulo = 0;
									this.txtProducto.Text = "";	
									//return;
								}
								else
								{
									this.txtProducto.Text = sCodigo;
								}
							}
						}
					}										
					else
					{
						MessageBox.Show("No existe el Producto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						idArticulo = 0;
					}
				}
			}
		}

		private void timer1_Tick(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Now;
		}

		private void timer2_Tick(object sender, System.EventArgs e)
		{
			this.Consulta();
		}

		private void uGridCedulas_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (e.Row.Cells["iEstado"].Value.ToString().Trim() == "1") 
			{
				e.Row.Appearance.BackColor = Color.Green; //EN ESPERA
				e.Row.Appearance.BackColor2 = Color.Green;
				e.Row.Appearance.ForeColor = Color.White;
				e.Row.Appearance.ForeColorDisabled = Color.White;
			}
			if (e.Row.Cells["iEstado"].Value.ToString().Trim() == "2") 
			{
				if (e.Row.Cells["idEstado"].Value.ToString().Trim() == "3")
				{
					e.Row.Appearance.BackColor = Color.Red; // RESPONDIDO NEGADO
					e.Row.Appearance.BackColor2 = Color.Red;
					e.Row.Appearance.ForeColor = Color.White;
					e.Row.Appearance.ForeColorDisabled = Color.White;
				}
				else
				{
					e.Row.Appearance.BackColor = Color.Blue; // RESPONDIDO APROBADO
					e.Row.Appearance.BackColor2 = Color.Blue;
					e.Row.Appearance.ForeColor = Color.White;
					e.Row.Appearance.ForeColorDisabled = Color.White;
				}
			}			
		}

		private void uGridCedulas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkAutomatico_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.chkAutomatico.Checked)
			{
				this.Consulta();
				this.timer2.Enabled = true;				
			}
			else
			{
				this.timer2.Enabled = false;
			}
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			this.Consulta();
		}

		private void optFiltro_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkCarpa_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.chkCarpa.Checked)
			{
				this.optAplica.Enabled = true;
				this.txtObservaciones.Enabled = true;
				this.txtObservaciones.Focus();
			}
			else
			{
				this.optAplica.Enabled = false;
				this.txtObservaciones.Enabled = false;
				this.txtObservaciones.Text = "";
			}
		}

		private void txtNotas_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnEnviar.Focus();
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null) return;
			if (this.dtFecha.Value == System.DBNull.Value) return;

			this.cmbVendedor.Value = System.DBNull.Value;

			this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

			bool PointMovil = Funcion.bEscalarSQL(cdsSeteoF, string.Format("select PointMovil from  Bodega where Bodega = {0}", (int)this.cmbBodega.Value), true);
			if (!PointMovil)
			{
				this.cmbEntidadFinanciera.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera  where idEntidadFinanciera in (1) Order By idEntidadFinanciera");
			}
			else
			{
				this.cmbEntidadFinanciera.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera  where idEntidadFinanciera in (1,14) Order By idEntidadFinanciera");
			}
		}

		private void cmbTrabajo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtTelefono.Focus();
		}

		private void dtFechaNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFechaVivienda.Focus();
		}

		private void dtFechaVivienda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbVivienda.Focus();
		}

		private void cmbVivienda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFechaTrabajo.Focus();
		}

		private void dtFechaTrabajo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTrabajo.Focus();
		}

		private void label18_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbEntidadFinanciera_ValueChanged(object sender, System.EventArgs e)
		{
			idArticulo = 0;
			this.txtProducto.Text = "";	
		}

		private void linkLabel2_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			System.Diagnostics.Process.Start("chrome.exe", "https://storage.googleapis.com/sparta_bucket/Principal/DESAPOINT/0916878820001_000001213_1279118/QR_NUOVO.PNG");
		}

		private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			System.Diagnostics.Process.Start("chrome.exe", "https://www.samsungknox.com/en/knox-platform/supported-devices");
		}

		private void linkLabel3_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			System.Diagnostics.Process.Start("chrome.exe", "https://storage.googleapis.com/sparta_bucket/Principal/DESAPOINT/0901020396001_000000524_1279743/QR_NUOVO.PNG");
			
		}

		private void mnuAccionesCopiarUltimoMes_Click(object sender, System.EventArgs e)
		{
			//this.linkLabel3_LinkClicked(sender,e);
			System.Diagnostics.Process.Start("chrome.exe", "https://storage.googleapis.com/sparta_bucket/Principal/DESAPOINT/1792605504001_000236550_1274424/QrDescaragNuovo.png");
			
		}

		private void mnuAccionesAgregarNuevoLocal_Click(object sender, System.EventArgs e)
		{
		//	this.linkLabel2_LinkClicked(sender,e);
		System.Diagnostics.Process.Start("chrome.exe", "https://storage.googleapis.com/sparta_bucket/Principal/DESAPOINT/0916878820001_000001213_1279118/QR_NUOVO.PNG");
		
		}

		private void mnuAccionesActualizar_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process.Start("chrome.exe", "https://www.samsungknox.com/en/knox-platform/supported-devices");
		
		}

		private void btnComandos_Click(object sender, System.EventArgs e)
		{
			this.mnuAcciones.Show(this.btnComandos, new Point(1, 25));
	
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process.Start("chrome.exe", "https://storage.googleapis.com/sparta_bucket/Principal/DESAPOINT/1204875023001_000001228_1279133/Qr_Huawei.png");
		
		}
	}
}

