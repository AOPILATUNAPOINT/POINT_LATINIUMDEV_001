using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_CambioVendedor.
	/// </summary>
	public class frmCom_CambioVendedor : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumeroDias;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCom_AsignacionDeVendedores;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblDescuento;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCom_AsignacionDeVendedoresN;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRango;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnImprimir;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCargo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpHasta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEnvioEMail;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEnvio;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiarioCarpa;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMensualCarpa;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGrupo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtContado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSemanal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGEx;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdUsuario;
		private C1.Data.C1DataSet cdsSeteo;
		private C1.Data.C1DataSet cdsSeguridad;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_CambioVendedor()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idPersonal = 0;
		int idCom_AsignaPersonal = 0;		
		bool bNuevo = false;
		string sPeriodo = ""; 
		int idEstado = 0;
		int idCargo = 0;
		int idRango = 0;
		DateTime dtDesde = DateTime.Today;
		DateTime dtHasta = DateTime.Today;
		DateTime dtUltimoDiaMes = DateTime.Today;
		DateTime dtUltimoDiaSemana = DateTime.Today;
		int iBodega = 0;
		int iNumero = 0;
		int idUsuario = 0;

		public frmCom_CambioVendedor(int IdPersonal, int IdCom_AsignaPersonal, bool BNuevo, string SPeriodo, int IdEstado, int IdCargo, int IdRango, DateTime DTDesde, DateTime DTHasta, DateTime DTUltimoDiaMes, int IBodega, int IdUsuario)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idPersonal = IdPersonal;
			idCom_AsignaPersonal = IdCom_AsignaPersonal;			
			sPeriodo = SPeriodo;
			bNuevo = BNuevo;
			idEstado = IdEstado;
			idCargo = IdCargo;
			idRango = IdRango;
			dtDesde = DTDesde;
			dtHasta = DTHasta;
			dtUltimoDiaMes = DTUltimoDiaMes;
			iBodega = IBodega;
			idUsuario = IdUsuario;
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		/*  ASIGNACION SEMANAL */
		public frmCom_CambioVendedor(int IdPersonal, int IdCom_AsignaPersonal, bool BNuevo, int IdEstado, int IdCargo, int IdRango, DateTime DTDesde, DateTime DTHasta, DateTime DTUltimoDiaSemana, int IBodega, int INumero)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idPersonal = IdPersonal;
			idCom_AsignaPersonal = IdCom_AsignaPersonal;						
			bNuevo = BNuevo;
			idEstado = IdEstado;
			idCargo = IdCargo;
			idRango = IdRango;
			dtDesde = DTDesde;
			dtHasta = DTHasta;
			dtUltimoDiaSemana = DTUltimoDiaSemana;
			iBodega = IBodega;
			iNumero = INumero;
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rango");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_CambioVendedor));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.dtpDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.dtpHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbRango = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNumeroDias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCom_AsignacionDeVendedores = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.txtContado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSemanal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCom_AsignacionDeVendedoresN = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.cmbCargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtGEx = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.dtFechaEnvio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkEnvioEMail = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtDiarioCarpa = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMensualCarpa = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkGrupo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtIdUsuario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.dtpDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRango)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroDias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCom_AsignacionDeVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSemanal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCom_AsignacionDeVendedoresN)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGEx)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEnvio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiarioCarpa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMensualCarpa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			this.SuspendLayout();
			// 
			// dtpDesde
			// 
			appearance1.FontData.Name = "Tahoma";
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpDesde.Appearance = appearance1;
			this.dtpDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpDesde.DateButtons.Add(dateButton1);
			this.dtpDesde.Enabled = false;
			this.dtpDesde.Format = "dd/MM/yyyy";
			this.dtpDesde.Location = new System.Drawing.Point(72, 96);
			this.dtpDesde.Name = "dtpDesde";
			this.dtpDesde.NonAutoSizeHeight = 23;
			this.dtpDesde.Size = new System.Drawing.Size(112, 21);
			this.dtpDesde.SpinButtonsVisible = true;
			this.dtpDesde.TabIndex = 2;
			this.dtpDesde.Value = new System.DateTime(2018, 8, 7, 0, 0, 0, 0);
			this.dtpDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			this.dtpDesde.ValueChanged += new System.EventHandler(this.dtpDesde_ValueChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 98);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(35, 17);
			this.label6.TabIndex = 169;
			this.label6.Text = "Desde";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpHasta
			// 
			appearance2.FontData.Name = "Tahoma";
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpHasta.Appearance = appearance2;
			this.dtpHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtpHasta.DateButtons.Add(dateButton2);
			this.dtpHasta.Enabled = false;
			this.dtpHasta.Format = "dd/MM/yyyy";
			this.dtpHasta.Location = new System.Drawing.Point(256, 96);
			this.dtpHasta.Name = "dtpHasta";
			this.dtpHasta.NonAutoSizeHeight = 23;
			this.dtpHasta.Size = new System.Drawing.Size(112, 21);
			this.dtpHasta.SpinButtonsVisible = true;
			this.dtpHasta.TabIndex = 3;
			this.dtpHasta.Value = new System.DateTime(2018, 8, 7, 0, 0, 0, 0);
			this.dtpHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtHasta_KeyDown);
			this.dtpHasta.ValueChanged += new System.EventHandler(this.dtpHasta_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(200, 98);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 17);
			this.label1.TabIndex = 171;
			this.label1.Text = "Hasta";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.Location = new System.Drawing.Point(8, 131);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 17);
			this.lblBodega.TabIndex = 173;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(8, 163);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 17);
			this.label2.TabIndex = 175;
			this.label2.Text = "Cargo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(8, 195);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(52, 17);
			this.label3.TabIndex = 177;
			this.label3.Text = "Categoria";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbRango
			// 
			this.cmbRango.CausesValidation = false;
			this.cmbRango.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRango.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRango.DataSource = this.ultraDataSource3;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 300;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbRango.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbRango.DisplayMember = "Rango";
			this.cmbRango.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRango.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRango.Enabled = false;
			this.cmbRango.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRango.Location = new System.Drawing.Point(72, 192);
			this.cmbRango.MaxDropDownItems = 30;
			this.cmbRango.Name = "cmbRango";
			this.cmbRango.Size = new System.Drawing.Size(296, 22);
			this.cmbRango.TabIndex = 6;
			this.cmbRango.ValueMember = "idCom_Rango";
			this.cmbRango.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbRango_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(8, 67);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(38, 17);
			this.label4.TabIndex = 179;
			this.label4.Text = "Estado";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEstado
			// 
			this.cmbEstado.CausesValidation = false;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.ultraDataSource1;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEstado.DisplayMember = "Estado";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Enabled = false;
			this.cmbEstado.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(72, 64);
			this.cmbEstado.MaxDropDownItems = 30;
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(296, 22);
			this.cmbEstado.TabIndex = 1;
			this.cmbEstado.ValueMember = "idCom_Estado";
			this.cmbEstado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstado_KeyDown);
			this.cmbEstado.ValueChanged += new System.EventHandler(this.cmbEstado_ValueChanged);
			this.cmbEstado.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEstado_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// txtNumeroDias
			// 
			this.txtNumeroDias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroDias.Enabled = false;
			this.txtNumeroDias.Location = new System.Drawing.Point(472, 64);
			this.txtNumeroDias.Name = "txtNumeroDias";
			this.txtNumeroDias.PromptChar = ' ';
			this.txtNumeroDias.Size = new System.Drawing.Size(80, 22);
			this.txtNumeroDias.TabIndex = 275;
			// 
			// txtIdCom_AsignacionDeVendedores
			// 
			this.txtIdCom_AsignacionDeVendedores.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCom_AsignacionDeVendedores.Enabled = false;
			this.txtIdCom_AsignacionDeVendedores.Location = new System.Drawing.Point(512, 304);
			this.txtIdCom_AsignacionDeVendedores.Name = "txtIdCom_AsignacionDeVendedores";
			this.txtIdCom_AsignacionDeVendedores.PromptChar = ' ';
			this.txtIdCom_AsignacionDeVendedores.Size = new System.Drawing.Size(16, 22);
			this.txtIdCom_AsignacionDeVendedores.TabIndex = 274;
			this.txtIdCom_AsignacionDeVendedores.Visible = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(384, 67);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(69, 17);
			this.label5.TabIndex = 276;
			this.label5.Text = "Numero Dias";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubTotal.Location = new System.Drawing.Point(384, 98);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(45, 17);
			this.lblSubTotal.TabIndex = 278;
			this.lblSubTotal.Text = "Mensual";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(384, 227);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(46, 17);
			this.lblDescuento.TabIndex = 277;
			this.lblDescuento.Text = "Contado";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblDescuento.Visible = false;
			// 
			// txtContado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtContado.Appearance = appearance3;
			this.txtContado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContado.Enabled = false;
			this.txtContado.FormatString = "#,##0.00";
			this.txtContado.Location = new System.Drawing.Point(472, 224);
			this.txtContado.Name = "txtContado";
			this.txtContado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtContado.PromptChar = ' ';
			this.txtContado.Size = new System.Drawing.Size(80, 22);
			this.txtContado.TabIndex = 280;
			this.txtContado.Visible = false;
			// 
			// txtSemanal
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSemanal.Appearance = appearance4;
			this.txtSemanal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSemanal.Enabled = false;
			this.txtSemanal.FormatString = "#,##0.00";
			this.txtSemanal.Location = new System.Drawing.Point(472, 95);
			this.txtSemanal.Name = "txtSemanal";
			this.txtSemanal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSemanal.PromptChar = ' ';
			this.txtSemanal.Size = new System.Drawing.Size(80, 22);
			this.txtSemanal.TabIndex = 279;
			// 
			// txtIdCom_AsignacionDeVendedoresN
			// 
			this.txtIdCom_AsignacionDeVendedoresN.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCom_AsignacionDeVendedoresN.Enabled = false;
			this.txtIdCom_AsignacionDeVendedoresN.Location = new System.Drawing.Point(528, 304);
			this.txtIdCom_AsignacionDeVendedoresN.Name = "txtIdCom_AsignacionDeVendedoresN";
			this.txtIdCom_AsignacionDeVendedoresN.PromptChar = ' ';
			this.txtIdCom_AsignacionDeVendedoresN.Size = new System.Drawing.Size(16, 22);
			this.txtIdCom_AsignacionDeVendedoresN.TabIndex = 281;
			this.txtIdCom_AsignacionDeVendedoresN.Visible = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 290);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(562, 8);
			this.groupBox1.TabIndex = 282;
			this.groupBox1.TabStop = false;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 306);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 9;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.CausesValidation = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 306);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 7;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 10);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(52, 17);
			this.label7.TabIndex = 286;
			this.label7.Text = "Vendedor";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(568, 8);
			this.groupBox2.TabIndex = 287;
			this.groupBox2.TabStop = false;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(168, 306);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(72, 23);
			this.btnImprimir.TabIndex = 8;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Visible = false;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// cmbCargo
			// 
			this.cmbCargo.CausesValidation = false;
			this.cmbCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargo.DataSource = this.ultraDataSource2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 300;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCargo.DisplayMember = "Cargo";
			this.cmbCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCargo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCargo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargo.Location = new System.Drawing.Point(72, 160);
			this.cmbCargo.MaxDropDownItems = 30;
			this.cmbCargo.Name = "cmbCargo";
			this.cmbCargo.Size = new System.Drawing.Size(296, 22);
			this.cmbCargo.TabIndex = 5;
			this.cmbCargo.ValueMember = "idCargo";
			this.cmbCargo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCargo_KeyDown);
			this.cmbCargo.ValueChanged += new System.EventHandler(this.cmbCargo_ValueChanged);
			this.cmbCargo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCargo_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource4;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 300;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(72, 128);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(296, 22);
			this.cmbBodega.TabIndex = 4;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
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
			// cmbVendedor
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance5;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataSource = this.ultraDataSource5;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 472;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(80, 8);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(472, 22);
			this.cmbVendedor.TabIndex = 0;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVendedor_KeyDown);
			this.cmbVendedor.ValueChanged += new System.EventHandler(this.cmbVendedor_ValueChanged);
			this.cmbVendedor.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbVendedor_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(384, 163);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(40, 17);
			this.label8.TabIndex = 289;
			this.label8.Text = "Credito";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(384, 195);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(27, 17);
			this.label9.TabIndex = 288;
			this.label9.Text = "G Ex";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtGEx
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGEx.Appearance = appearance6;
			this.txtGEx.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGEx.Enabled = false;
			this.txtGEx.FormatString = "#,##0.00";
			this.txtGEx.Location = new System.Drawing.Point(472, 192);
			this.txtGEx.Name = "txtGEx";
			this.txtGEx.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGEx.PromptChar = ' ';
			this.txtGEx.Size = new System.Drawing.Size(80, 22);
			this.txtGEx.TabIndex = 291;
			// 
			// txtCredito
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCredito.Appearance = appearance7;
			this.txtCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCredito.Enabled = false;
			this.txtCredito.FormatString = "#,##0.00";
			this.txtCredito.Location = new System.Drawing.Point(472, 160);
			this.txtCredito.Name = "txtCredito";
			this.txtCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCredito.PromptChar = ' ';
			this.txtCredito.Size = new System.Drawing.Size(80, 22);
			this.txtCredito.TabIndex = 290;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(8, 258);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(98, 17);
			this.label10.TabIndex = 293;
			this.label10.Text = "Fecha Envio E-Mail";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaEnvio
			// 
			appearance8.FontData.Name = "Tahoma";
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEnvio.Appearance = appearance8;
			this.dtFechaEnvio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaEnvio.DateButtons.Add(dateButton3);
			this.dtFechaEnvio.Enabled = false;
			this.dtFechaEnvio.Format = "dd/MM/yyyy";
			this.dtFechaEnvio.Location = new System.Drawing.Point(128, 256);
			this.dtFechaEnvio.Name = "dtFechaEnvio";
			this.dtFechaEnvio.NonAutoSizeHeight = 23;
			this.dtFechaEnvio.Size = new System.Drawing.Size(112, 21);
			this.dtFechaEnvio.SpinButtonsVisible = true;
			this.dtFechaEnvio.TabIndex = 292;
			this.dtFechaEnvio.Value = ((object)(resources.GetObject("dtFechaEnvio.Value")));
			// 
			// chkEnvioEMail
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEnvioEMail.Appearance = appearance9;
			this.chkEnvioEMail.BackColor = System.Drawing.Color.Transparent;
			this.chkEnvioEMail.CausesValidation = false;
			this.chkEnvioEMail.Checked = true;
			this.chkEnvioEMail.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkEnvioEMail.Location = new System.Drawing.Point(8, 224);
			this.chkEnvioEMail.Name = "chkEnvioEMail";
			this.chkEnvioEMail.Size = new System.Drawing.Size(136, 22);
			this.chkEnvioEMail.TabIndex = 627;
			this.chkEnvioEMail.Text = "Enviar E-Mail al Grabar";
			this.chkEnvioEMail.CheckedChanged += new System.EventHandler(this.chkEnvioEMail_CheckedChanged);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(384, 131);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(33, 17);
			this.label11.TabIndex = 629;
			this.label11.Text = "Carpa";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(384, 258);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(66, 17);
			this.label12.TabIndex = 628;
			this.label12.Text = "Diario Carpa";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label12.Visible = false;
			// 
			// txtDiarioCarpa
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiarioCarpa.Appearance = appearance10;
			this.txtDiarioCarpa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiarioCarpa.Enabled = false;
			this.txtDiarioCarpa.FormatString = "#,##0.00";
			this.txtDiarioCarpa.Location = new System.Drawing.Point(472, 255);
			this.txtDiarioCarpa.Name = "txtDiarioCarpa";
			this.txtDiarioCarpa.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDiarioCarpa.PromptChar = ' ';
			this.txtDiarioCarpa.Size = new System.Drawing.Size(80, 22);
			this.txtDiarioCarpa.TabIndex = 631;
			this.txtDiarioCarpa.Visible = false;
			// 
			// txtMensualCarpa
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMensualCarpa.Appearance = appearance11;
			this.txtMensualCarpa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMensualCarpa.Enabled = false;
			this.txtMensualCarpa.FormatString = "#,##0";
			this.txtMensualCarpa.Location = new System.Drawing.Point(472, 128);
			this.txtMensualCarpa.Name = "txtMensualCarpa";
			this.txtMensualCarpa.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMensualCarpa.PromptChar = ' ';
			this.txtMensualCarpa.Size = new System.Drawing.Size(80, 22);
			this.txtMensualCarpa.TabIndex = 630;
			// 
			// chkGrupo
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGrupo.Appearance = appearance12;
			this.chkGrupo.BackColor = System.Drawing.Color.Transparent;
			this.chkGrupo.CausesValidation = false;
			this.chkGrupo.Enabled = false;
			this.chkGrupo.Location = new System.Drawing.Point(152, 224);
			this.chkGrupo.Name = "chkGrupo";
			this.chkGrupo.Size = new System.Drawing.Size(64, 22);
			this.chkGrupo.TabIndex = 632;
			this.chkGrupo.Text = "Puerteo";
			// 
			// txtIdUsuario
			// 
			this.txtIdUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdUsuario.Enabled = false;
			this.txtIdUsuario.Location = new System.Drawing.Point(432, 304);
			this.txtIdUsuario.MaskInput = "nn";
			this.txtIdUsuario.MinValue = 0;
			this.txtIdUsuario.Name = "txtIdUsuario";
			this.txtIdUsuario.PromptChar = ' ';
			this.txtIdUsuario.Size = new System.Drawing.Size(80, 22);
			this.txtIdUsuario.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtIdUsuario.TabIndex = 749;
			this.txtIdUsuario.Visible = false;
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.FillOnRequest = false;
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// frmCom_CambioVendedor
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(562, 336);
			this.ControlBox = false;
			this.Controls.Add(this.txtIdUsuario);
			this.Controls.Add(this.chkGrupo);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtDiarioCarpa);
			this.Controls.Add(this.txtMensualCarpa);
			this.Controls.Add(this.chkEnvioEMail);
			this.Controls.Add(this.dtFechaEnvio);
			this.Controls.Add(this.txtGEx);
			this.Controls.Add(this.txtCredito);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.cmbCargo);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtIdCom_AsignacionDeVendedoresN);
			this.Controls.Add(this.txtContado);
			this.Controls.Add(this.txtSemanal);
			this.Controls.Add(this.txtNumeroDias);
			this.Controls.Add(this.txtIdCom_AsignacionDeVendedores);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.cmbRango);
			this.Controls.Add(this.dtpHasta);
			this.Controls.Add(this.dtpDesde);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmCom_CambioVendedor";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Movimiento Vendedor";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCom_CambioVendedor_KeyDown);
			this.Load += new System.EventHandler(this.frmCom_CambioVendedor_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtpDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRango)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroDias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCom_AsignacionDeVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSemanal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCom_AsignacionDeVendedoresN)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGEx)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEnvio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiarioCarpa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMensualCarpa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}	
		
		private void frmCom_CambioVendedor_Load(object sender, System.EventArgs e)
		{
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 7, 0");						
			Puerteo(iBodega);//this.cmbCargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 59, 0");			
			this.cmbRango.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 60, 0");
			this.cmbEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 61, 0");
			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Com_ListaVendedoresMovimiento '{0}', {1}, {2}, 1", dtDesde.ToString("yyyyMMdd"), idPersonal, bNuevo));

			this.txtIdCom_AsignacionDeVendedores.Value = idCom_AsignaPersonal;							
			
			if (iNumero == 0)
				iNumero = Funcion.iEscalarSQL(cdsSeteoF, "Select Top 1 Numero From Com_AsignacionDeVendedores Order By Numero Desc");

			if (bNuevo)
			{
				this.cmbVendedor.Enabled = true;
				this.cmbEstado.Value = 3;
				this.cmbCargo.Value = 3;
				this.cmbRango.Value = 2;
				this.cmbBodega.Value = iBodega;

				this.dtpDesde.CalendarInfo.MinDate = dtDesde;
				this.dtpDesde.CalendarInfo.MaxDate = dtHasta;

				this.dtpHasta.CalendarInfo.MinDate = dtHasta;
				this.dtpHasta.CalendarInfo.MaxDate = dtHasta;

				this.dtpDesde.Value = DateTime.Today;
				this.dtpHasta.Value = dtHasta;
			}
			else
			{
				this.cmbVendedor.Value = idPersonal;				
				this.cmbEstado.Value = idEstado;
				this.cmbCargo.Value = idCargo;
				this.cmbRango.Value = idRango;	
				
				this.dtpDesde.CalendarInfo.MinDate = dtDesde;//DateTime.Today;//dtDesde;
				this.dtpDesde.CalendarInfo.MaxDate = dtHasta;
				this.dtpDesde.Value = DateTime.Today;
				
				this.dtpHasta.CalendarInfo.MinDate = dtHasta;
				this.dtpHasta.CalendarInfo.MaxDate = dtHasta;				
				this.dtpHasta.Value = dtHasta;				
			}			

			this.dtpDesde.Enabled = true;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			int Cargo = Convert.ToInt32(cmbCargo.Value);
			string nUsuario = "";
			string sClave = "";

			if (cmbVendedor.SelectedRow != null)
			{
				nUsuario = cmbVendedor.SelectedRow.Cells["Codigo"].Text;

				string sSQL = string.Format("SELECT Nidentificacion FROM Nomina WHERE Codigo = '{0}'", nUsuario);
				sClave = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
			}

			#region Validacion
			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione el Vendedor")) return;	

			if (!Validacion.vbFechaEnDocumentos(this.dtpDesde, "Seleccione la Fecha Inicial", false, "", cdsSeteoF)) return;
			
			if (!Validacion.vbFechaEnDocumentos(this.dtpHasta, "Seleccione la Fecha Final", false, "", cdsSeteoF)) return;

			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtpDesde.Value), Convert.ToDateTime(this.dtpHasta.Value), this.dtpDesde, "La Fecha Final no puede ser menor a la Fecha Inicial", false)) return;
			
			if (!Validacion.vbComboVacio(this.cmbEstado, "Seleccione el Estado")) return;
			
			if ((int)this.cmbEstado.Value == 2)
			{
				MessageBox.Show("El Estado no puede ser 'INACTIVO'", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbEstado.Focus();
				return;
			}

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione el Local")) return;
			
			if (!Validacion.vbComboVacio(this.cmbCargo, "Seleccione el Cargo")) return;

			if ((int)this.cmbCargo.Value != 1 && (int)this.cmbCargo.Value != 4 && (int)this.cmbCargo.Value != 7  && (int)this.cmbCargo.Value != 20)
				if (!Validacion.vbComboVacio(this.cmbRango, "Seleccione la Categoria")) return;

			if (iBodega == (int)this.cmbBodega.Value)
			{
				if (idCargo == (int)this.cmbCargo.Value)
				{
					MessageBox.Show("No puede generar un cambio con el mismo cargo en el mismo local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbBodega.Focus();
					return;
				}
			}

			if (!this.chkEnvioEMail.Checked)
				if (!Validacion.vbFechaEnDocumentos(this.dtFechaEnvio, "Seleccione la Fecha de Envio de E-Mail", false, "", cdsSeteoF)) return;			
			#endregion Validacion

			#region Guarda Movimiento
			int idRango = 0;
			if (this.cmbRango.ActiveRow != null) idRango = (int)this.cmbRango.Value;
			
			DateTime dtFechaEnvio = DateTime.Parse("01/01/2001");

			if (!this.chkEnvioEMail.Checked) dtFechaEnvio = Convert.ToDateTime(this.dtFechaEnvio.Value);

			if (!bNuevo && Convert.ToDateTime(this.dtpDesde.Value).Day == 1) this.txtIdCom_AsignacionDeVendedoresN.Value = (int)this.txtIdCom_AsignacionDeVendedores.Value;
			else
			{
				string sSQLA = string.Format("Exec Com_ActualizaMetasAVSEnMovimientos {0}, '{1}', {2}", 
					(int)this.txtIdCom_AsignacionDeVendedores.Value, Convert.ToDateTime(this.dtpDesde.Value).AddDays(-1).ToString("yyyyMMdd"), 
					DateTime.DaysInMonth(Convert.ToDateTime(this.dtpDesde.Value).Year, Convert.ToDateTime(this.dtpDesde.Value).Month));
				Funcion.EjecutaSQL(cdsSeteoF, sSQLA);
			}

			string sSQL5 = string.Format("Exec Com_GuardaAsignacionVendedores {0}, '{1}', {2}, {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, '{18}', {19}, {20}, {21}, {22}, {23}",
				(int)this.txtIdCom_AsignacionDeVendedoresN.Value, sPeriodo, (int)this.cmbBodega.Value, (int)this.cmbCargo.Value, (int)this.cmbVendedor.Value, idRango, //5
				Convert.ToDateTime(this.dtpDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtpHasta.Value).ToString("yyyyMMdd"), (int)this.txtNumeroDias.Value,  //7
				0.00m, 0.00m, 0.00m, 0.00m, 0.00m, Convert.ToDecimal(this.txtMensualCarpa.Value), 
				(int)this.cmbEstado.Value, 0, false, "20010101", (bool)this.chkGrupo.Checked, 
				Convert.ToDecimal(this.txtSemanal.Value), Convert.ToDecimal(this.txtContado.Value), Convert.ToDecimal(this.txtCredito.Value), Convert.ToDecimal(this.txtGEx.Value));
			this.txtIdCom_AsignacionDeVendedoresN.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL5);

			//			string sSQL = string.Format("Exec Com_GuardaAsignacionSemanalVendedores {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', '{9}', {10}, {11}, {12}, {13}, {14}, {15}, '{16}'",
			//				(int)this.txtIdCom_AsignacionDeVendedoresN.Value, iNumero, (int)this.cmbBodega.Value, (bool)this.chkGrupo.Checked, 
			//				(int)this.cmbVendedor.Value, (int)this.cmbCargo.Value, idRango, (int)this.cmbEstado.Value, 
			//				Convert.ToDateTime(this.dtpDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtpHasta.Value).ToString("yyyyMMdd"), (int)this.txtNumeroDias.Value, 
			//				Convert.ToDecimal(this.txtSemanal.Value), Convert.ToDecimal(this.txtContado.Value), Convert.ToDecimal(this.txtCredito.Value), Convert.ToDecimal(this.txtGEx.Value),				
			//				(bool)this.chkEnvioEMail.Checked, dtFechaEnvio.ToString("yyyyMMdd"));
			//			this.txtIdCom_AsignacionDeVendedoresN.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);	

			string sSQL6 = string.Format("Exec Com_ActualizaDatosNominaAV {0}, {1}, {2}, {3}, {4}", 
				(int)this.cmbBodega.Value, (int)this.cmbCargo.Value, (int)this.cmbEstado.Value, (int)this.cmbRango.Value, (int)this.cmbVendedor.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL6);
			#endregion Guarda Movimiento

			#region Mail
			if (this.chkEnvioEMail.Checked)
			{
				string sSQLMail = string.Format(@"Exec Com_MailIngresoPersonal {0}, '{1}', {2}, {3}", 
					(int)this.txtIdCom_AsignacionDeVendedoresN.Value, sPeriodo, (int)this.cmbBodega.Value, (int)this.cmbEstado.Value);
				string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLMail, true);
			}
			#endregion Mail

			#region Asigancion de bodega

			if (idUsuario > 0)
			{
				int nuevaBodega = Convert.ToInt32(cmbBodega.Value);
			
				string sSQLConsulta = string.Format("SELECT idUsuarioBodega FROM UsuariosBodegas WHERE idUsuario = {0}", idUsuario);
				int idUsuarioBodegaExistente = Funcion.iEscalarSQL(cdsSeteoF, sSQLConsulta);

				if (idUsuarioBodegaExistente > 0)
				{
					string sSQLEliminar = string.Format("EXEC GuardaUsuarioBodega {0}, {1}, {2}, 0",idUsuarioBodegaExistente, idUsuario, nuevaBodega);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLEliminar, true);
				}
				string sSQLInsertar = string.Format("EXEC GuardaUsuarioBodega 0, {0}, {1}, 1",idUsuario, nuevaBodega);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLInsertar, true);
			}

			#endregion Asignacion de bodega 


			if (bNuevo)
			{
				
				if ( Cargo == 1 || Cargo == 21 )
				{
					#region Validacion
					string sSQL = string.Format("Select Count(*) From Usuario Where Nombre = '{0}'", nUsuario);
					int iContUsuario = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
					if (iContUsuario > 0)
					{
						string sSQLGrupo = string.Format("SELECT IDUSUARIO FROM Usuario WHERE Nombre = '{0}' AND IDGRUPO = 36", nUsuario.Replace("'", "''"));
						int idUsuarioGrupo = Funcion.iEscalarSQL(cdsSeteoF, sSQLGrupo);

						if (idUsuarioGrupo > 0)
						{
							string sSQLUpdate = string.Format("UPDATE Usuario SET IDGRUPO = 22, ACTIVO = 1 WHERE IDUSUARIO = {0}", idUsuarioGrupo);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLUpdate, true);
						}
						else
						{
							MessageBox.Show(
								string.Format("El nombre de usuario '{0}' ya existe.", nUsuario),
								"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}
					}
					#endregion Validacion

					#region Crea usuario en la base de datos
					string stExec = string.Format("Exec UsuarioCrear {0}, '{1}', '{2}', {3}, {4}, {5}", 0, nUsuario, sClave, 22, 1, 1);
					this.txtIdUsuario.Value = Funcion.iEscalarSQL(cdsSeteo, stExec);
					#endregion Crea usuario en la base de datos

					#region Crea usuario en el servidor
					string stSelect = string.Format("SELECT Count(*) FROM Master.Dbo.syslogins Where LoginName = '{0}'", nUsuario);
					int iCuentaBases = Funcion.iEscalarSQL(cdsSeguridad, stSelect, true);
					if (iCuentaBases == 0)
					{
						string sSQL2 = string.Format("Create Login {0} With Password = '{1}'; Exec sp_addsrvrolemember '{0}', 'sysadmin'; Create User {0} For Login {0}", nUsuario, "Bl45o6$9");
						Funcion.EjecutaSQL(cdsSeguridad, sSQL2);
					}
					#endregion Crea usuario en el servidor			

					#region Asignacion de bodega
					string sSQL3 = string.Format("Exec GuardaUsuarioBodega {0}, {1}, {2}, {3}", 0, (int)this.txtIdUsuario.Value, iBodega, 1);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL3, true);
					#endregion Asignacion de bodega

					#region Permisos
					string sSQLUsuarios = string.Format("Exec PermisosUsuariosCreaUsuarios '{0}', {1}", nUsuario, 1);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLUsuarios, true);
					#endregion Permisos
					
					#region Mensaje de confirmación
					string sMensaje = "";
					sMensaje = "Usuario Registrado.";
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					#endregion Mensaje de confirmación
				}
				else if (Cargo == 3 || Cargo == 22  || Cargo == 23  || Cargo == 20)
				{
					#region Validacion
					string sSQL = string.Format("Select Count(*) From Usuario Where Nombre = '{0}'", nUsuario);
					int iContUsuario = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
					if (iContUsuario > 0)
					{
						string sSQLGrupo = string.Format("SELECT IDUSUARIO FROM Usuario WHERE Nombre = '{0}' AND IDGRUPO = 36", nUsuario.Replace("'", "''"));
						int idUsuarioGrupo = Funcion.iEscalarSQL(cdsSeteoF, sSQLGrupo);

						if (idUsuarioGrupo > 0)
						{
							string sSQLUpdate = string.Format("UPDATE Usuario SET IDGRUPO = 23, ACTIVO = 1 WHERE IDUSUARIO = {0}", idUsuarioGrupo);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLUpdate, true);
						}
						else
						{
							MessageBox.Show(
								string.Format("El nombre de usuario '{0}' ya existe.", nUsuario),
								"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}
					}
					#endregion Validacion

					#region Crea usuario en la base de datos
					string stExec = string.Format("Exec UsuarioCrear {0}, '{1}', '{2}', {3}, {4}, {5}", 0, nUsuario, sClave, 23, 1, 1);
					this.txtIdUsuario.Value = Funcion.iEscalarSQL(cdsSeteo, stExec);
					#endregion Crea usuario en la base de datos

					#region Crea usuario en el servidor
					string stSelect = string.Format("SELECT Count(*) FROM Master.Dbo.syslogins Where LoginName = '{0}'", nUsuario);
					int iCuentaBases = Funcion.iEscalarSQL(cdsSeguridad, stSelect, true);
					if (iCuentaBases == 0)
					{
						string sSQL2 = string.Format("Create Login {0} With Password = '{1}'; Exec sp_addsrvrolemember '{0}', 'sysadmin'; Create User {0} For Login {0}", nUsuario, "Bl45o6$9");
						Funcion.EjecutaSQL(cdsSeguridad, sSQL2);
					}
					#endregion Crea usuario en el servidor			

					#region Asignacion de bodega
					string sSQL3 = string.Format("Exec GuardaUsuarioBodega {0}, {1}, {2}, {3}", 0, (int)this.txtIdUsuario.Value, iBodega, 1);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL3, true);
					#endregion Asignacion de bodega

					#region Permisos
					string sSQLUsuarios = string.Format("Exec PermisosUsuariosCreaUsuarios '{0}', {1}", nUsuario, 2);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLUsuarios, true);
					#endregion Permisos
					
					#region Mensaje de confirmación
					string sMensaje = "";
					sMensaje = "Usuario Registrado.";
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					#endregion Mensaje de confirmación
				}
			}

			string sMsj = "Movimiento Guardado Correctamente";

			if (this.chkEnvioEMail.Checked) sMsj = sMsj + "\n\nE-Mail Enviado Correctamente";
			if (!this.chkEnvioEMail.Checked) sMsj = sMsj + string.Format("\n\nEnvio de E-Mail programado en la siguiente fecha '{0}'", Convert.ToDateTime(this.dtFechaEnvio.Value).ToString("dd/MM/yyyy"));

			MessageBox.Show(sMsj, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
			this.btnImprimir_Click(sender, e);

			this.DialogResult = DialogResult.OK;
		}
		
		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
				if (this.dtpHasta.Enabled) 
					this.dtpHasta.Focus(); else this.cmbBodega.Focus();
		}

		private void txtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCargo.Focus();
		}

		private void cmbCargo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void cmbRango_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void cmbEstado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtpDesde.Focus();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void frmCom_CambioVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void cmbCargo_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbCargo.ActiveRow == null) return;

			this.cmbRango.Value = idRango;

			if (bNuevo && (int)this.cmbCargo.Value != 5 && (int)this.cmbCargo.Value != 7) this.cmbRango.Value = 2;
			else if (bNuevo && (int)this.cmbCargo.Value == 5 && (int)this.cmbCargo.Value == 7 && (int)this.cmbCargo.Value == 20)	this.cmbRango.Value = System.DBNull.Value;
			else if (!bNuevo && (int)this.cmbCargo.Value == 5 && (int)this.cmbCargo.Value == 7 && (int)this.cmbCargo.Value == 20)	this.cmbRango.Value = System.DBNull.Value;			
			else
			{
				string sSQL = string.Format("Exec Com_ValidaCategoriaVendedores {0}, {1}, {2}", idCargo, (int)this.cmbCargo.Value, (int)this.cmbRango.Value);
				this.cmbRango.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			}

			CalculaMetasDias();
		}

		private void cmbEstado_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtpDesde.Focus();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			#region Parametros
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfId = new ParameterField ();
			ParameterDiscreteValue ValId = new ParameterDiscreteValue ();
			pfId.ParameterFieldName = "@idCom_AsignacionDeVendedores";
			ValId.Value = (int)this.txtIdCom_AsignacionDeVendedoresN.Value;
			pfId.CurrentValues.Add (ValId);
			paramFields.Add (pfId);
				
			ParameterField pfPeriodo = new ParameterField ();
			ParameterDiscreteValue ValPeriodo = new ParameterDiscreteValue ();
			pfPeriodo.ParameterFieldName = "@Periodo";
			ValPeriodo.Value = sPeriodo;
			pfPeriodo.CurrentValues.Add (ValPeriodo);
			paramFields.Add (pfPeriodo);
										
			ParameterField pfBodega = new ParameterField ();
			ParameterDiscreteValue ValBodega = new ParameterDiscreteValue ();
			pfBodega.ParameterFieldName = "@Bodega";
			ValBodega.Value = (int)this.cmbBodega.Value;
			pfBodega.CurrentValues.Add (ValBodega);
			paramFields.Add (pfBodega);
			#endregion Parametros

			Reporte Reporte = new Reporte("Com_ReporteIngresoDePersonal.rpt", "");			
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Págare");
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.ShowDialog();
		}

		private void CalculaNumeroDias()
		{
			if (this.dtpDesde.Value != System.DBNull.Value)
				if (this.dtpHasta.Value != System.DBNull.Value)
					this.txtNumeroDias.Value = FuncionesProcedimientos.CalculaDias((DateTime)this.dtpDesde.Value, (DateTime)this.dtpHasta.Value, cdsSeteoF);							
		}
		
		private void cmbEstado_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int)this.cmbEstado.Value > 3)
				this.cmbBodega.Value = iBodega;

			CalculaMetasDias();
		}

		private void CalculaMetasDias()
		{
			if (this.cmbEstado.ActiveRow != null && this.cmbCargo.ActiveRow != null && this.cmbBodega.ActiveRow != null && this.cmbVendedor.ActiveRow != null)
			{				
				this.txtSemanal.Value = 0.00m;
				this.txtContado.Value = 0.00m;
				this.txtCredito.Value = 0.00m;
				this.txtGEx.Value = 0.00m;
				this.txtMensualCarpa.Value = 0.00m;

				this.txtSemanal.Value = FuncionesProcedimientos.CalculaMeta((int)this.cmbBodega.Value, Convert.ToDateTime(this.dtpDesde.Value).Month, Convert.ToDateTime(this.dtpDesde.Value).Year, (int)this.cmbCargo.Value, (int)this.txtNumeroDias.Value, 1, cdsSeteoF);
				//this.txtContado.Value = FuncionesProcedimientos.CalculaMeta((int)this.cmbBodega.Value, Convert.ToDateTime(this.dtpDesde.Value).Month, Convert.ToDateTime(this.dtpDesde.Value).Year, (int)this.cmbCargo.Value, (int)this.txtNumeroDias.Value, 2, cdsSeteoF);
				this.txtCredito.Value = FuncionesProcedimientos.CalculaMeta((int)this.cmbBodega.Value, Convert.ToDateTime(this.dtpDesde.Value).Month, Convert.ToDateTime(this.dtpDesde.Value).Year, (int)this.cmbCargo.Value, (int)this.txtNumeroDias.Value, 3, cdsSeteoF);
				this.txtGEx.Value = FuncionesProcedimientos.CalculaMeta((int)this.cmbBodega.Value, Convert.ToDateTime(this.dtpDesde.Value).Month, Convert.ToDateTime(this.dtpDesde.Value).Year, (int)this.cmbCargo.Value, (int)this.txtNumeroDias.Value, 4, cdsSeteoF);
				this.txtMensualCarpa.Value = FuncionesProcedimientos.CalculaMeta((int)this.cmbBodega.Value, Convert.ToDateTime(this.dtpDesde.Value).Month, Convert.ToDateTime(this.dtpDesde.Value).Year, (int)this.cmbCargo.Value, (int)this.txtNumeroDias.Value, 5, cdsSeteoF);
								
//				int iDiasTrabaja = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(DAY, dbo.NominaRetornaFechaEntradaPersonal({0}, 0, 0), '{1}')", 
//					(int)this.cmbVendedor.Value, Convert.ToDateTime(this.dtpHasta.Value).ToString("yyyyMMdd")));
//
//				if (iDiasTrabaja > 30)					
//					this.txtMensualCarpa.Value = FuncionesProcedimientos.CalculaMetaCarpa((int)this.cmbEstado.Value, sPeriodo, 
//						(int)this.cmbBodega.Value, (int)this.cmbCargo.Value, 
//						(DateTime)this.dtpDesde.Value, (int)this.txtNumeroDias.Value, cdsSeteoF);
//				else this.txtMensualCarpa.Value = 0.00m;
//
//
//				if ((int)this.txtNumeroDias.Value > 0) 
//				{
//					this.txtDiario.Value = Convert.ToDecimal(this.txtMensual.Value)/(int)this.txtNumeroDias.Value;
//					this.txtDiarioGEX.Value = Convert.ToDecimal(this.txtMensualGEX.Value)/(int)this.txtNumeroDias.Value;
//
//					if (iDiasTrabaja > 30)
//						this.txtDiarioCarpa.Value = Math.Round(Convert.ToDecimal(this.txtMensualCarpa.Value)/(int)this.txtNumeroDias.Value, 2);
//					else this.txtDiarioCarpa.Value = 0.00m;
//				}
			}
		}
		
		private void Puerteo(int iAlmcacen)
		{
			bool bPuerteo = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Puerteo From Bodega Where Bodega = {0}", iAlmcacen), true);
 
			int iPuerteo = 0;
			if (bPuerteo)
			{
				this.chkGrupo.Enabled = true;
				iPuerteo = 1;
			}

			this.cmbCargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 59, {0}, 0", iPuerteo));
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			CalculaMetasDias();

			this.chkGrupo.Enabled = false;
			this.chkGrupo.Checked = false;

			Puerteo((int)this.cmbBodega.Value);
		}

		private void dtpHasta_ValueChanged(object sender, System.EventArgs e)
		{
			CalculaNumeroDias();

			CalculaMetasDias();
		}

		private void dtpDesde_ValueChanged(object sender, System.EventArgs e)
		{
			CalculaNumeroDias();

			CalculaMetasDias();
		}

		private void cmbVendedor_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbVendedor.ActiveRow != null)
			{
				string sSQL = string.Format("Exec NominaConsultaFechaEntrada {0}", (int)this.cmbVendedor.Value);
				DateTime dtFecEntrada = Funcion.dtEscalarSQL(cdsSeteoF, sSQL);
			
				if (dtFecEntrada < dtDesde)
				{
					this.dtpDesde.CalendarInfo.MinDate = dtDesde;
					this.dtpDesde.Value = dtDesde;
				}
				else
				{
					this.dtpDesde.CalendarInfo.MinDate = dtFecEntrada;
					this.dtpDesde.Value = dtFecEntrada;
				}

				CalculaNumeroDias();
				CalculaMetasDias();
			}
		}

		private void chkEnvioEMail_CheckedChanged(object sender, System.EventArgs e)
		{				
			if (!this.chkEnvioEMail.Checked)
			{
				this.dtFechaEnvio.CalendarInfo.MinDate = DateTime.Today;
				this.dtFechaEnvio.CalendarInfo.MaxDate = (DateTime)this.dtpDesde.Value;
				this.dtFechaEnvio.Value = (DateTime)this.dtpDesde.Value;
				this.dtFechaEnvio.Focus();
			}

			if (this.chkEnvioEMail.Checked) this.dtFechaEnvio.Value = System.DBNull.Value;
			this.dtFechaEnvio.Enabled = !this.chkEnvioEMail.Checked;
		}

		private void cmbCargo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbVendedor_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}
	}
}
