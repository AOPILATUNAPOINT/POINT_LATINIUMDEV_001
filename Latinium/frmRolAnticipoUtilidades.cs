using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRolAnticipoUtilidades.
	/// </summary>
	public class frmRolAnticipoUtilidades : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnCerrar;
		private System.Windows.Forms.Button btnSeleccionar;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnAcreditacion;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEDepartamento;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpPeriodo;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridComisiones;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECentroCosto;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ContextMenu mnuAcreditacion;
		private System.Windows.Forms.MenuItem mnuAcreditacionTransferencias;
		private System.Windows.Forms.MenuItem mnuAcreditacionCheques;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAntipoUtilidades;

		int Estado = 0;
		int idAntipoUtilidades =0;
		int tipoPantalla =0;

		public frmRolAnticipoUtilidades(int IdAntipoUtilidades ,int TipoPantalla )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			idAntipoUtilidades =IdAntipoUtilidades;
			tipoPantalla = TipoPantalla;
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmRolAnticipoUtilidades()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRolAnticipoUtilidades));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolDetalleAnticipoUtilidades");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolAnticipoUtilidades");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellidos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Salida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDePago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Centro_De_Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valorQuincenal");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valorMensual");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("total");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnFC", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnOtros", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnticipos", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMultas", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCredito", 4);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "total", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "total", 19, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valorMensual", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valorMensual", 18, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "valorQuincenal", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "valorQuincenal", 17, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolDetalleAnticipoUtilidades");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolAnticipoUtilidades");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Apellidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Salida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDePago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Centro_De_Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Departamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("valorQuincenal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("valorMensual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("total");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.btnCerrar = new System.Windows.Forms.Button();
			this.btnSeleccionar = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnAcreditacion = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbNEDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label33 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtIdAntipoUtilidades = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.dtpPeriodo = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridComisiones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExportar = new System.Windows.Forms.Button();
			this.cmbNECentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.mnuAcreditacion = new System.Windows.Forms.ContextMenu();
			this.mnuAcreditacionTransferencias = new System.Windows.Forms.MenuItem();
			this.mnuAcreditacionCheques = new System.Windows.Forms.MenuItem();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAntipoUtilidades)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpPeriodo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridComisiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			this.SuspendLayout();
			// 
			// btnCerrar
			// 
			this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCerrar.CausesValidation = false;
			this.btnCerrar.Image = ((System.Drawing.Image)(resources.GetObject("btnCerrar.Image")));
			this.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCerrar.Location = new System.Drawing.Point(1073, 3);
			this.btnCerrar.Name = "btnCerrar";
			this.btnCerrar.TabIndex = 768;
			this.btnCerrar.Text = "Cerrar";
			this.btnCerrar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
			// 
			// btnSeleccionar
			// 
			this.btnSeleccionar.CausesValidation = false;
			this.btnSeleccionar.Enabled = false;
			this.btnSeleccionar.Image = ((System.Drawing.Image)(resources.GetObject("btnSeleccionar.Image")));
			this.btnSeleccionar.Location = new System.Drawing.Point(8, 96);
			this.btnSeleccionar.Name = "btnSeleccionar";
			this.btnSeleccionar.Size = new System.Drawing.Size(30, 22);
			this.btnSeleccionar.TabIndex = 767;
			this.btnSeleccionar.Visible = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(654, 67);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 16);
			this.lblContador.TabIndex = 765;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(422, 6);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 16);
			this.lblEstado.TabIndex = 764;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnAcreditacion
			// 
			this.btnAcreditacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAcreditacion.CausesValidation = false;
			this.btnAcreditacion.Image = ((System.Drawing.Image)(resources.GetObject("btnAcreditacion.Image")));
			this.btnAcreditacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAcreditacion.Location = new System.Drawing.Point(896, 3);
			this.btnAcreditacion.Name = "btnAcreditacion";
			this.btnAcreditacion.Size = new System.Drawing.Size(88, 23);
			this.btnAcreditacion.TabIndex = 763;
			this.btnAcreditacion.Text = "Acreditación";
			this.btnAcreditacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAcreditacion.Click += new System.EventHandler(this.btnAcreditacion_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(6, 99);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1232, 8);
			this.groupBox1.TabIndex = 762;
			this.groupBox1.TabStop = false;
			// 
			// cmbNEDepartamento
			// 
			this.cmbNEDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridBand1.AddButtonCaption = "RolTipo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 208;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbNEDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbNEDepartamento.DisplayMember = "Descripcion";
			this.cmbNEDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEDepartamento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEDepartamento.Location = new System.Drawing.Point(422, 35);
			this.cmbNEDepartamento.Name = "cmbNEDepartamento";
			this.cmbNEDepartamento.Size = new System.Drawing.Size(200, 22);
			this.cmbNEDepartamento.TabIndex = 761;
			this.cmbNEDepartamento.ValueMember = "idDepartamento";
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(14, 37);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(88, 16);
			this.label33.TabIndex = 758;
			this.label33.Text = "Centro De Costo";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(14, 68);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(84, 16);
			this.label7.TabIndex = 754;
			this.label7.Text = "Cedula/Nombre";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(14, 6);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(43, 16);
			this.label6.TabIndex = 749;
			this.label6.Text = "Periodo";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(334, 37);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(76, 16);
			this.label12.TabIndex = 760;
			this.label12.Text = "Departamento";
			// 
			// txtIdAntipoUtilidades
			// 
			this.txtIdAntipoUtilidades.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAntipoUtilidades.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAntipoUtilidades.Enabled = false;
			this.txtIdAntipoUtilidades.Location = new System.Drawing.Point(611, 4);
			this.txtIdAntipoUtilidades.Name = "txtIdAntipoUtilidades";
			this.txtIdAntipoUtilidades.PromptChar = ' ';
			this.txtIdAntipoUtilidades.Size = new System.Drawing.Size(8, 21);
			this.txtIdAntipoUtilidades.TabIndex = 756;
			this.txtIdAntipoUtilidades.Visible = false;
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(110, 67);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(200, 19);
			this.txtBuscar.TabIndex = 753;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(334, 65);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(64, 23);
			this.btnVer.TabIndex = 752;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1155, 3);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 751;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// dtpPeriodo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpPeriodo.Appearance = appearance1;
			this.dtpPeriodo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpPeriodo.DateButtons.Add(dateButton1);
			this.dtpPeriodo.Format = "MM/yyyy";
			this.dtpPeriodo.Location = new System.Drawing.Point(110, 4);
			this.dtpPeriodo.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpPeriodo.Name = "dtpPeriodo";
			this.dtpPeriodo.NonAutoSizeHeight = 23;
			this.dtpPeriodo.Size = new System.Drawing.Size(104, 21);
			this.dtpPeriodo.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtpPeriodo.TabIndex = 750;
			this.dtpPeriodo.Value = ((object)(resources.GetObject("dtpPeriodo.Value")));
			this.dtpPeriodo.ValueChanged += new System.EventHandler(this.dtpPeriodo_ValueChanged);
			// 
			// uGridComisiones
			// 
			this.uGridComisiones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridComisiones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridComisiones.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridComisiones.DisplayLayout.Appearance = appearance2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Header.Caption = "N°";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 30;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 150;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 150;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 70;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Width = 70;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Width = 120;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.VisiblePosition = 13;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.Caption = "Centro De Costo";
			ultraGridColumn17.Header.VisiblePosition = 14;
			ultraGridColumn17.Width = 200;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Format = "";
			ultraGridColumn18.Header.VisiblePosition = 16;
			ultraGridColumn18.Width = 150;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn19.Header.VisiblePosition = 15;
			ultraGridColumn19.Width = 150;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance4;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "Valor Quincenal";
			ultraGridColumn20.Header.VisiblePosition = 17;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance5;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Valor Mensual";
			ultraGridColumn21.Header.VisiblePosition = 18;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance6;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "Total";
			ultraGridColumn22.Header.VisiblePosition = 19;
			ultraGridColumn23.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn23.Header.Caption = "F C";
			ultraGridColumn23.Header.VisiblePosition = 21;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn23.Width = 30;
			ultraGridColumn24.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn24.Header.Caption = "O";
			ultraGridColumn24.Header.VisiblePosition = 24;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn24.Width = 30;
			ultraGridColumn25.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn25.Header.Caption = "A";
			ultraGridColumn25.Header.VisiblePosition = 23;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn25.Width = 30;
			ultraGridColumn26.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn26.Header.Caption = "M";
			ultraGridColumn26.Header.VisiblePosition = 22;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn26.Width = 30;
			ultraGridColumn27.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn27.Header.Caption = "C";
			ultraGridColumn27.Header.VisiblePosition = 20;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn27.Width = 30;
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
																										 ultraGridColumn27});
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance7;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance8;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance9;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand2.SummaryFooterCaption = "Totales";
			this.uGridComisiones.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridComisiones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridComisiones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridComisiones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridComisiones.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridComisiones.DisplayLayout.Override.CellAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8.5F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridComisiones.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.uGridComisiones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridComisiones.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridComisiones.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridComisiones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridComisiones.Location = new System.Drawing.Point(14, 115);
			this.uGridComisiones.Name = "uGridComisiones";
			this.uGridComisiones.Size = new System.Drawing.Size(1216, 240);
			this.uGridComisiones.TabIndex = 748;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn12.DataType = typeof(int);
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
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(991, 3);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.TabIndex = 757;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// cmbNECentroCosto
			// 
			this.cmbNECentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridBand3.AddButtonCaption = "RolTipo";
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 85;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Width = 208;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			this.cmbNECentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbNECentroCosto.DisplayMember = "Nombre";
			this.cmbNECentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECentroCosto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECentroCosto.Location = new System.Drawing.Point(110, 35);
			this.cmbNECentroCosto.Name = "cmbNECentroCosto";
			this.cmbNECentroCosto.Size = new System.Drawing.Size(200, 22);
			this.cmbNECentroCosto.TabIndex = 759;
			this.cmbNECentroCosto.ValueMember = "idProyecto";
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
			// mnuAcreditacion
			// 
			this.mnuAcreditacion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																										this.mnuAcreditacionTransferencias,
																																										this.mnuAcreditacionCheques});
			// 
			// mnuAcreditacionTransferencias
			// 
			this.mnuAcreditacionTransferencias.Index = 0;
			this.mnuAcreditacionTransferencias.Text = "Transferencias";
			this.mnuAcreditacionTransferencias.Click += new System.EventHandler(this.mnuAcreditacionTransferencias_Click);
			// 
			// mnuAcreditacionCheques
			// 
			this.mnuAcreditacionCheques.Index = 1;
			this.mnuAcreditacionCheques.Text = "Cheques";
			this.mnuAcreditacionCheques.Click += new System.EventHandler(this.mnuAcreditacionCheques_Click);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// ultraDataSource7
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// frmRolAnticipoUtilidades
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1245, 358);
			this.Controls.Add(this.btnCerrar);
			this.Controls.Add(this.btnSeleccionar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnAcreditacion);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbNEDepartamento);
			this.Controls.Add(this.label33);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtIdAntipoUtilidades);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.dtpPeriodo);
			this.Controls.Add(this.uGridComisiones);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.cmbNECentroCosto);
			this.Name = "frmRolAnticipoUtilidades";
			this.Text = "Rol Anticipo Utilidades";
			this.Load += new System.EventHandler(this.frmRolAnticipoUtilidades_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAntipoUtilidades)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpPeriodo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridComisiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRolAnticipoUtilidades_Load(object sender, System.EventArgs e)
		{
			DateTime dtFecha = DateTime.Parse("01/01/2024");			
			this.dtpPeriodo.CalendarInfo.MinDate = dtFecha.Date;
			this.dtpPeriodo.CalendarInfo.MaxDate = Convert.ToDateTime(DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month).ToString() + "/" + DateTime.Today.Month.ToString() + "/" + DateTime.Today.Year.ToString());

			this.cmbNECentroCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 45, 0");
			this.cmbNEDepartamento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 48, 0");

			this.dtpPeriodo.Value = DateTime.Today.AddDays(-31);

			if (idAntipoUtilidades > 0)
			{
				//this.dtpPeriodo.Value = DateTime.Today.AddDays(-31);
				dtpPeriodo_ValueChanged(sender, e);
				this.txtIdAntipoUtilidades.Value = idAntipoUtilidades;
				this.btnVer_Click(sender, e);
			}

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void dtpPeriodo_ValueChanged(object sender, System.EventArgs e)
		{	
			this.btnAcreditacion.Enabled = false;
			this.btnExportar.Enabled = false;			
			this.btnSeleccionar.Enabled = false;
			this.btnCerrar.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(false, this.uGridComisiones);
			
			string sSQL = string.Format("Exec Com_ValidaComisionesExisteUtilidades '{0}'", Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"));			
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdAntipoUtilidades.Value = dr.GetInt32(0);
				this.lblEstado.Text = dr.GetString(1);
				Estado = dr.GetInt32(2);	
				dr.Close();

				if (Estado == 1)
				{
					this.btnAcreditacion.Enabled = true;
					this.btnExportar.Enabled = true;
					this.btnSeleccionar.Enabled = true;
				
					FuncionesProcedimientos.EstadoGrids(true, this.uGridComisiones);
				}
				else if (Estado == 2) 
				{
					this.btnCerrar.Enabled = true;
					this.btnExportar.Enabled = true;
				}
				else if (Estado == 3)
				{
					this.btnExportar.Enabled = true;
				}
				if (Estado == 4)
				{
					this.btnAcreditacion.Enabled = true;
					this.btnExportar.Enabled = true;
					this.btnSeleccionar.Enabled = true;
				
					FuncionesProcedimientos.EstadoGrids(true, this.uGridComisiones);
				}
				else if (Estado == 5) 
				{
					this.btnCerrar.Enabled = true;
					this.btnExportar.Enabled = true;
				}
				else if (Estado == 6)
				{
					this.btnExportar.Enabled = true;
				}
				else this.btnExportar.Enabled = true;	
							
			}
			else
			{
				dr.Close();				
				this.txtIdAntipoUtilidades.Value = 0;
				this.lblEstado.Text = "PENDIENTE DE GENERAR";

				FuncionesProcedimientos.EstadoGrids(true, this.uGridComisiones);
			
			}
			dr.Close();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int idProyecto = 0;
			int idDepartamento = 0;

			if (this.cmbNECentroCosto.ActiveRow != null) idProyecto = (int)this.cmbNECentroCosto.Value;
			if (this.cmbNEDepartamento.ActiveRow != null) idDepartamento = (int)this.cmbNEDepartamento.Value;

			string sSQL = string.Format("Exec Com_ConsultaDetalleaAnticipoUtilidades {0}, {1}, {2}, '{3}'", (int)this.txtIdAntipoUtilidades.Value, idProyecto, idDepartamento, this.txtBuscar.Text.ToString());
			this.uGridComisiones.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			this.lblContador.Text = this.uGridComisiones.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnAcreditacion_Click(object sender, System.EventArgs e)
		{
			this.mnuAcreditacion.Show(this.btnAcreditacion, new Point(1, 25));
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridComisiones);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void mnuAcreditacionTransferencias_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());

			using (frmRolAcreditacion RAC = new frmRolAcreditacion((int)this.txtIdAntipoUtilidades.Value, 0, dtFecha, tipoPantalla))
			{
				if (DialogResult.OK == RAC.ShowDialog())
				{
		
				}

				this.dtpPeriodo_ValueChanged(sender, e);
			}	
		}

		private void mnuAcreditacionCheques_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());

			using (frmRolAutorizacionPagoCheques RAPC = new frmRolAutorizacionPagoCheques((int)this.txtIdAntipoUtilidades.Value, dtFecha, tipoPantalla))
			{
				if (DialogResult.OK == RAPC.ShowDialog())
				{
		
				}

				this.dtpPeriodo_ValueChanged(sender, e);
			}
		}

		private void btnCerrar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show("Esta seguro de cerrar.", "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaActualizaAnticipoUtilidades {0}",  (int)this.txtIdAntipoUtilidades.Value));
	
				dtpPeriodo_ValueChanged(sender, e);
			}
		}
	}
}
