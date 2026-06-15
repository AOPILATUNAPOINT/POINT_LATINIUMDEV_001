using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmrequerimientoreporte.
	/// </summary>
	public class frmrequerimientoreporte : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label lblRegistros;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblAsignacionesMias;
		public System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.CheckBox chkAsignacionesMias;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnConsultar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.GroupBox groupBox4;
		public Infragistics.Win.UltraWinGrid.UltraGrid udgSeguimiento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimiento;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsResponsable;
		private System.Windows.Forms.Button btnSalir;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoReporteHistorial;
		public Infragistics.Win.UltraWinGrid.UltraGrid udgRequerimientoReporteHistorial;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optReporte;
		private System.Windows.Forms.TabControl tabReporte;
		int idRequerimientoDepartamento = 0;

		public frmrequerimientoreporte(int IdRequerimientoDepartamento)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			idRequerimientoDepartamento = IdRequerimientoDepartamento;
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmrequerimientoreporte));
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Titulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UltimaNota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPrueba");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAprobacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaResuelto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCerrado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeguimientoAnadir", 0);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 2);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Titulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UltimaNota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPrueba");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAprobacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaResuelto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCerrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Titulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UltimaNota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeguimientoAnadir", 0);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEstado", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 2);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Titulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UltimaNota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label12 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.optReporte = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.button1 = new System.Windows.Forms.Button();
			this.lblAsignacionesMias = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.chkAsignacionesMias = new System.Windows.Forms.CheckBox();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.lblRegistros = new System.Windows.Forms.Label();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.tabReporte = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.udgSeguimiento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimiento = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.udgRequerimientoReporteHistorial = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsRequerimientoReporteHistorial = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btnSalir = new System.Windows.Forms.Button();
			this.udsResponsable = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.groupBox4.SuspendLayout();
			this.tabReporte.SuspendLayout();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgSeguimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimiento)).BeginInit();
			this.tabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoReporteHistorial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoReporteHistorial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Controls.Add(this.label1);
			this.groupBox3.Controls.Add(this.cmbResponsable);
			this.groupBox3.Controls.Add(this.txtNumero);
			this.groupBox3.Controls.Add(this.cmbEstado);
			this.groupBox3.Controls.Add(this.label12);
			this.groupBox3.Controls.Add(this.lblEstado);
			this.groupBox3.Location = new System.Drawing.Point(10, 18);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1420, 65);
			this.groupBox3.TabIndex = 1052;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Datos principales";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(10, 28);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(60, 18);
			this.label1.TabIndex = 1038;
			this.label1.Text = "* Usuario";
			// 
			// cmbResponsable
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance1;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(125, 28);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.ReadOnly = true;
			this.cmbResponsable.Size = new System.Drawing.Size(297, 25);
			this.cmbResponsable.TabIndex = 1042;
			this.cmbResponsable.ValueMember = "idPersonal";
			// 
			// txtNumero
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance2;
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(614, 28);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(144, 24);
			this.txtNumero.TabIndex = 1050;
			this.txtNumero.Visible = false;
			// 
			// cmbEstado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance3;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 217;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(864, 28);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(182, 25);
			this.cmbEstado.TabIndex = 1055;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			this.cmbEstado.Visible = false;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(538, 28);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(47, 18);
			this.label12.TabIndex = 1051;
			this.label12.Text = "Buscar";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label12.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Location = new System.Drawing.Point(778, 28);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(124, 26);
			this.lblEstado.TabIndex = 1054;
			this.lblEstado.Text = "* Estado:";
			this.lblEstado.Visible = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.optReporte);
			this.groupBox2.Controls.Add(this.button1);
			this.groupBox2.Controls.Add(this.lblAsignacionesMias);
			this.groupBox2.Controls.Add(this.label17);
			this.groupBox2.Controls.Add(this.label18);
			this.groupBox2.Controls.Add(this.chkAsignacionesMias);
			this.groupBox2.Controls.Add(this.dtHasta);
			this.groupBox2.Controls.Add(this.dtDesde);
			this.groupBox2.Controls.Add(this.btnConsultar);
			this.groupBox2.Controls.Add(this.lblRegistros);
			this.groupBox2.Location = new System.Drawing.Point(10, 102);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1420, 73);
			this.groupBox2.TabIndex = 1051;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Búsqueda";
			// 
			// optReporte
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.optReporte.Appearance = appearance4;
			this.optReporte.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optReporte.CheckedIndex = 0;
			this.optReporte.ItemAppearance = appearance5;
			this.optReporte.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "General";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Detalle";
			this.optReporte.Items.Add(valueListItem1);
			this.optReporte.Items.Add(valueListItem2);
			this.optReporte.ItemSpacingVertical = 10;
			this.optReporte.Location = new System.Drawing.Point(568, 24);
			this.optReporte.Name = "optReporte";
			this.optReporte.Size = new System.Drawing.Size(168, 29);
			this.optReporte.TabIndex = 1060;
			this.optReporte.Text = "General";
			this.optReporte.ValueChanged += new System.EventHandler(this.optRegiones_ValueChanged);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.CausesValidation = false;
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(1306, 28);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(93, 27);
			this.button1.TabIndex = 1057;
			this.button1.Text = "&Exportar";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// lblAsignacionesMias
			// 
			this.lblAsignacionesMias.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblAsignacionesMias.ForeColor = System.Drawing.Color.Red;
			this.lblAsignacionesMias.Location = new System.Drawing.Point(1085, 120);
			this.lblAsignacionesMias.Name = "lblAsignacionesMias";
			this.lblAsignacionesMias.Size = new System.Drawing.Size(230, 27);
			this.lblAsignacionesMias.TabIndex = 1053;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(278, 28);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(44, 18);
			this.label17.TabIndex = 1047;
			this.label17.Text = "Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(19, 28);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(47, 18);
			this.label18.TabIndex = 1046;
			this.label18.Text = "Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkAsignacionesMias
			// 
			this.chkAsignacionesMias.Location = new System.Drawing.Point(970, 120);
			this.chkAsignacionesMias.Name = "chkAsignacionesMias";
			this.chkAsignacionesMias.Size = new System.Drawing.Size(96, 28);
			this.chkAsignacionesMias.TabIndex = 1052;
			this.chkAsignacionesMias.Text = "Asignaciones mías";
			this.chkAsignacionesMias.Visible = false;
			// 
			// dtHasta
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance6;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(346, 28);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(192, 24);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1045;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance7;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(77, 28);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(182, 24);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1044;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(1200, 28);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(94, 27);
			this.btnConsultar.TabIndex = 1019;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// lblRegistros
			// 
			this.lblRegistros.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRegistros.ForeColor = System.Drawing.Color.Red;
			this.lblRegistros.Location = new System.Drawing.Point(760, 24);
			this.lblRegistros.Name = "lblRegistros";
			this.lblRegistros.Size = new System.Drawing.Size(212, 26);
			this.lblRegistros.TabIndex = 1056;
			// 
			// groupBox4
			// 
			this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox4.Controls.Add(this.tabReporte);
			this.groupBox4.Location = new System.Drawing.Point(10, 194);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1555, 406);
			this.groupBox4.TabIndex = 1053;
			this.groupBox4.TabStop = false;
			// 
			// tabReporte
			// 
			this.tabReporte.Controls.Add(this.tabPage1);
			this.tabReporte.Controls.Add(this.tabPage2);
			this.tabReporte.Location = new System.Drawing.Point(16, 24);
			this.tabReporte.Name = "tabReporte";
			this.tabReporte.SelectedIndex = 0;
			this.tabReporte.Size = new System.Drawing.Size(1512, 368);
			this.tabReporte.TabIndex = 13;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.udgSeguimiento);
			this.tabPage1.Location = new System.Drawing.Point(4, 25);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(1504, 339);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "General";
			// 
			// udgSeguimiento
			// 
			this.udgSeguimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgSeguimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgSeguimiento.DataSource = this.udsRequerimiento;
			appearance8.BackColor = System.Drawing.Color.White;
			this.udgSeguimiento.DisplayLayout.Appearance = appearance8;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Número";
			ultraGridColumn5.Header.Fixed = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 135;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Título";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 317;
			ultraGridColumn7.AutoEdit = false;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Descripción";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 425;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Width = 227;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 123;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Última nota";
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Width = 443;
			ultraGridColumn11.AutoEdit = false;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 114;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Width = 110;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn14.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn14.Header.Caption = "Fecha prueba";
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.Width = 135;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn15.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn15.Header.Caption = "Fecha aprobación";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 130;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn16.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn16.Header.Caption = "Fecha resuelto";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 118;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn17.Format = "dd/MM/yyyy HH:mm:ss";
			ultraGridColumn17.Header.Caption = "Fecha cerrado";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 107;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Width = 124;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			ultraGridColumn19.CellAppearance = appearance9;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			ultraGridColumn19.CellButtonAppearance = appearance10;
			ultraGridColumn19.Header.Caption = "...";
			ultraGridColumn19.Header.Fixed = true;
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 22;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.Caption = "Estado";
			ultraGridColumn20.Header.VisiblePosition = 13;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn20.Width = 122;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn21.Header.Caption = "N°";
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Width = 29;
			ultraGridBand3.Columns.AddRange(new object[] {
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
			ultraGridBand3.Header.Caption = "Detalle";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgSeguimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.udgSeguimiento.DisplayLayout.GroupByBox.Hidden = true;
			this.udgSeguimiento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgSeguimiento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgSeguimiento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.udgSeguimiento.DisplayLayout.Override.CardAreaAppearance = appearance11;
			this.udgSeguimiento.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 10F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgSeguimiento.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgSeguimiento.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgSeguimiento.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.udgSeguimiento.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgSeguimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.udgSeguimiento.Location = new System.Drawing.Point(16, 16);
			this.udgSeguimiento.Name = "udgSeguimiento";
			this.udgSeguimiento.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.udgSeguimiento.Size = new System.Drawing.Size(1472, 304);
			this.udgSeguimiento.TabIndex = 12;
			// 
			// udsRequerimiento
			// 
			ultraDataColumn2.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn3.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.ReadOnly = Infragistics.Win.DefaultableBoolean.True;
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(System.DateTime);
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn13.DataType = typeof(System.DateTime);
			this.udsRequerimiento.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn14});
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.udgRequerimientoReporteHistorial);
			this.tabPage2.Location = new System.Drawing.Point(4, 25);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Size = new System.Drawing.Size(1504, 339);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "Detalle";
			// 
			// udgRequerimientoReporteHistorial
			// 
			this.udgRequerimientoReporteHistorial.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgRequerimientoReporteHistorial.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgRequerimientoReporteHistorial.DataSource = this.udsRequerimientoReporteHistorial;
			appearance15.BackColor = System.Drawing.Color.White;
			this.udgRequerimientoReporteHistorial.DisplayLayout.Appearance = appearance15;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn27.Header.VisiblePosition = 6;
			ultraGridColumn28.Header.Caption = "Nota";
			ultraGridColumn28.Header.VisiblePosition = 7;
			ultraGridColumn28.Width = 443;
			ultraGridColumn29.Header.VisiblePosition = 8;
			ultraGridColumn30.Header.VisiblePosition = 9;
			ultraGridColumn31.Header.VisiblePosition = 10;
			appearance16.Image = ((object)(resources.GetObject("appearance16.Image")));
			ultraGridColumn32.CellAppearance = appearance16;
			appearance17.Image = ((object)(resources.GetObject("appearance17.Image")));
			ultraGridColumn32.CellButtonAppearance = appearance17;
			ultraGridColumn32.Header.Caption = "...";
			ultraGridColumn32.Header.Fixed = true;
			ultraGridColumn32.Header.VisiblePosition = 12;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn32.Width = 22;
			ultraGridColumn33.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.Caption = "Estado";
			ultraGridColumn33.Header.VisiblePosition = 11;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn33.Width = 122;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn34.Header.Caption = "N°";
			ultraGridColumn34.Header.VisiblePosition = 0;
			ultraGridColumn34.Width = 29;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			ultraGridBand4.Header.Caption = "Detalle";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand4.SummaryFooterCaption = "Sum Detalle:";
			this.udgRequerimientoReporteHistorial.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.udgRequerimientoReporteHistorial.DisplayLayout.GroupByBox.Hidden = true;
			this.udgRequerimientoReporteHistorial.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udgRequerimientoReporteHistorial.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udgRequerimientoReporteHistorial.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.udgRequerimientoReporteHistorial.DisplayLayout.Override.CardAreaAppearance = appearance18;
			this.udgRequerimientoReporteHistorial.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 10F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgRequerimientoReporteHistorial.DisplayLayout.Override.HeaderAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoReporteHistorial.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgRequerimientoReporteHistorial.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.udgRequerimientoReporteHistorial.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgRequerimientoReporteHistorial.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.udgRequerimientoReporteHistorial.Location = new System.Drawing.Point(16, 16);
			this.udgRequerimientoReporteHistorial.Name = "udgRequerimientoReporteHistorial";
			this.udgRequerimientoReporteHistorial.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.udgRequerimientoReporteHistorial.Size = new System.Drawing.Size(1472, 304);
			this.udgRequerimientoReporteHistorial.TabIndex = 1062;
			// 
			// udsRequerimientoReporteHistorial
			// 
			ultraDataColumn15.DataType = typeof(System.UInt32);
			ultraDataColumn20.DataType = typeof(System.DateTime);
			this.udsRequerimientoReporteHistorial.Band.Columns.AddRange(new object[] {
																																								 ultraDataColumn15,
																																								 ultraDataColumn16,
																																								 ultraDataColumn17,
																																								 ultraDataColumn18,
																																								 ultraDataColumn19,
																																								 ultraDataColumn20,
																																								 ultraDataColumn21,
																																								 ultraDataColumn22,
																																								 ultraDataColumn23,
																																								 ultraDataColumn24});
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
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1469, 28);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(96, 26);
			this.btnSalir.TabIndex = 1054;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnSalir, "Salir");
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// udsResponsable
			// 
			ultraDataColumn25.DataType = typeof(System.UInt32);
			this.udsResponsable.Band.Columns.AddRange(new object[] {
																															 ultraDataColumn25,
																															 ultraDataColumn26});
			// 
			// frmrequerimientoreporte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(1584, 616);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Name = "frmrequerimientoreporte";
			this.Text = "Reporte de requerimientos";
			this.Load += new System.EventHandler(this.frmrequerimientoreporte_Load);
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.groupBox4.ResumeLayout(false);
			this.tabReporte.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgSeguimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimiento)).EndInit();
			this.tabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgRequerimientoReporteHistorial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoReporteHistorial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsResponsable)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmrequerimientoreporte_Load(object sender, System.EventArgs e)
		{
			//Cargar el nombre de la empresa que se logeo
			string sempresa = Funcion.sEjecucion(frmRequerimiento.sconexionInicio, string.Format("Select Top 1 e.Empresa From Empresas e Where e.BaseDeDatos = '{0}'", MenuLatinium.stDirFacturacion.Trim()));
			//CargarResponsable
			this.cmbResponsable.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionInicio, "Exec RequerimientoResponsable");
			//Traer de la tabla de inicio los usuarios y el logeado
			try
			{
				//Seleccionar el personal que ingreso al sistema que es responsable
				this.cmbResponsable.Value = frmRequerimiento.idUsuarioInicio;

				if((int)this.cmbResponsable.Value == 0)
				{
					MessageBox.Show(string.Format("Su usuario {0}, no se encuentra registrado como usuario de la empresa {1}", 
						MenuLatinium.stUsuario,
						sempresa
						), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			//Ubicar por default las fechas al día de hoy para no estesar la base
			#region Fechas default
			DateTime dtInicio = DateTime.Today.AddMonths(-540);
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
			this.dtDesde.Value = Convert.ToDateTime("14/02/2024");
			this.dtHasta.Value = DateTime.Today;
			#endregion Fechas default
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			RequerimientoReporte();
		}

		private void RequerimientoSetterarGrid()
		{
			System.Data.DataTable dt;

			dt = this.udgSeguimiento.DataSource as System.Data.DataTable;
			if (dt != null)
			{
				dt.Rows.Clear();
			}

			dt = this.udgRequerimientoReporteHistorial.DataSource as System.Data.DataTable;
			if (dt != null)
			{
				dt.Rows.Clear();
			}
								this.lblRegistros.Text = "";

		}

		private void RequerimientoReporte()
		{
			int idResponsable = 0;
			string dtDate = "";
			string dtHasta = "";
			string Numero = "";
			int idEstado = 0;
			int iOpcion = 0;

			try { idResponsable = (int)this.cmbResponsable.Value;} 
			catch(Exception ex){string e = ex.Message;}
			try { dtDate = Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd");} 
			catch(Exception ex){string e = ex.Message;}
			try { dtHasta = Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd");} 
			catch(Exception ex){string e = ex.Message;}
			try { Numero = (string)this.txtNumero.Text;} 
			catch(Exception ex){string e = ex.Message;}
			try { idEstado = (int)this.cmbEstado.Value;} 
			catch(Exception ex){string e = ex.Message;}
			try { iOpcion = (int)this.optReporte.Value;} 
			catch(Exception ex){string e = ex.Message;}

			RequerimientoSetterarGrid();

			if(iOpcion == 0)//Reporte General
			{
				if (tabReporte.TabPages.Contains(tabPage1))
					tabReporte.SelectedTab = tabPage1;

				try
				{
					// Mostrar el modal de "procesando"
					MenuLatinium.MostrarModalProcesando();

					string squery = string.Format("EXEC RequerimientoReporteReq '{0}','{1}',{2},{3},'{4}',{5}", 
						dtDate,
						dtHasta,
						idResponsable,
						idEstado,
						Numero,
						idRequerimientoDepartamento
						);
					//Cargar la tabla de  seguimientos
					this.udgSeguimiento.DataSource = 
						FuncionesProcedimientos.dtEjecucionGeneral(frmRequerimiento.sconexionPoint,squery); 
					//Contador de filas
					Funcion.ContadorFilas(this.udgSeguimiento, "Row");
					this.lblRegistros.Text = this.udgSeguimiento.Rows.Count + " REGISTROS";
				}
				catch (Exception ex)
				{
					// Manejar cualquier excepción que pueda ocurrir durante la consulta
					MessageBox.Show("Ocurrió un error durante la consulta Reporte General: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				finally
				{
					// Ocultar el modal de "procesando" una vez que se haya completado la consulta
					MenuLatinium.OcultarModalProcesando();
				}
			}
			if(iOpcion == 1)//Reporte General
			{
				if (tabReporte.TabPages.Contains(tabPage2))
					tabReporte.SelectedTab = tabPage2;

				try
				{
					// Mostrar el modal de "procesando"
					MenuLatinium.MostrarModalProcesando();

					string squery = string.Format("EXEC RequerimientoReporteHistorial '{0}','{1}',{2},{3},'{4}',{5}", 
						dtDate,
						dtHasta,
						idResponsable,
						idEstado,
						Numero,
						idRequerimientoDepartamento
						);

					//Cargar la tabla de historial
					this.udgRequerimientoReporteHistorial.DataSource = 
						FuncionesProcedimientos.dtEjecucionGeneral(frmRequerimiento.sconexionPoint,squery); 

					//Contador de filas
					this.udgRequerimientoReporteHistorial.DataSource =
						FuncionesProcedimientos.dtEjecucionGeneral(frmRequerimiento.sconexionPoint, squery);

					// Habilitar agrupación
					this.udgRequerimientoReporteHistorial.DisplayLayout.Override.AllowGroupBy =
						Infragistics.Win.DefaultableBoolean.True;

					this.udgRequerimientoReporteHistorial.DisplayLayout.GroupByBox.Hidden = false;

					this.udgRequerimientoReporteHistorial.DisplayLayout.GroupByBox.Prompt =
						"Arrastre aquí una columna para agrupar";

					// Contador
					Funcion.ContadorFilas(this.udgRequerimientoReporteHistorial, "Row");
										this.lblRegistros.Text = this.udgRequerimientoReporteHistorial.Rows.Count + " REGISTROS";
				}
				catch (Exception ex)
				{
					// Manejar cualquier excepción que pueda ocurrir durante la consulta
					MessageBox.Show("Ocurrió un error durante la consulta Reporte Detallado: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				finally
				{
					// Ocultar el modal de "procesando" una vez que se haya completado la consulta
					MenuLatinium.OcultarModalProcesando();
				}
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			try
			{
				int iOpcion = 0;
				
				if(this.optReporte.Value != null) iOpcion = (int)this.optReporte.Value;

				if(iOpcion == 0)
				{
					// Mostrar el modal de "procesando"
					MenuLatinium.MostrarModalProcesando();
					FuncionesProcedimientos.ExportaExcel(this.udgSeguimiento);
				}
				if(iOpcion == 1)
				{
					// Mostrar el modal de "procesando"
					MenuLatinium.MostrarModalProcesando();
					FuncionesProcedimientos.ExportaExcel(this.udgRequerimientoReporteHistorial);
				}

			}
			catch (Exception ex)
			{
				// Manejar cualquier excepción que pueda ocurrir durante la consulta
				MessageBox.Show("Ocurrió un error durante la consulta: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				// Ocultar el modal de "procesando" una vez que se haya completado la consulta
				MenuLatinium.OcultarModalProcesando();
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void optRegiones_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

	}
}
