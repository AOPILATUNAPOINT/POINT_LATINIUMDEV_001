using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmTiempoContratacion.
	/// </summary>
	public class frmTiempoContratacion : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidTiempoContratacion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaSolicitud;
		private System.Windows.Forms.Label lblFechaSolicitud;
		private System.Windows.Forms.Label lblSolicita;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtSolicita;
		private System.Windows.Forms.Label lblidDepartamento;
		private System.Windows.Forms.Label lblidCargo;
		private System.Windows.Forms.Label lblFechacontratacion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaContratacion;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.Misc.UltraButton btnNuevo;
		private Infragistics.Win.Misc.UltraButton btnEditar;
		public Infragistics.Win.UltraWinGrid.UltraGrid udtG;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblBuscar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.Misc.UltraButton btnCancela;
		private Infragistics.Win.Misc.UltraButton btnGrabar;
		private Infragistics.Win.Misc.UltraButton btnSalir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbidCargo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbidDepartamento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		
		bool bNuevo = false;
		bool bEdicion = false;
		private Infragistics.Win.Misc.UltraButton btnbuscar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.ComponentModel.Container components = null;

		public frmTiempoContratacion()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTiempoContratacion));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTiempoContratacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Solicita", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContratacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("R_D.Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("R_C.Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTiempoContratacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSolicitud");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Solicita");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaContratacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("R_D.Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("R_C.Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.txtidTiempoContratacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblFechaSolicitud = new System.Windows.Forms.Label();
			this.lblSolicita = new System.Windows.Forms.Label();
			this.lblidDepartamento = new System.Windows.Forms.Label();
			this.dtFechaSolicitud = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtSolicita = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblidCargo = new System.Windows.Forms.Label();
			this.lblFechacontratacion = new System.Windows.Forms.Label();
			this.dtFechaContratacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnNuevo = new Infragistics.Win.Misc.UltraButton();
			this.btnEditar = new Infragistics.Win.Misc.UltraButton();
			this.udtG = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblBuscar = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnCancela = new Infragistics.Win.Misc.UltraButton();
			this.btnGrabar = new Infragistics.Win.Misc.UltraButton();
			this.btnSalir = new Infragistics.Win.Misc.UltraButton();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbidCargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbidDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnbuscar = new Infragistics.Win.Misc.UltraButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.txtidTiempoContratacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSolicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSolicita)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaContratacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udtG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// txtidTiempoContratacion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidTiempoContratacion.Appearance = appearance1;
			this.txtidTiempoContratacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidTiempoContratacion.Enabled = false;
			this.txtidTiempoContratacion.FormatString = "0";
			this.txtidTiempoContratacion.Location = new System.Drawing.Point(856, 136);
			this.txtidTiempoContratacion.MinValue = 0;
			this.txtidTiempoContratacion.Name = "txtidTiempoContratacion";
			this.txtidTiempoContratacion.PromptChar = ' ';
			this.txtidTiempoContratacion.Size = new System.Drawing.Size(48, 22);
			this.txtidTiempoContratacion.TabIndex = 289;
			this.txtidTiempoContratacion.Visible = false;
			// 
			// lblFechaSolicitud
			// 
			this.lblFechaSolicitud.AutoSize = true;
			this.lblFechaSolicitud.Location = new System.Drawing.Point(8, 10);
			this.lblFechaSolicitud.Name = "lblFechaSolicitud";
			this.lblFechaSolicitud.Size = new System.Drawing.Size(79, 17);
			this.lblFechaSolicitud.TabIndex = 292;
			this.lblFechaSolicitud.Text = "Fecha Solicitud";
			this.lblFechaSolicitud.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblSolicita
			// 
			this.lblSolicita.AutoSize = true;
			this.lblSolicita.Location = new System.Drawing.Point(248, 10);
			this.lblSolicita.Name = "lblSolicita";
			this.lblSolicita.Size = new System.Drawing.Size(39, 17);
			this.lblSolicita.TabIndex = 293;
			this.lblSolicita.Text = "Solicita";
			this.lblSolicita.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblidDepartamento
			// 
			this.lblidDepartamento.AutoSize = true;
			this.lblidDepartamento.Location = new System.Drawing.Point(8, 42);
			this.lblidDepartamento.Name = "lblidDepartamento";
			this.lblidDepartamento.Size = new System.Drawing.Size(76, 17);
			this.lblidDepartamento.TabIndex = 294;
			this.lblidDepartamento.Text = "Departamento";
			this.lblidDepartamento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaSolicitud
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaSolicitud.Appearance = appearance2;
			this.dtFechaSolicitud.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaSolicitud.DateButtons.Add(dateButton1);
			this.dtFechaSolicitud.Enabled = false;
			this.dtFechaSolicitud.Format = "dd/MM/yyyy";
			this.dtFechaSolicitud.Location = new System.Drawing.Point(104, 8);
			this.dtFechaSolicitud.Name = "dtFechaSolicitud";
			this.dtFechaSolicitud.NonAutoSizeHeight = 23;
			this.dtFechaSolicitud.Size = new System.Drawing.Size(112, 21);
			this.dtFechaSolicitud.SpinButtonsVisible = true;
			this.dtFechaSolicitud.TabIndex = 295;
			this.dtFechaSolicitud.Value = ((object)(resources.GetObject("dtFechaSolicitud.Value")));
			this.dtFechaSolicitud.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaSolicitud_BeforeDropDown);
			// 
			// txtSolicita
			// 
			appearance3.BackColorDisabled = System.Drawing.Color.Silver;
			appearance3.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance3.BorderColor = System.Drawing.Color.Black;
			appearance3.BorderColor3DBase = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSolicita.Appearance = appearance3;
			this.txtSolicita.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSolicita.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSolicita.Enabled = false;
			this.txtSolicita.Location = new System.Drawing.Point(304, 8);
			this.txtSolicita.Name = "txtSolicita";
			this.txtSolicita.Size = new System.Drawing.Size(320, 22);
			this.txtSolicita.TabIndex = 296;
			this.txtSolicita.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSolicita_KeyDown);
			this.txtSolicita.ValueChanged += new System.EventHandler(this.txtSolicita_ValueChanged);
			// 
			// lblidCargo
			// 
			this.lblidCargo.AutoSize = true;
			this.lblidCargo.Location = new System.Drawing.Point(368, 42);
			this.lblidCargo.Name = "lblidCargo";
			this.lblidCargo.Size = new System.Drawing.Size(33, 17);
			this.lblidCargo.TabIndex = 299;
			this.lblidCargo.Text = "Cargo";
			this.lblidCargo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblFechacontratacion
			// 
			this.lblFechacontratacion.AutoSize = true;
			this.lblFechacontratacion.Location = new System.Drawing.Point(655, 10);
			this.lblFechacontratacion.Name = "lblFechacontratacion";
			this.lblFechacontratacion.Size = new System.Drawing.Size(117, 17);
			this.lblFechacontratacion.TabIndex = 300;
			this.lblFechacontratacion.Text = "Fecha De Contratación";
			this.lblFechacontratacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaContratacion
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaContratacion.Appearance = appearance4;
			this.dtFechaContratacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaContratacion.DateButtons.Add(dateButton2);
			this.dtFechaContratacion.Enabled = false;
			this.dtFechaContratacion.Format = "dd/MM/yyyy";
			this.dtFechaContratacion.Location = new System.Drawing.Point(798, 8);
			this.dtFechaContratacion.Name = "dtFechaContratacion";
			this.dtFechaContratacion.NonAutoSizeHeight = 23;
			this.dtFechaContratacion.Size = new System.Drawing.Size(120, 21);
			this.dtFechaContratacion.SpinButtonsVisible = true;
			this.dtFechaContratacion.TabIndex = 301;
			this.dtFechaContratacion.Value = ((object)(resources.GetObject("dtFechaContratacion.Value")));
			this.dtFechaContratacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaContratacion_KeyDown);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(655, 41);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 19);
			this.lblEstado.TabIndex = 302;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnNuevo
			// 
			this.btnNuevo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnNuevo.Location = new System.Drawing.Point(8, 88);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(75, 25);
			this.btnNuevo.TabIndex = 303;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnEditar.Enabled = false;
			this.btnEditar.Location = new System.Drawing.Point(88, 88);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(75, 25);
			this.btnEditar.TabIndex = 305;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// udtG
			// 
			this.udtG.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udtG.Cursor = System.Windows.Forms.Cursors.Default;
			this.udtG.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			this.udtG.DisplayLayout.Appearance = appearance5;
			ultraGridBand1.AddButtonCaption = "Cliente";
			ultraGridColumn1.Header.Caption = "Tiempo Contratacion";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 95;
			ultraGridColumn2.Header.Caption = "Fecha Solicitud";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 95;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 220;
			ultraGridColumn4.Header.Caption = "Departamento";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 95;
			ultraGridColumn5.Header.Caption = "Cargo";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.Caption = "Fecha Contratacion";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 95;
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 220;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 220;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 220;
			ultraGridColumn10.Header.Caption = "Estado";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 220;
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
																										 ultraGridColumn10});
			this.udtG.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.udtG.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udtG.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udtG.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.udtG.DisplayLayout.Override.CardAreaAppearance = appearance6;
			this.udtG.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.udtG.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udtG.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.udtG.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udtG.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udtG.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.udtG.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.udtG.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.udtG.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.udtG.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.udtG.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.udtG.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.udtG.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.udtG.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.udtG.Location = new System.Drawing.Point(8, 168);
			this.udtG.Name = "udtG";
			this.udtG.Size = new System.Drawing.Size(912, 200);
			this.udtG.TabIndex = 308;
			this.udtG.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.udtG_BeforeSelectChange);
			this.udtG.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.udtG_InitializeLayout);
			this.udtG.Click += new System.EventHandler(this.udtG_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(int);
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
			// lblBuscar
			// 
			this.lblBuscar.AutoSize = true;
			this.lblBuscar.Location = new System.Drawing.Point(8, 139);
			this.lblBuscar.Name = "lblBuscar";
			this.lblBuscar.Size = new System.Drawing.Size(111, 17);
			this.lblBuscar.TabIndex = 310;
			this.lblBuscar.Text = "Nombre/Fecha/Cargo";
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
			// btnCancela
			// 
			this.btnCancela.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnCancela.Enabled = false;
			this.btnCancela.Location = new System.Drawing.Point(248, 88);
			this.btnCancela.Name = "btnCancela";
			this.btnCancela.Size = new System.Drawing.Size(75, 25);
			this.btnCancela.TabIndex = 311;
			this.btnCancela.Text = "&Cancelar";
			this.btnCancela.Click += new System.EventHandler(this.btnCancela_Click);
			// 
			// btnGrabar
			// 
			this.btnGrabar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnGrabar.Enabled = false;
			this.btnGrabar.Location = new System.Drawing.Point(168, 88);
			this.btnGrabar.Name = "btnGrabar";
			this.btnGrabar.Size = new System.Drawing.Size(75, 25);
			this.btnGrabar.TabIndex = 314;
			this.btnGrabar.Text = "&Grabar";
			this.btnGrabar.Click += new System.EventHandler(this.btnGrabar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnSalir.Location = new System.Drawing.Point(328, 88);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(75, 25);
			this.btnSalir.TabIndex = 315;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// cmbidCargo
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbidCargo.Appearance = appearance10;
			this.cmbidCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbidCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidCargo.DataSource = this.ultraDataSource4;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 100;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 180;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbidCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbidCargo.DisplayMember = "Descripcion";
			this.cmbidCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbidCargo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbidCargo.Enabled = false;
			this.cmbidCargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidCargo.Location = new System.Drawing.Point(424, 40);
			this.cmbidCargo.Name = "cmbidCargo";
			this.cmbidCargo.Size = new System.Drawing.Size(200, 21);
			this.cmbidCargo.TabIndex = 317;
			this.cmbidCargo.ValueMember = "idCargo";
			this.cmbidCargo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbidCargo_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// cmbidDepartamento
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbidDepartamento.Appearance = appearance11;
			this.cmbidDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbidDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidDepartamento.DataSource = this.ultraDataSource3;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbidDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbidDepartamento.DisplayMember = "Descripcion";
			this.cmbidDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbidDepartamento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbidDepartamento.Enabled = false;
			this.cmbidDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidDepartamento.Location = new System.Drawing.Point(104, 40);
			this.cmbidDepartamento.Name = "cmbidDepartamento";
			this.cmbidDepartamento.Size = new System.Drawing.Size(200, 21);
			this.cmbidDepartamento.TabIndex = 318;
			this.cmbidDepartamento.ValueMember = "idDepartamento";
			this.cmbidDepartamento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbidDepartamento_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn15.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// txtBuscar
			// 
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(144, 136);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(260, 22);
			this.txtBuscar.TabIndex = 319;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			this.txtBuscar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscar_KeyPress);
			this.txtBuscar.TextChanged += new System.EventHandler(this.txtBuscar_TextChanged);
			// 
			// txtEstado
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstado.Appearance = appearance12;
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.FormatString = "0";
			this.txtEstado.Location = new System.Drawing.Point(808, 136);
			this.txtEstado.MaxValue = 2;
			this.txtEstado.MinValue = 0;
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(48, 22);
			this.txtEstado.TabIndex = 322;
			this.txtEstado.Value = 1;
			this.txtEstado.Visible = false;
			this.txtEstado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEstado_KeyDown);
			// 
			// btnbuscar
			// 
			this.btnbuscar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btnbuscar.Location = new System.Drawing.Point(424, 135);
			this.btnbuscar.Name = "btnbuscar";
			this.btnbuscar.Size = new System.Drawing.Size(75, 25);
			this.btnbuscar.TabIndex = 325;
			this.btnbuscar.Text = "&Buscar";
			this.btnbuscar.Click += new System.EventHandler(this.btnbuscar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(928, 8);
			this.groupBox1.TabIndex = 326;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 120);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(928, 8);
			this.groupBox2.TabIndex = 327;
			this.groupBox2.TabStop = false;
			// 
			// frmTiempoContratacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(928, 374);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnbuscar);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.cmbidDepartamento);
			this.Controls.Add(this.cmbidCargo);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGrabar);
			this.Controls.Add(this.btnCancela);
			this.Controls.Add(this.lblBuscar);
			this.Controls.Add(this.udtG);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.dtFechaContratacion);
			this.Controls.Add(this.lblFechacontratacion);
			this.Controls.Add(this.lblidCargo);
			this.Controls.Add(this.txtSolicita);
			this.Controls.Add(this.dtFechaSolicitud);
			this.Controls.Add(this.lblidDepartamento);
			this.Controls.Add(this.lblSolicita);
			this.Controls.Add(this.lblFechaSolicitud);
			this.Controls.Add(this.txtidTiempoContratacion);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmTiempoContratacion";
			this.Text = "frmTiempoContratacion";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmTiempoContratacion_KeyDown);
			this.Load += new System.EventHandler(this.frmTiempoContratacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidTiempoContratacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSolicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSolicita)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaContratacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udtG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		
		private void Aceptar()
		{
			if (this.udtG.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros encontrados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtBuscar.Focus();			
			}
			else if (this.udtG.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if (this.udtG.ActiveRow.Cells["idCliente"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else this.DialogResult = DialogResult.OK;
		}
		private void frmTiempoContratacion_Load(object sender, System.EventArgs e)
		{
			this.cmbidDepartamento.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idDepartamento, Descripcion from RolDepartamento");
			this.cmbidCargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idCargo,Descripcion from RolCargo");
//			this.dtFechaSolicitud.Value = DateTime.Today;
			this.udtG.DataSource =  Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaGeneralTiempoDeContratacion ''");
		}
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			#region vaciar y habilitar controles
			this.btnCancela_Click(sender, e);
			this.dtFechaSolicitud.Value = DateTime.Today;
			this.dtFechaContratacion.Value = DateTime.Today;
			this.txtSolicita.Enabled=true;
			this.cmbidDepartamento.Enabled=true;
			this.cmbidCargo.Enabled=true;
			this.dtFechaContratacion.Enabled=true;
			this.txtEstado.Enabled= true;
			this.txtEstado.Value=1;
      this.txtEstado.Visible= false;
			this.btnNuevo.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGrabar.Enabled = true;
			this.btnCancela.Enabled = true;	
			this.txtEstado.Visible = true;
			this.udtG.Enabled=false;
			this.lblEstado.Text="Estado";
			#endregion vaciar y habilitar controles
			bNuevo = true;
			bEdicion = true;
			this.txtSolicita.Focus();		
		}

		private void btnCancela_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;
			#region Desabilita Controles
			bNuevo = false;
			bEdicion = false;
			this.txtidTiempoContratacion.Value=0;
			this.dtFechaSolicitud.Value = System.DBNull.Value;
			this.txtSolicita.Text= "";
			this.txtSolicita.Enabled=false;
			this.cmbidDepartamento.Value = System.DBNull.Value;
			this.cmbidDepartamento.Enabled=false;
			this.cmbidCargo.Value = System.DBNull.Value;
			this.cmbidCargo.Enabled=false;
			this.dtFechaContratacion.Value = System.DBNull.Value;
			this.dtFechaContratacion.Enabled=false;	
			this.txtEstado.Value=1;
			this.txtEstado.Enabled=false;
			this.btnNuevo.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGrabar.Enabled = false;
			this.btnCancela.Enabled = false;
			this.txtEstado.Visible = false;
			this.lblEstado.Text="Estado";
			this.udtG.Enabled=true;
			#endregion Desabilita Controles	
		}

		private void btnGrabar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbTexto(this.txtSolicita, 5, 60, "Ingrese Nombre del Solicitante")) return;
			if (!Validacion.vbComboVacio(this.cmbidDepartamento, "Seleccione el departamento ")) return;
			if (!Validacion.vbComboVacio(this.cmbidCargo, "Seleccione el Cargo ")) return;
			#region Grabar Tiempo Contratacion
			string sSQL = string.Format("Exec GuardaTiempoContratacion {0},'{1}','{2}',{3},{4},'{5}',{6}",
				(int)this.txtidTiempoContratacion.Value, Convert.ToDateTime(this.dtFechaSolicitud.Value).ToString("yyyyMMdd"),
				this.txtSolicita.Text.ToString(),(int)this.cmbidDepartamento.Value ,
				(int)this.cmbidCargo.Value,Convert.ToDateTime(this.dtFechaContratacion.Value).ToString("yyyyMMdd"),
				(int) this.txtEstado.Value);
			this.txtidTiempoContratacion.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQL);
			#endregion Grabar Tiempo Contratacion
			#region desabilitar controles al guardar
			this.dtFechaContratacion.Enabled= false;
			this.txtSolicita.Enabled=false;
			this.cmbidDepartamento.Enabled=false;
			this.cmbidCargo.Enabled=false;
			this.dtFechaContratacion.Enabled=false;
			this.txtEstado.Enabled= false;
			this.txtEstado.Visible=false;
			this.btnNuevo.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGrabar.Enabled = false;
			this.btnCancela.Enabled = true;
			this.udtG.Enabled=false;
			this.udtG.Enabled=true;
			#endregion desabilita controles al guardar
			bNuevo = false;
			bEdicion = false;
		}
		private void udtG_Click(object sender, System.EventArgs e)
		{
			this.txtEstado.Visible= false;
			this.btnEditar.Enabled = true;
      this.btnNuevo.Enabled= true;
			this.btnCancela.Enabled = true;
		}

		private void txtBuscar_TextChanged(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ConsultaGeneralTiempoDeContratacion '{0}'", this.txtBuscar.Text.ToString());
			this.udtG.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			
		}

		private void dtFechaSolicitud_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void udtG_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			#region habilita controles al precionar editar
			this.lblEstado.Text="Estado";
      int Estado = (int)this.txtEstado.Value;

			if( Estado == 2){
//				this.txtSolicita.Enabled=false;
//				this.cmbidDepartamento.Enabled=false;
//				this.cmbidCargo.Enabled=false;
//				this.dtFechaContratacion.Enabled=false;	
//				this.txtEstado.Enabled=false;
//				this.btnNuevo.Enabled = false;
//				this.btnEditar.Enabled = false;
//				this.btnGrabar.Enabled = true;
//				this.btnCancela.Enabled = true;
//				this.udtG.Enabled=false;
				MessageBox.Show("No puede editar el registro seleccionado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}else{

			this.txtSolicita.Enabled=true;
			this.cmbidDepartamento.Enabled=true;
			this.cmbidCargo.Enabled=true;
			this.dtFechaContratacion.Enabled=true;	
			this.txtEstado.Enabled=true;
			this.btnNuevo.Enabled = false;
				this.txtEstado.Enabled=true;
				this.txtEstado.Visible= true;
			this.btnEditar.Enabled = false;
			this.btnGrabar.Enabled = true;
			this.btnCancela.Enabled = true;
			this.udtG.Enabled=false;
				bEdicion=true;
				this.txtSolicita.Focus();
			}
			#endregion habilita controles al precionar editar
			
			
		}

		private void udtG_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			#region Capturar datos del datagrid a los cajas de texto
			if (this.udtG.ActiveRow.Cells["idTiempoContratacion"].Value == DBNull.Value) return;
			this.txtidTiempoContratacion.Value = (int)this.udtG.ActiveRow.Cells["idTiempoContratacion"].Value;
			this.dtFechaSolicitud.Value = this.udtG.ActiveRow.Cells["FechaSolicitud"].Value;
			this.txtSolicita.Text = this.udtG.ActiveRow.Cells["Solicita"].Value.ToString();
			this.dtFechaContratacion.Value = this.udtG.ActiveRow.Cells["FechaContratacion"].Value;
			this.txtEstado.Value = this.udtG.ActiveRow.Cells["Estado"].Value;
			this.cmbidDepartamento.Value = (int)this.udtG.ActiveRow.Cells["idDepartamento"].Value;
			this.cmbidCargo.Value = (int)this.udtG.ActiveRow.Cells["idCargo"].Value;
			this.lblEstado.Text = this.udtG.ActiveRow.Cells["EstadoM"].Value.ToString();
			#endregion Capturar datos del datagrid a los cajas de texto
			
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void txtSolicita_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtSolicita_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtFechaContratacion.Focus();	
		}

		private void dtFechaContratacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbidCargo.Focus();	
		}

		private void cmbidCargo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbidDepartamento.Focus();	
		}

		private void cmbidDepartamento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtEstado.Focus();	
		}

		private void txtEstado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnGrabar.Focus();	
		}

		private void frmTiempoContratacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
			
				if ((int)this.txtidTiempoContratacion.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
//					if (e.KeyCode == Keys.F5) this.Consultar((int)this.txtIdBodega.Value);					
				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancela_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGrabar_Click(sender, e);				
			}
		}

		private void txtBuscar_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.udtG.DataSource =  Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaGeneralTiempoDeContratacion ''");
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode==Keys.Enter)
				this.btnbuscar_Click(sender,e);
		}

		private void btnbuscar_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ConsultaGeneralTiempoDeContratacion '{0}'", this.txtBuscar.Text.ToString());
			this.udtG.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}
	}
}
