using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmEntregaDeObsequiosClientesCumpleaneros.
	/// </summary>
	public class frmEntregaDeObsequiosClientesCumpleaneros : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblFecha;
		private System.Windows.Forms.Label lblCedula;
		private System.Windows.Forms.Label lblNombre;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCedula;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label lblnotas;
		private System.Windows.Forms.Label lblBusqueda;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtbusqueda;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		private System.Windows.Forms.Button btnCliente;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultgconsulta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor rtbNotas;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label lblCodigo;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtReferido;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtobsequio;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultgReferidos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmEntregaDeObsequiosClientesCumpleaneros()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmEntregaDeObsequiosClientesCumpleaneros));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Email");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("B.Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("p.Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ob.idPersonal");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("B.Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("p.Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ob.idPersonal");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClientesReferidos");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Email");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idClientesReferidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblFecha = new System.Windows.Forms.Label();
			this.lblCedula = new System.Windows.Forms.Label();
			this.lblNombre = new System.Windows.Forms.Label();
			this.txtCedula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblnotas = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblBusqueda = new System.Windows.Forms.Label();
			this.txtbusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.ultgconsulta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnCliente = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.rtbNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblCodigo = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtReferido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtobsequio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultgReferidos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtbusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgconsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rtbNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtobsequio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgReferidos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			this.SuspendLayout();
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(120, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 640;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);			
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// lblFecha
			// 
			this.lblFecha.AutoSize = true;
			this.lblFecha.Location = new System.Drawing.Point(8, 10);
			this.lblFecha.Name = "lblFecha";
			this.lblFecha.Size = new System.Drawing.Size(39, 16);
			this.lblFecha.TabIndex = 641;
			this.lblFecha.Text = "Fecha:";
			// 
			// lblCedula
			// 
			this.lblCedula.AutoSize = true;
			this.lblCedula.Location = new System.Drawing.Point(8, 42);
			this.lblCedula.Name = "lblCedula";
			this.lblCedula.Size = new System.Drawing.Size(43, 16);
			this.lblCedula.TabIndex = 642;
			this.lblCedula.Text = "Cedula:";
			// 
			// lblNombre
			// 
			this.lblNombre.AutoSize = true;
			this.lblNombre.Location = new System.Drawing.Point(216, 42);
			this.lblNombre.Name = "lblNombre";
			this.lblNombre.Size = new System.Drawing.Size(48, 16);
			this.lblNombre.TabIndex = 643;
			this.lblNombre.Text = "Nombre:";
			// 
			// txtCedula
			// 
			appearance2.BorderColor = System.Drawing.Color.Black;
			appearance2.BorderColor3DBase = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCedula.Appearance = appearance2;
			this.txtCedula.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCedula.Enabled = false;
			this.txtCedula.Location = new System.Drawing.Point(72, 40);
			this.txtCedula.MaxLength = 10;
			this.txtCedula.Name = "txtCedula";
			this.txtCedula.Size = new System.Drawing.Size(120, 21);
			this.txtCedula.TabIndex = 644;
			this.txtCedula.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCedula_KeyDown);
			this.txtCedula.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCedula_KeyPress);
			this.txtCedula.Validating += new System.ComponentModel.CancelEventHandler(this.txtCedula_Validating);
			this.txtCedula.Validated += new System.EventHandler(this.txtCedula_Validated);
			// 
			// txtNombre
			// 
			appearance3.BorderColor = System.Drawing.Color.Black;
			appearance3.BorderColor3DBase = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance3;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(280, 40);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(220, 21);
			this.txtNombre.TabIndex = 645;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);			
			// 
			// lblnotas
			// 
			this.lblnotas.AutoSize = true;
			this.lblnotas.Location = new System.Drawing.Point(8, 95);
			this.lblnotas.Name = "lblnotas";
			this.lblnotas.Size = new System.Drawing.Size(37, 16);
			this.lblnotas.TabIndex = 646;
			this.lblnotas.Text = "Notas:";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn6.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// lblBusqueda
			// 
			this.lblBusqueda.AutoSize = true;
			this.lblBusqueda.Location = new System.Drawing.Point(8, 346);
			this.lblBusqueda.Name = "lblBusqueda";
			this.lblBusqueda.Size = new System.Drawing.Size(84, 16);
			this.lblBusqueda.TabIndex = 650;
			this.lblBusqueda.Text = "Cedula/Nombre";
			// 
			// txtbusqueda
			// 
			appearance4.BackColorDisabled = System.Drawing.Color.Silver;
			appearance4.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance4.BorderColor = System.Drawing.Color.Black;
			appearance4.BorderColor3DBase = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtbusqueda.Appearance = appearance4;
			this.txtbusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtbusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtbusqueda.Location = new System.Drawing.Point(112, 344);
			this.txtbusqueda.Name = "txtbusqueda";
			this.txtbusqueda.Size = new System.Drawing.Size(240, 21);
			this.txtbusqueda.TabIndex = 651;
			this.txtbusqueda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtbusqueda_KeyDown);
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(360, 343);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 23);
			this.btnConsultar.TabIndex = 652;
			this.btnConsultar.Text = "Buscar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// ultgconsulta
			// 
			this.ultgconsulta.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultgconsulta.DataSource = this.ultraDataSource2;
			appearance5.BackColor = System.Drawing.Color.White;
			appearance5.ForegroundAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultgconsulta.DisplayLayout.Appearance = appearance5;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 5;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 61;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 7;
			ultraGridColumn2.Width = 220;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 75;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 9;
			ultraGridColumn4.Width = 300;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Cedula";
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Width = 120;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 72;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 160;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Local";
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 160;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Vendedor";
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 220;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 81;
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
																										 ultraGridColumn11});
			this.ultgconsulta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultgconsulta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultgconsulta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultgconsulta.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8.5F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultgconsulta.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgconsulta.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgconsulta.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultgconsulta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultgconsulta.Location = new System.Drawing.Point(8, 376);
			this.ultgconsulta.Name = "ultgconsulta";
			this.ultgconsulta.Size = new System.Drawing.Size(968, 152);
			this.ultgconsulta.TabIndex = 653;
			this.ultgconsulta.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultgconsulta_BeforeSelectChange);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn17});			
			// 
			// btnGuardar
			// 
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(168, 296);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 659;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(88, 296);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(72, 23);
			this.btnEditar.TabIndex = 656;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(256, 296);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 657;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(336, 296);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 658;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 296);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 654;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// txtidCliente
			// 
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.Location = new System.Drawing.Point(816, 296);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(16, 21);
			this.txtidCliente.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidCliente.TabIndex = 660;
			this.txtidCliente.Visible = false;			
			// 
			// btnCliente
			// 
			this.btnCliente.Enabled = false;
			this.btnCliente.Location = new System.Drawing.Point(512, 39);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(24, 22);
			this.btnCliente.TabIndex = 661;
			this.btnCliente.Text = "...";
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
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
			// rtbNotas
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.rtbNotas.Appearance = appearance10;
			this.rtbNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.rtbNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.rtbNotas.Enabled = false;
			this.rtbNotas.Location = new System.Drawing.Point(72, 72);
			this.rtbNotas.Multiline = true;
			this.rtbNotas.Name = "rtbNotas";
			this.rtbNotas.Size = new System.Drawing.Size(784, 63);
			this.rtbNotas.TabIndex = 665;
			this.rtbNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.rtbNotas_KeyDown_1);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(216, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 669;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(560, 10);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(53, 16);
			this.lblVendedor.TabIndex = 668;
			this.lblVendedor.Text = "Vendedor";
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(280, 7);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(220, 22);
			this.cmbBodega.TabIndex = 666;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodega_Validating);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// ultraDataSource3
			// 
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// cmbVendedor
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance11;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 300;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Width = 120;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(632, 7);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(220, 22);
			this.cmbVendedor.TabIndex = 667;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVendedor_KeyDown);
			// 
			// lblCodigo
			// 
			this.lblCodigo.AutoSize = true;
			this.lblCodigo.Location = new System.Drawing.Point(560, 42);
			this.lblCodigo.Name = "lblCodigo";
			this.lblCodigo.Size = new System.Drawing.Size(43, 16);
			this.lblCodigo.TabIndex = 670;
			this.lblCodigo.Text = "Codigo:";
			// 
			// txtCodigo
			// 
			appearance12.BorderColor = System.Drawing.Color.Black;
			appearance12.BorderColor3DBase = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance12;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(632, 40);
			this.txtCodigo.MaxLength = 15;
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(112, 21);
			this.txtCodigo.TabIndex = 671;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			this.txtCodigo.Validating += new System.ComponentModel.CancelEventHandler(this.txtCodigo_Validating);			
			// 
			// txtReferido
			// 
			this.txtReferido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtReferido.Enabled = false;
			this.txtReferido.Location = new System.Drawing.Point(864, 296);
			this.txtReferido.Name = "txtReferido";
			this.txtReferido.PromptChar = ' ';
			this.txtReferido.Size = new System.Drawing.Size(16, 21);
			this.txtReferido.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtReferido.TabIndex = 672;
			this.txtReferido.Visible = false;
			// 
			// txtobsequio
			// 
			this.txtobsequio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtobsequio.Enabled = false;
			this.txtobsequio.Location = new System.Drawing.Point(840, 296);
			this.txtobsequio.Name = "txtobsequio";
			this.txtobsequio.PromptChar = ' ';
			this.txtobsequio.Size = new System.Drawing.Size(16, 21);
			this.txtobsequio.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtobsequio.TabIndex = 673;
			this.txtobsequio.Visible = false;
			// 
			// ultgReferidos
			// 
			this.ultgReferidos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultgReferidos.DataSource = this.ultraDataSource4;
			appearance13.BackColor = System.Drawing.Color.White;
			this.ultgReferidos.DisplayLayout.Appearance = appearance13;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.Header.Caption = " Apellidos y Nombres ";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Width = 220;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 90;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Width = 90;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Width = 150;
			ultraGridColumn19.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Width = 300;
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 43;
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 80;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			ultraGridBand3.Header.Caption = "REFERIDOS";
			ultraGridBand3.HeaderVisible = true;
			this.ultgReferidos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultgReferidos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultgReferidos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultgReferidos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.ultgReferidos.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8.5F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultgReferidos.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgReferidos.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgReferidos.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.ultgReferidos.Enabled = false;
			this.ultgReferidos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultgReferidos.Location = new System.Drawing.Point(8, 144);
			this.ultgReferidos.Name = "ultgReferidos";
			this.ultgReferidos.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ultgReferidos.Size = new System.Drawing.Size(888, 128);
			this.ultgReferidos.TabIndex = 674;
			this.ultgReferidos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultgReferidos_AfterRowInsert);
			this.ultgReferidos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultgReferidos_KeyPress);
			this.ultgReferidos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultgReferidos_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27});
			// 
			// ultraDataSource5
			// 
			ultraDataColumn29.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// gbBotones
			// 
			this.gbBotones.Location = new System.Drawing.Point(0, 280);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(904, 8);
			this.gbBotones.TabIndex = 675;
			this.gbBotones.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 328);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(904, 8);
			this.groupBox1.TabIndex = 676;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(456, 345);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 677;
			// 
			// frmEntregaDeObsequiosClientesCumpleaneros
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(904, 534);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.ultgReferidos);
			this.Controls.Add(this.txtobsequio);
			this.Controls.Add(this.txtReferido);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.lblCodigo);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.rtbNotas);
			this.Controls.Add(this.btnCliente);
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.ultgconsulta);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.txtbusqueda);
			this.Controls.Add(this.lblBusqueda);
			this.Controls.Add(this.lblnotas);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtCedula);
			this.Controls.Add(this.lblNombre);
			this.Controls.Add(this.lblCedula);
			this.Controls.Add(this.lblFecha);
			this.Controls.Add(this.dtFecha);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmEntregaDeObsequiosClientesCumpleaneros";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Entrega De Obsequios Clientes Cumpleañeros";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmEntregaDeObsequiosClientesCumpleaneros_KeyDown);
			this.Load += new System.EventHandler(this.frmEntregaDeObsequiosClientesCumpleaneros_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtbusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgconsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rtbNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtobsequio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgReferidos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		int idCuponesWeb = 0;

		private void frmEntregaDeObsequiosClientesCumpleaneros_Load(object sender, System.EventArgs e)
		{
			
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			#region Habilitacion de controles 
			this.btnCancelar_Click(sender,e);

			this.dtFecha.Value = DateTime.Today;

			this.txtCodigo.Enabled=true;
			this.txtCedula.Enabled= true;
			this.rtbNotas.Enabled = true;
			this.cmbBodega.Enabled=true;
			this.cmbVendedor.Enabled=true;
			this.txtbusqueda.Text="";

			this.btnNuevo.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnConsultar.Enabled = false;
			this.txtbusqueda.Enabled= false;

			FuncionesProcedimientos.EstadoGrids(true,this.ultgReferidos);

			bNuevo = true;
			bEdicion = true;

			this.cmbBodega.Focus();
			#endregion Habilitacion de controles 
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region Bloqueo y vaciado de controles 
			this.txtbusqueda.Text="";
			string sSQL = string.Format("Exec ConsultaGeneralEntregaDeObsequiosClientesCumpleaneros  '{0}'", this.txtbusqueda.Text.ToString());
			this.ultgconsulta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			this.txtidCliente.Value = 0;
			this.txtCedula.Enabled= false;
			this.txtNombre.Enabled = false;
			this.rtbNotas.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(false,this.ultgReferidos);
			this.cmbBodega.Enabled=false;
			this.cmbVendedor.Enabled=false;
			this.rtbNotas.Enabled=false;
			this.dtFecha.Value = System.DBNull.Value;
			this.txtCedula.Text ="";
			this.txtNombre.Text ="";
			this.rtbNotas.Text ="";
			this.txtCodigo.Text="";
			this.txtCodigo.Enabled=false;
			this.cmbBodega.Text="";
			this.cmbVendedor.Text="";
			this.rtbNotas.Text="";
			this.txtbusqueda.Text="";			
			this.ultgReferidos.DataSource  = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientesReferidos,idCliente,Nombre,Telefono,Celular,Email,Nota From ClientesReferidos Where idCliente = 0");

			this.btnNuevo.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnConsultar.Enabled = true;
			this.txtbusqueda.Enabled= true;
			idCuponesWeb = 0;
			#endregion Bloqueo y vaciado de controles 
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();		
		}

		private void txtCedula_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtCedula_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCodigo.Focus();
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCodigo.Focus();
		}

		private void rtbNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) 
			{
				this.ultgReferidos.DisplayLayout.Bands[0].AddNew();
				this.ultgReferidos.ActiveCell = this.ultgReferidos.Rows[0].Cells["Nombre"];
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnCliente_Click(object sender, System.EventArgs e)
		{
			#region abrir pantalla datos del cliente
			if (this.txtCedula.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el Numero de Identificacion del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCedula.Focus();
			}
			else
			{
				FuncionesProcedimientos.AbreFacturacionClientes(5, this.txtCedula.Text.ToString(), true);
			}
			#endregion abrir pantalla datos del cliente
		}

		private void txtCedula_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if(!bEdicion)return;

			if (this.txtCedula.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el Numero de Identificacion del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCedula.Focus();
				e.Cancel = true;
				return;
			}
			
			this.txtidCliente.Value = Funcion.iEscalarSQL(cdsSeteoF,string.Format("if Exists(select idCliente from Cliente where Ruc= '{0}' and Proveedor = 0)select idCliente from Cliente where Ruc= '{0}' and Proveedor = 0  Else Select CONVERT(Int, 0)",this.txtCedula.Text.ToString()));
				
			if ((int)this.txtidCliente.Value == 0)
			{
				MessageBox.Show("Cliente no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCedula.Focus();
				e.Cancel = true;
				return;
			}

			int Existencia = Funcion.iEscalarSQL(cdsSeteoF,string.Format("if Exists(select idCliente from ClientesReferidos where idCliente ={0})select idCliente from ClientesReferidos where idCliente ={0} Else Select CONVERT(Int, 0)",(int)this.txtidCliente.Value));
			
			if (Existencia == 0)
			{
				MessageBox.Show("Cliente ya fue registrado en la Entrega de obsequios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCedula.Focus();
				e.Cancel = true;
				return;
			}

			/*
			string  Nombre = Funcion.sEscalarSQL(cdsSeteoF,string.Format("select Nombre from Cliente where idCliente = {0}",this.txtidCliente.Value));
			this.txtNombre.Text = Nombre;
			if (Existencia > 0)
			{							
				DateTime Fecha = Funcion.dtEscalarSQL(cdsSeteoF,string.Format("select Fecha from ClientesObsequios where idCliente = {0}",this.txtidCliente.Value));
				int Bodega = Funcion.iEscalarSQL(cdsSeteoF,string.Format("select Bodega from ClientesObsequios where idCliente = {0}",this.txtidCliente.Value));
				int IdPersonal = Funcion.iEscalarSQL(cdsSeteoF,string.Format("select IdPersonal from ClientesObsequios where idCliente = {0}",this.txtidCliente.Value));
				string codigo = Funcion.sEscalarSQL(cdsSeteoF,string.Format("select Codigo from ClientesObsequios where idCliente = {0}",this.txtidCliente.Value));
				string Notas = Funcion.sEscalarSQL(cdsSeteoF,string.Format("select ISNULL(Notas,'') from ClientesObsequios where idCliente = {0}",this.txtidCliente.Value));
				this.dtFecha.Value= Fecha;
				this.cmbBodega.Value= Bodega;
				this.cmbVendedor.Value= IdPersonal;
				this.txtCodigo.Text= codigo;
				this.rtbNotas.Text=Notas;
				this.ultgReferidos.DataSource  = Funcion.dvProcedimiento(cdsSeteoF, "select idClientesReferidos,idCliente,Nombre,Telefono,Celular,Email,Nota from ClientesReferidos where idCliente = '" + IdCliente+ "'");
				if (this.ultgReferidos.Rows.Count > 0) this.ultgReferidos.Focus(); else this.cmbBodega.Focus();
              
				this.rtbNotas.Focus();

				this.cmbVendedor.Enabled=false;
				this.cmbBodega.Enabled=false;
				this.txtCodigo.Enabled=false;
				this.txtCedula.Enabled=false;
				this.btnCliente.Enabled= true;
			}
			else
			{
				this.dtFecha.Value = DateTime.Today;
			}	
			*/		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ConsultaGeneralEntregaDeObsequiosClientesCumpleaneros  '{0}'", this.txtbusqueda.Text.ToString());
			this.ultgconsulta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			
			if (this.ultgconsulta.Rows.Count > 0) this.ultgconsulta.Focus(); else this.txtbusqueda.Focus();			

			this.btnNuevo.Enabled = true;
			this.btnEditar.Enabled = true;			
			this.btnCancelar.Enabled = true;			
		}

		private void txtbusqueda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode==Keys.Enter)
				this.btnConsultar_Click(sender,e);
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 43, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)	
			{				
				if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;
				if (this.cmbBodega.ActiveRow != null) 
					this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (this.cmbBodega.ActiveRow != null)
				{
					if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;
					this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
				}
			}
		}

		private void cmbBodega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{			
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.rtbNotas.Enabled=true;
			this.ultgReferidos.Enabled=true;
			this.btnNuevo.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnConsultar.Enabled = false;
			this.txtbusqueda.Enabled= false;
			
			FuncionesProcedimientos.EstadoGrids(true,this.ultgReferidos);

			bNuevo = false;
			bEdicion = true;	
		}

		private void ultgconsulta_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			#region Capturar datos del datagrid a los cajas de texto
			if(!bEdicion)
			{
				if (this.ultgconsulta.ActiveRow.Cells["idCliente"].Value == DBNull.Value) return;

				this.txtidCliente.Value = (int)this.ultgconsulta.ActiveRow.Cells["idCliente"].Value;
				this.dtFecha.Value = this.ultgconsulta.ActiveRow.Cells["Fecha"].Value;
				this.txtNombre.Text = this.ultgconsulta.ActiveRow.Cells["Nombre"].Value.ToString();
				this.rtbNotas.Text= this.ultgconsulta.ActiveRow.Cells["Notas"].Value.ToString();
				this.txtCedula.Text = this.ultgconsulta.ActiveRow.Cells["Ruc"].Value.ToString();
				this.txtCodigo.Text = this.ultgconsulta.ActiveRow.Cells["Codigo"].Value.ToString();
				this.cmbBodega.Value = (int)this.ultgconsulta.ActiveRow.Cells["Bodega"].Value;
				this.cmbVendedor.Value = (int)this.ultgconsulta.ActiveRow.Cells["IdPersonal"].Value;
				
				this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

				this.ultgReferidos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idClientesReferidos, idCliente, Nombre, Telefono, Celular, Email, Nota From ClientesReferidos Where idCliente = {0}", (int)this.txtidCliente.Value));
			}
			#endregion Capturar datos del datagrid a los cajas de texto
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region validacion campos generales
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una bodega")) return;
			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un vendedor")) return;
			if (!Validacion.vbIdentificacion(this.txtCedula, 5, cdsSeteoF)) return;
			if (!Validacion.vbTexto(this.txtCodigo, 6, 50, "Código")) return;

			if (this.ultgReferidos.Rows.Count <2)
			{
				MessageBox.Show("Para guardar debe ingresar al menos dos Referidos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);		
				return;
			}
			#endregion validacion campos generales

			#region validacion ultgReferidos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultgReferidos.Rows.All)
			{				
				if (dr.Cells["Nombre"].Value.ToString().Trim().Length == 0 )
				{
					MessageBox.Show("Debe ingresar el nombre del referido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					
					this.ultgReferidos.ActiveRow = dr;
					this.ultgReferidos.ActiveRow.Selected = true;
					return;
				}

				if (dr.Cells["Telefono"].Value.ToString().Trim().Length == 0 && dr.Cells["Celular"].Value.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Debe Ingresar el Telefono Convencional o el Telefono Celular del Contacto Seleccionado ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					
					this.ultgReferidos.ActiveRow = dr;
					this.ultgReferidos.ActiveRow.Selected = true;
					return;
				}	
			
				#region Telefono
				if (dr.Cells["Telefono"].Value.ToString().Trim().Length > 0)
				{
					if (dr.Cells["Telefono"].Value.ToString().Trim().Length != 9)
					{
						MessageBox.Show("El Telefono Convencional debe tener 9 Digitos.\n\nLos dos primeros corresponden al Codigo de Provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						
						this.ultgReferidos.ActiveRow = dr;
						this.ultgReferidos.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Telefono
				#region Celular
				if (dr.Cells["Celular"].Value.ToString().Trim().Length > 0)
				{
					if (dr.Cells["Celular"].Value.ToString().Trim().Length != 10)
					{
						MessageBox.Show("El Telefono Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					
						this.ultgReferidos.ActiveRow = dr;
						this.ultgReferidos.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Celular
				#region valida email data grid referidos 
				if (dr.Cells["Email"].Value.ToString().Trim().Length > 0)
				{
					if (!Validacion.bEMail(dr.Cells["Email"].Value.ToString())) 
					{
						this.ultgReferidos.ActiveRow = dr;
						this.ultgReferidos.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion valida email data grid referidos 
			}
			#endregion validacion ultgReferidos

			#region Guardar Cliente
			string sSQL = string.Format("Exec GuardaEntregaDeObsequiosClientesCumpleaneros {0},'{1}','{2}','{3}',{4},{5}",
				(int) this.txtidCliente.Value,
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),
				this.rtbNotas.Text.ToString(), 
				this.txtCodigo.Text.ToString(), 
				(int) this.cmbBodega.Value,
				(int) this.cmbVendedor.Value);
			this.txtReferido.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQL);
			#endregion Guardar Cliente

			#region Guarda Clientes Referidos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.ultgReferidos.Rows.All)
			{
				string sSQLDetalle = string.Format("Exec GuardaEntregaClientesReferidos  {0},{1},'{2}','{3}','{4}','{5}','{6}'", 
					drr.Cells["idClientesReferidos"].Value,
					(int)this.txtidCliente.Value,
					drr.Cells["Nombre"].Value.ToString(),
					drr.Cells["Telefono"].Value.ToString(),
					drr.Cells["Celular"].Value.ToString(),
					drr.Cells["Email"].Value.ToString(),
					drr.Cells["Nota"].Value.ToString());     
					
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			}	

			MessageBox.Show("Registrado guardado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);			
			#endregion Guarda Clientes 
  			
			string ssSQL = string.Format("Update CuponesWeb Set Estado = 2 Where idCuponesWeb = {0}", idCuponesWeb);
			Funcion.EjecutaSQL(cdsSeteoF, ssSQL);
		  
			int idCompra = 0;
			if (bNuevo)
			{
				string Sql = string.Format("Exec CreaEgresoBodegaObsequio {0}, {1}, {2}",
					(int)this.txtReferido.Value, (int)this.cmbBodega.Value, (int)this.txtidCliente.Value);
				idCompra = Funcion.iEscalarSQL(cdsSeteoF, Sql);

				Facturacion.Impresion(idCompra, false, 3, true, true, false, false, 8, 0.00m, 0, 0, "", DateTime.Today, 0, "", cdsSeteoF);
			}

			#region Controles
			this.ultgReferidos.Enabled=false;
			this.txtCedula.Enabled= false;
			this.txtNombre.Enabled = false;
			this.rtbNotas.Enabled = false;
			this.ultgReferidos.Enabled=false;
			this.cmbBodega.Enabled=false;
			this.cmbVendedor.Enabled=false;
			this.rtbNotas.Enabled=false;

			FuncionesProcedimientos.EstadoGrids(false, this.ultgReferidos);

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.txtbusqueda.Enabled= true;
			this.txtCodigo.Enabled=false;
			
			bNuevo = false;
			bEdicion = false;					
			#endregion Controles
		}
		
		private void ultgReferidos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ultgReferidos.ActiveCell.Column.Key.ToString().ToUpper() == "TELEFONO" || ultgReferidos.ActiveCell.Column.Key.ToString().ToUpper() == "CELULAR")			{
				Funcion.CamposNumericos(sender, e);}
		}

		private void rtbNotas_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) 
			{
				this.ultgReferidos.DisplayLayout.Bands[0].AddNew();
				this.ultgReferidos.ActiveCell = this.ultgReferidos.Rows[0].Cells["Nombre"];
			}
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbVendedor.Focus();
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCedula.Focus();
		}

		private void ultgReferidos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idClientesReferidos"].Value = 0;
		}

		private void ultgReferidos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.ultgReferidos);
		}

		private void txtCodigo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtCodigo.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el codigo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				e.Cancel = true;
				return;
			}
			
			idCuponesWeb = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Not Exists(Select idCuponesWeb From CuponesWeb Where Cupon = '{0}') Select CONVERT(Int , -1) Else If Exists(Select idCuponesWeb From CuponesWeb Where Cupon = '{0}' And Estado = 0) Select idCuponesWeb From CuponesWeb Where Cupon = '{0}' And Estado = 0 Else Select CONVERT(Int , 0)",
				this.txtCodigo.Text.ToString()));		
			
			if (idCuponesWeb == -1)
			{
				MessageBox.Show("El Codigo ingresado no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				e.Cancel = true;
				return;
			}

			if (idCuponesWeb == 0)
			{
				MessageBox.Show("El Codigo ingresado ya fue utilizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.rtbNotas.Focus();
		}

		private void frmEntregaDeObsequiosClientesCumpleaneros_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{			
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidCliente.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);													
				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}			
		}

		private void txtCedula_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
				FuncionesProcedimientos.AbreFacturacionClientes(5, this.txtCedula.Text.ToString(), false);
		}
	}
}
