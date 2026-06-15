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
	/// Descripción breve de frmRequerimientosDeSuministros1.
	/// </summary>
	public class frmRequerimientosDeSuministros1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;

		#region Variables
		bool bAsientoAutomatico = false;
		bool bImpreso = false;
		int idTipoPrecio = 0;
		bool BUsoNotaDeCredito = false;
		decimal dSaldoNotaDeCredito = 0;
		bool BRetencion = false;
		bool ModSerial = false;
		bool bEliminaAlValidar = false;
		bool bNuevo = false;
		bool bEdicion = false;
		decimal dAnticipo = 0;
		bool bUsoAnticipo = false;
		bool bPromocionCombo = false;
		bool bPromocion = false;
		bool bCombo = false;
		int IdSigno = 2;
		int idBloqueaTransacciones = 0;
		int idBloqueaTransacciones1 = 0;
		bool bSerialesL = false;
		bool bFacturacManual = false;
		bool bSerialNoExiste = false;
		bool bConsignacion = false;
		bool bIncluyeIva = false;
		private Acceso miAcceso;
		private bool bActualiza = true;
		bool bRastreo = false;
		bool bNumeracionAutomatica = true;
		DateTime dtFechPrimPago = DateTime.Today;
		bool bAnulado = false;
		int idTipoOrigen = 0;
		int iBodega = 0;
		int IdAP = 0;
		#endregion Variables

		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaAprobacion;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtUsuario;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotasAprobacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidArticulo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor idFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCedula;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnFacturar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidProyecto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstadoFactura;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientosDeSuministros1()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmRequerimientosDeSuministros1(int idAP)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			IdAP = idAP;

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientosDeSuministros1));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleSuministros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoSuministros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Existencia");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleSuministros");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoSuministros");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Existencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label4 = new System.Windows.Forms.Label();
			this.txtUsuario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNotasAprobacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.dtFechaAprobacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtBodega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstadoFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.idFormaPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCedula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnFacturar = new System.Windows.Forms.Button();
			this.txtidProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotasAprobacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAprobacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.idFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidProyecto)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 237;
			this.label6.Text = "Fecha";
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
			this.dtFecha.Location = new System.Drawing.Point(64, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(128, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 236;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 42);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(85, 17);
			this.lblBodega.TabIndex = 238;
			this.lblBodega.Text = "Centro de Costo";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(8, 75);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 17);
			this.lblVendedor.TabIndex = 241;
			this.lblVendedor.Text = "Vendedor";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbVendedor
			// 
			appearance2.FontData.Name = "Tahoma";
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance2;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 335;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(104, 72);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(256, 22);
			this.cmbVendedor.TabIndex = 240;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVendedor_KeyDown);
			this.cmbVendedor.ValueChanged += new System.EventHandler(this.cmbVendedor_ValueChanged);
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance3;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNotas.Location = new System.Drawing.Point(104, 104);
			this.txtNotas.MaxLength = 100;
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(488, 84);
			this.txtNotas.TabIndex = 831;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			this.txtNotas.ValueChanged += new System.EventHandler(this.txtNotas_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 138);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 17);
			this.label1.TabIndex = 832;
			this.label1.Text = "Notas";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Enabled = false;
			this.lblNumero.Location = new System.Drawing.Point(832, 11);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 17);
			this.lblNumero.TabIndex = 834;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance4;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(888, 8);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 833;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.txtUsuario);
			this.groupBox1.Controls.Add(this.txtNotasAprobacion);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.dtFechaAprobacion);
			this.groupBox1.Enabled = false;
			this.groupBox1.Location = new System.Drawing.Point(616, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(384, 168);
			this.groupBox1.TabIndex = 835;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Aprobación";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 51);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(42, 17);
			this.label4.TabIndex = 836;
			this.label4.Text = "Usuario";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtUsuario
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUsuario.Appearance = appearance5;
			this.txtUsuario.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUsuario.Enabled = false;
			this.txtUsuario.Location = new System.Drawing.Point(64, 48);
			this.txtUsuario.Name = "txtUsuario";
			this.txtUsuario.Size = new System.Drawing.Size(312, 22);
			this.txtUsuario.TabIndex = 835;
			// 
			// txtNotasAprobacion
			// 
			this.txtNotasAprobacion.AcceptsReturn = true;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotasAprobacion.Appearance = appearance6;
			this.txtNotasAprobacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotasAprobacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotasAprobacion.Enabled = false;
			this.txtNotasAprobacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNotasAprobacion.Location = new System.Drawing.Point(64, 80);
			this.txtNotasAprobacion.Multiline = true;
			this.txtNotasAprobacion.Name = "txtNotasAprobacion";
			this.txtNotasAprobacion.Size = new System.Drawing.Size(312, 80);
			this.txtNotasAprobacion.TabIndex = 834;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 100);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 17);
			this.label3.TabIndex = 833;
			this.label3.Text = "Notas";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 18);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 17);
			this.label2.TabIndex = 239;
			this.label2.Text = "Fecha";
			// 
			// dtFechaAprobacion
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaAprobacion.Appearance = appearance7;
			this.dtFechaAprobacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaAprobacion.DateButtons.Add(dateButton2);
			this.dtFechaAprobacion.Enabled = false;
			this.dtFechaAprobacion.Format = "dd/MM/yyyy";
			this.dtFechaAprobacion.Location = new System.Drawing.Point(64, 16);
			this.dtFechaAprobacion.Name = "dtFechaAprobacion";
			this.dtFechaAprobacion.NonAutoSizeHeight = 23;
			this.dtFechaAprobacion.Size = new System.Drawing.Size(168, 21);
			this.dtFechaAprobacion.SpinButtonsVisible = true;
			this.dtFechaAprobacion.TabIndex = 238;
			this.dtFechaAprobacion.Value = ((object)(resources.GetObject("dtFechaAprobacion.Value")));
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(624, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 836;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 208);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1024, 8);
			this.groupBox2.TabIndex = 837;
			this.groupBox2.TabStop = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance8.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance8;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 200;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance9;
			ultraGridColumn8.Format = "#,##0";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridColumn9.CellAppearance = appearance10;
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Width = 340;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance11;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 60;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Width = 281;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance13;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 10F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance14;
			appearance15.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance15.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.ultraGrid1.Enabled = false;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 224);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(992, 280);
			this.ultraGrid1.TabIndex = 838;
			this.ultraGrid1.Error += new Infragistics.Win.UltraWinGrid.ErrorEventHandler(this.ultraGrid1_Error);
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(bool);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(-8, 512);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1008, 8);
			this.groupBox3.TabIndex = 839;
			this.groupBox3.TabStop = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 528);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 24);
			this.btnNuevo.TabIndex = 845;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(86, 528);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 844;
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
			this.btnEditar.Location = new System.Drawing.Point(170, 528);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 24);
			this.btnEditar.TabIndex = 843;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(252, 528);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 842;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(334, 528);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 24);
			this.btnAnular.TabIndex = 841;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(497, 528);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 24);
			this.btnCancelar.TabIndex = 840;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			// txtBodega
			// 
			this.txtBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodega.Enabled = false;
			this.txtBodega.Location = new System.Drawing.Point(200, 8);
			this.txtBodega.Name = "txtBodega";
			this.txtBodega.PromptChar = ' ';
			this.txtBodega.Size = new System.Drawing.Size(16, 22);
			this.txtBodega.TabIndex = 850;
			this.txtBodega.Visible = false;
			// 
			// txtidCliente
			// 
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.Location = new System.Drawing.Point(264, 8);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(16, 22);
			this.txtidCliente.TabIndex = 849;
			this.txtidCliente.Value = 171570;
			this.txtidCliente.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(248, 8);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 848;
			this.txtEstado.Visible = false;
			// 
			// txtEstadoFactura
			// 
			this.txtEstadoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstadoFactura.Enabled = false;
			this.txtEstadoFactura.Location = new System.Drawing.Point(232, 8);
			this.txtEstadoFactura.Name = "txtEstadoFactura";
			this.txtEstadoFactura.PromptChar = ' ';
			this.txtEstadoFactura.Size = new System.Drawing.Size(16, 22);
			this.txtEstadoFactura.TabIndex = 847;
			this.txtEstadoFactura.Visible = false;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(216, 8);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtIdCompra.TabIndex = 846;
			this.txtIdCompra.Visible = false;
			// 
			// txtidArticulo
			// 
			this.txtidArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidArticulo.Enabled = false;
			this.txtidArticulo.Location = new System.Drawing.Point(280, 8);
			this.txtidArticulo.Name = "txtidArticulo";
			this.txtidArticulo.PromptChar = ' ';
			this.txtidArticulo.Size = new System.Drawing.Size(16, 22);
			this.txtidArticulo.TabIndex = 851;
			this.txtidArticulo.Visible = false;
			// 
			// idFormaPago
			// 
			this.idFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.idFormaPago.Enabled = false;
			this.idFormaPago.Location = new System.Drawing.Point(296, 8);
			this.idFormaPago.Name = "idFormaPago";
			this.idFormaPago.PromptChar = ' ';
			this.idFormaPago.Size = new System.Drawing.Size(16, 22);
			this.idFormaPago.TabIndex = 853;
			this.idFormaPago.Value = 9;
			this.idFormaPago.Visible = false;
			// 
			// txtCedula
			// 
			this.txtCedula.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCedula.Appearance = appearance18;
			this.txtCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCedula.Enabled = false;
			this.txtCedula.Location = new System.Drawing.Point(888, 528);
			this.txtCedula.MaxLength = 9;
			this.txtCedula.Name = "txtCedula";
			this.txtCedula.Size = new System.Drawing.Size(104, 22);
			this.txtCedula.TabIndex = 854;
			this.txtCedula.Visible = false;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.NullText = "248";
			ultraGridColumn14.Width = 248;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(104, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(256, 21);
			this.cmbBodega.TabIndex = 239;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodega_Validating);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.SystemColors.Control;
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(575, 529);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 855;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnFacturar
			// 
			this.btnFacturar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnFacturar.BackColor = System.Drawing.SystemColors.Control;
			this.btnFacturar.CausesValidation = false;
			this.btnFacturar.Enabled = false;
			this.btnFacturar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnFacturar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnFacturar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFacturar.ImageIndex = 62;
			this.btnFacturar.Location = new System.Drawing.Point(416, 529);
			this.btnFacturar.Name = "btnFacturar";
			this.btnFacturar.TabIndex = 856;
			this.btnFacturar.Text = "Facturar";
			this.btnFacturar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnFacturar.Click += new System.EventHandler(this.btnFacturar_Click);
			// 
			// txtidProyecto
			// 
			this.txtidProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidProyecto.Enabled = false;
			this.txtidProyecto.Location = new System.Drawing.Point(312, 8);
			this.txtidProyecto.Name = "txtidProyecto";
			this.txtidProyecto.PromptChar = ' ';
			this.txtidProyecto.Size = new System.Drawing.Size(16, 22);
			this.txtidProyecto.TabIndex = 857;
			this.txtidProyecto.Visible = false;
			// 
			// frmRequerimientosDeSuministros1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1008, 558);
			this.Controls.Add(this.txtidProyecto);
			this.Controls.Add(this.btnFacturar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtCedula);
			this.Controls.Add(this.idFormaPago);
			this.Controls.Add(this.txtidArticulo);
			this.Controls.Add(this.txtBodega);
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtEstadoFactura);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtFecha);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Location = new System.Drawing.Point(410, 0);
			this.Name = "frmRequerimientosDeSuministros1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Requerimientos De Suministros";
			this.Load += new System.EventHandler(this.frmRequerimientosDeSuministros1_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotasAprobacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaAprobacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.idFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidProyecto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.Close();
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			bNuevo = true;
			bEdicion = true;
			
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);
			CargaLocal();

			this.cmbBodega.Enabled = true;
			this.cmbVendedor.Enabled = true;	
//			this.btnNuevo.Enabled = false;
//			this.btnConsultar.Enabled = false;
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = false;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.txtNotas.Enabled = true;
			this.cmbVendedor.Enabled = true;
			this.txtEstado.Value = 1;
			this.lblEstado.Text = "PENDIENTE";
			this.cmbBodega.Focus ();

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid1, "idArticulo");
			bEliminaAlValidar = false;

			#region Validacion 
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha del Documento", true, "BODEGA", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una bodega")) return;
			if (!Validacion.vbComboVacio(this.cmbVendedor , "Seleccione una Vendedor")) return;
	
			if (!FuncionesInventario.ValidaBodegaEnInventario((int)this.txtBodega.Value, Convert.ToDateTime(this.dtFecha.Value), cdsSeteoF)) return;

			if (!FuncionesInventario.ValidaBodegaInventarios(this.ultraGrid1, (int)this.txtBodega.Value, Convert.ToDateTime(this.dtFecha.Value), cdsSeteoF)) return;
			
			//			if (!Validacion.vbGrids(this.ultraGrid1, "Codigo", 9)) return;		

			if (!Validacion.vbTexto(this.txtNotas, 5, 100, "Ingrese la sugerencia")) return;

			#endregion Validacion

			#region Cantidad

			int iCantidad =0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDTi in this.ultraGrid1.Rows.All)
			{
				if((int)drDTi.Cells["Cantidad"].Value == 0)	iCantidad++;
			}

			if(iCantidad > 0) 
			{
				MessageBox.Show("No Puede Ingresar Articulo en 0 " , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
				ultraGrid1.ActiveRow.Selected = true;
				return;
			}
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los Suministros que requiere.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}
			}
			#endregion Cantidad

			#region Existencia
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				int Existencia = Funcion.iEscalarSQL(cdsSeteoF,String.Format ("Select dbo.RetornaExistencia('{0}', {1}, {2}, {3}, {4}, {5})", 
					Convert.ToDateTime(dtFecha.Value).ToString("yyyyMMdd"), 0, iBodega, (int)dr.Cells["idArticulo"].Value, 65, 1));

				if (Existencia < 1)
				{
					MessageBox.Show("Este Articulo No Tiene Existencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
					ultraGrid1.ActiveRow = dr;
					ultraGrid1 .ActiveRow.Selected = true;
					return;
				}

				if ((int)dr.Cells["Cantidad"].Value > Existencia)
				{	
					MessageBox.Show(string.Format("Existencia Actual : {0}", Existencia), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					ultraGrid1.ActiveRow = dr;
					ultraGrid1 .ActiveRow.Selected = true;
					return;
				}
			}
			#endregion Existencia

			#region Guardar


			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;

				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					if (bNuevo && bNumeracionAutomatica)
					{
						string sSQLNumero = string.Format("Exec NumeracionLocales 65, {0}, 0", (int)this.cmbBodega.Value);
						oCmdActualiza.CommandText = sSQLNumero;
						this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString(); 
					}	

					#region Graba Maestro

					int IdBodega = 0;
					int IdVendedor = 0;
					
					FuncionesProcedimientos.RetornaFechaServidor(this.dtFechaAprobacion, cdsSeteoF, true);
			
					if (this.cmbBodega.ActiveRow != null) IdBodega = (int)this.cmbBodega.Value;
					if (this.cmbVendedor.ActiveRow != null) IdVendedor = (int)this.cmbVendedor.Value;
					if (this.dtFechaAprobacion.Value != null) this.dtFechaAprobacion.Value = DateTime.Today;
																																												 
					if(this.txtUsuario.Text.Equals(""))this.txtUsuario.Text="";
					if(this.txtNotasAprobacion.Text.Equals(""))this.txtNotasAprobacion.Text="";

					string  sSQL1 = string.Format("Exec GrabaRequerimientoSuministros {0}, '{1}', {2}, {3}, '{4}',{5},'{6}','{7}','{8}',{9}" ,
						(int)this.txtIdCompra.Value,
						this.txtNumero.Text,
						IdBodega,
						IdVendedor,
						this.txtNotas.Text,
						(int)txtEstado.Value,
						this.txtUsuario.Text,
						this.txtNotasAprobacion.Text,
						Convert.ToDateTime(this.dtFechaAprobacion.Value).ToString("yyyyMMdd"),
						(int)this.txtidProyecto.Value);
					oCmdActualiza.CommandText = sSQL1;
					this.txtIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro

					#region Detalle
					
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{	
						if(dr.Cells["Notas"].Value == System.DBNull.Value )dr.Cells["Notas"].Value =" ";
						if(dr.Cells["Estado"].Value == System.DBNull.Value )dr.Cells["Estado"].Value = 0;

						string sSQLDetalle = string.Format("Exec GrabaDetRequerimientoSuministros {0}, {1}, {2}, {3},{4},{5},'{6}'",
							(int)dr.Cells["idDetalleSuministros"].Value,
							(int)this.txtIdCompra.Value,
							(int)dr.Cells["idArticulo"].Value, 
							(int)dr.Cells["Cantidad"].Value,
							(int)dr.Cells["Existencia"].Value,
							Convert.ToInt32(dr.Cells["Estado"].Value),
							Convert.ToString(dr.Cells ["Notas"].Value.ToString()));

						oCmdActualiza.CommandText = sSQLDetalle;
						oCmdActualiza.ExecuteNonQuery();
					}

					string sSQLActDetalle = string.Format("Exec ConsultaDetalleSuministros {0}", (int)this.txtIdCompra.Value);
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle = new DataTable();
					oDADetalle.Fill(oDTDetalle);
					this.ultraGrid1.DataSource = oDTDetalle;
			
					#endregion Detalle	

					#region Bloqueo
					if (!bNuevo)
					{
						string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtIdCompra.Value);
						oCmdActualiza.CommandText = sSQLBloqueo;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Bloqueo

					oTransaction.Commit();
					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
					#region Controles
					
					this.cmbBodega.Enabled = false;
					this.btnGuardar.Enabled = false;
					this.cmbVendedor.Enabled = false;	
					this.btnNuevo.Enabled = true;
					this.btnConsultar.Enabled = true;
					this.txtNumero.Enabled = false;
					this.dtFecha.Enabled = false;
					this.txtNotas.Enabled = false;
					FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
					bNuevo = false;
					bEdicion = false;
					//					this.ConsultaRegistro((int)this.txtIdCompra.Value);
					#endregion Controles
				}
				catch(Exception ex)
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
				#endregion Guardar
		}

		#region Controles
		private void Botones()
		{
		this.txtBodega.Value = 0;
		this.txtidArticulo.Value = 0;
		this.txtIdCompra.Value = 0;
		this.txtEstado.Value = 0;
		this.txtNumero.Text = "";	
		this.txtNotas.Text = "";	
		this.lblEstado.Text = "";
		this.txtUsuario.Text =  "";
		this.txtNotasAprobacion.Text = "";
		this.txtidProyecto.Value = 0;
		//			this.dtFecha.Value = System.DBNull.Value; 
		this.dtFechaAprobacion.Value = System.DBNull.Value; 

		this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleSuministros  0");

		this.cmbBodega.Enabled = false;
		this.cmbVendedor.Enabled = false;	
		this.cmbBodega.Value = System.DBNull.Value;
		this.cmbVendedor.Value = System.DBNull.Value;
		this.txtNotas.Enabled = false;	

		FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
	
	}
		#endregion Controles

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

      Botones();

//			this.btnNuevo.Enabled = true;
//			this.btnConsultar.Enabled = true;
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;		
	
			bAnulado = false;
			bEliminaAlValidar = false;
		
			this.btnNuevo.Focus();
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Codigo Articulo
			if (!bActualiza) return;
			string sCodigo = "";
			if (e.Cell.Column.ToString() == "Codigo")
			{
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{					
					string sSQLArticulo = string.Format("Exec ConsultaInformacionArticulosCombos '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idDetalleSuministros"].Value = 0;
						e.Cell.Row.Cells["idRequerimientoSuministros"].Value = 0;
						e.Cell.Row.Cells["idArticulo"].Value = drArticulo.GetInt32(0);
						sCodigo = drArticulo.GetString(1);
						e.Cell.Row.Cells["Cantidad"].Value = 1;	
						e.Cell.Row.Cells["Articulo"].Value = drArticulo.GetString(6);
					}
					drArticulo.Close();
					bActualiza = false;
					e.Cell.Row.Cells["Codigo"].Value = sCodigo;
					bActualiza = true;
					string sSQLocal = string.Format("Exec RSRetornaBodega {0}", (int)this.cmbBodega.Value);
					iBodega= Funcion.iEscalarSQL(cdsSeteoF, sSQLocal);	

					e.Cell.Row.Cells["Existencia"].Value = Funcion.iEscalarSQL(cdsSeteoF,String.Format ("Select dbo.RetornaExistencia('{0}', {1}, {2}, {3}, {4},{5})", 
						Convert.ToDateTime(dtFecha.Value).ToString("yyyyMMdd"), 0, iBodega, (int)e.Cell.Row.Cells["idArticulo"].Value,65,1));
				}
				else if (iCuenta != 1)
				{
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODIGO", iBodega, 65, 1, DateTime.Today);					
				}
			}
			#endregion Codigo Articulo
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				if(!bActualiza)
				{
					e.Row.Cells["idDetalleSuministros"].Value = 0;
					e.Row.Cells["idRequerimientoSuministros"].Value = 0;
					e.Row.Cells["idArticulo"].Value = 0;
					e.Row.Cells["Codigo"].Value = "";
					e.Row.Cells["Articulo"].Value = "";
					e.Row.Cells["Cantidad"].Value = 1;	
					e.Row.Cells["Existencia"].Value = 0;
				}
			}
		}
		private void CargaLocal()
		{
			string sSQLoc = string.Format("Exec RSListaCentrosDeCosto '{0}','{1}'",MenuLatinium.stUsuario,Convert.ToDateTime(dtFecha.Value).ToString("yyyyMMdd"));	
			this.txtBodega.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLoc);
			if ((int)this.txtBodega.Value > 0) 
			{
//				this.cmbBodega.Enabled = true;
				this.cmbBodega.Value = (int)this.txtBodega.Value;
					this.txtidProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select isNull(idProyecto, 0) From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));
			}
		}
		
		private void ListaBodegas (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, string Usuario, DateTime dtDesde, C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec RSListaCentrosDeCosto '{0}','{1}'",Usuario,dtDesde.ToString("yyyyMMdd")));
		}

		private void ListaPersonal (Infragistics.Win.UltraWinGrid.UltraCombo uCombo,int idProyecto,   DateTime dtDesde, string Usuario, C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec RSListaPersonal {0}, '{1}','{2}'", idProyecto,  dtDesde.ToString("yyyyMMdd"),Usuario));
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{

		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			#region Inventario
			if (e.KeyCode == Keys.F3)
			{				
				if (this.ultraGrid1.DisplayLayout.Override.AllowUpdate == Infragistics.Win.DefaultableBoolean.False) return;
				
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO")
				{
					string sSQLocal = string.Format("Exec RSRetornaBodega {0}", (int)this.cmbBodega.Value);
					iBodega= Funcion.iEscalarSQL(cdsSeteoF, sSQLocal);	

					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODIGO", iBodega, 65, 1, DateTime.Today);										
				}
			}
			#endregion Inventario

			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CANTIDAD")
				Funcion.CamposNumericos(sender, e);
		}

		private void ultraGrid1_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e)
		{
			if (e.MaskErrorInfo == null)
			{
				e.Cancel = true;
				string stError = e.DataErrorInfo.InvalidValue.ToString();
				string stCol = e.DataErrorInfo.Cell.Column.ToString();
				string stMensaje = "Error Detalle: VENTAS";
				
				string stAudita = "Exec AuditaCrear 53, 7, '" + stMensaje + "'"; // Se graba como salida de pantalla
				Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);

				MessageBox.Show("Valor: " + stError + " no valido en columna: " + stCol + "\n1.- GRABE LOS DATOS. \n2.- SALGA DE LA PANTALLA.\n3.- VUELVA A INGRESAR", 
					"Informacion de Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{			
			if (this.dtFecha.Value != System.DBNull.Value)
			{
				ListaBodegas(this.cmbBodega, MenuLatinium.stUsuario,(DateTime)this.dtFecha.Value,cdsSeteoF);
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)
			{
				//				this.cmbVendedor.Enabled = true;
				bNuevo = true ;
				if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;
				ListaPersonal(this.cmbVendedor,(int)this.cmbBodega.Value,(DateTime)this.dtFecha.Value,MenuLatinium.stUsuario, cdsSeteoF);
			}
		}

		private void cmbBodega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;
		}

		private void cmbVendedor_ValueChanged(object sender, System.EventArgs e)
		{
			//			if (this.cmbVendedor.ActiveRow != null)
			//			{
			//				this.txtNotas.Enabled = true;
			//			}
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
				bEdicion = true;
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if (!bEliminaAlValidar) 
				{
					if (!FuncionesInventario.ValidaEliminaFilasGrids(sender, e, 9, this.ultraGrid1, cdsSeteoF)) return;		
				}									
				else e.DisplayPromptMsg = false;			
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
		
		private void frmRequerimientosDeSuministros1_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0222");

			if (!Funcion.Permiso("1274", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Solicitudes Suministros", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
						if (!miAcceso.Nuevo) this.btnNuevo.Enabled = false;
						if (!miAcceso.Buscar) this.btnConsultar.Enabled = false;



			if (IdAP > 0) this.ConsultaRegistro(IdAP);
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

			string sSQLoc = string.Format("Select Ruc from Cliente where idCliente = {0}",(int)this.txtidCliente.Value);	
			this.txtCedula.Text = Funcion.sEscalarSQL(cdsSeteoF, sSQLoc);
		}

		private void ConsultaRegistro (int idCompra)
		{
			#region botones 
//			this.btnNuevo.Enabled = true;
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.cmbBodega.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.txtNotas.Enabled = false;

			#endregion botones

			this.txtIdCompra.Value = idCompra;
			int EstadoFac = 0; 

			string sSQL = string.Format("Exec ConsultaIndividualSuministros {0}", this.txtIdCompra.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdCompra.Value = dr.GetInt32(0);
				if (dr.GetValue(1) != System.DBNull.Value) this.dtFecha.Value = dr.GetDateTime(1);
				if (dr.GetValue(2) != System.DBNull.Value) this.txtNumero.Text = dr.GetString(2);
				this.cmbBodega.Value = dr.GetInt32(3);
				this.cmbVendedor.Value = dr.GetInt32(4);
				this.txtNotas.Text = dr.GetString(5);
				this.txtEstado.Value = dr.GetInt32(6);
				this.txtNotasAprobacion.Text = dr.GetString(7);
				this.txtUsuario.Text = dr.GetString(8);
				if (dr.GetValue(9) != System.DBNull.Value) this.dtFechaAprobacion.Value = dr.GetDateTime(9);
				this.txtidProyecto.Value = dr.GetInt32(10);
				this.txtEstadoFactura.Value = dr.GetInt32(11);
			}
			dr.Close();
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleSuministros {0}", (int)this.txtIdCompra.Value));   
			
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

			if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";
			if ((int)this.txtEstado.Value == 2) this.lblEstado.Text = "ANULADO";	
			if ((int)this.txtEstado.Value == 3){
				if((int)this.txtEstadoFactura.Value > 0){
					string sSQEst = string.Format("Select Estado from Compra where idCompra = {0}",(int)this.txtEstadoFactura.Value);	
					EstadoFac = Funcion.iEscalarSQL(cdsSeteoF, sSQEst);
					if(EstadoFac == 6)this.lblEstado.Text = "TEMPORAL";
					if(EstadoFac == 10)this.lblEstado.Text = "FACTURADO";
				}
			}	
			if ((int)this.txtEstado.Value == 5) this.lblEstado.Text = "APROBADO";
			if ((int)this.txtEstado.Value == 4) this.lblEstado.Text = "RECHAZADO";			

			if ((int)this.txtEstado.Value == 5) 
			{
				if (miAcceso.BFacturar) this.btnFacturar.Enabled = true;
			};
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbVendedor.Focus();
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNotas.Focus();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaSTSuministros Busqueda = new frmBuscaSTSuministros((int)this.txtIdCompra.Value))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					ConsultaRegistro((int)Busqueda.ultraGrid1.ActiveRow.Cells["idRequerimientoSuministros"].Value);
//					CargaLocal();
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{			
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("No puede editar el documento Procesado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else if ((int)this.txtEstado.Value == 5)
			{
				MessageBox.Show("No puede editar el documento esta Aprobado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("No puede editar el documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else if ((int)this.txtEstado.Value == 9)
			{
				MessageBox.Show("No puede editar el documento esta Facturado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else 
			{				
				#region botones
//				this.btnNuevo.Enabled = false;
//				this.btnConsultar.Enabled = false;
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = false;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = true;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = true;
				this.cmbBodega.Enabled = true;
				this.cmbVendedor.Enabled = true;
				this.txtNotas.Enabled = true;
				this.cmbBodega.Focus();
				#endregion Botones

				FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);			
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtIdCompra.Value);
				idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);	

				bNuevo = false;
				bEdicion = true;
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
		
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnFacturar_Click(object sender, System.EventArgs e)
		{			
			#region Graba Factura
			int idCompra = 0;

			if (DialogResult.Yes == MessageBox.Show("Esta seguro de GENERAR LA FACTURA?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{
				int iBodega = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Bodega From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));

				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 120;
						
					try
					{							
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;
					
						#region Maestro
						string sSQLMaestro = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, '{46}', {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, '{59}', {60}, {61}, '{62}', '{63}', {64}, '{65}', {66}, {67}, '{68}', {69}", 
							0, 0, 0, 1, 15, 6, false, 2, 6, iBodega,(int)this.txtidProyecto.Value , 205, 
							"", this.txtNumero.Text.ToString(), 4, 171570, "1791774582001", "SUPERMERCADO DE COMPUTADORAS COMPUBUSSINES CIA LTD", DateTime.Today.ToString("yyyyMMdd"), 
							"", "", DateTime.Today.ToString("yyyyMMdd"), false, DateTime.Today.ToString("yyyyMMdd"), DateTime.Today.ToString("yyyyMMdd hh:mm"), 0, "FACTURA GENERADA POR REQUERIMIENTO DE SUMINISTROS",
							false, 0, 0, 0, DateTime.Today.ToString("yyyyMMdd"), 0, 0, 0, 0, 0, 0,	
							false, false, false, false, false, true, iBodega, 0, "", 0,
							0, 0, 0, 0, false, 0, 0, 0, 0, 0, 0, "", false, 0, "", "", 0, "", 0, 0, "", 0, 0, "");
						oCmdActualiza.CommandText = sSQLMaestro;
						idCompra = (int)oCmdActualiza.ExecuteScalar();
						#endregion Maestro

						#region Detalle 	
						string sSQLDetalle = string.Format("Exec GuardaDetalleFacturarequerimientos {0}, {1}", 
							(int)this.txtIdCompra.Value, idCompra);
						oCmdActualiza.CommandText = sSQLDetalle;
						oCmdActualiza.ExecuteNonQuery();
						#endregion Detalle						
						
						oTransaction.Commit();
						oConexion.Close();
					}
					catch (Exception ex)
					{
						oTransaction.Rollback();
						
						oConexion.Close();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);														
						return;
					}
					finally
					{
						oConexion.Close();					

						this.btnFacturar.Enabled = false;

						MessageBox.Show("Factura generada correctamente en Estado Temporal", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
						//
						//						frmVenta Factura = new frmVenta(idCompra, true);
						//						Factura.MdiParent = this;
						//						Factura.Show();
						
					}
				}
			}
			#endregion Graba Factura
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void txtNotas_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

	}
}
