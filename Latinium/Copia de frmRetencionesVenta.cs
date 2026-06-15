using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRetencionesVenta.
	/// </summary>
	public class frmRetencionesVenta : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtCaducidad;
		private System.Windows.Forms.GroupBox groupBox1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label5;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetenidoRenta;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetenidoIVA;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblSubTotal;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetenido;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalFactura;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdRetencion;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFactura;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtAutFactura;
		private System.Windows.Forms.Label label31;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblNumero;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieFact;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.ComponentModel.IContainer components;

		Infragistics.Win.UltraWinGrid.UltraGrid UGridVentas;
		private Acceso miAcceso;
		int IdCompra = 0;
		int IdCliente = 0;
		string Factura = "";
		DateTime DtFecha = DateTime.Today;
		decimal DTotal = 0;		
		bool bEliminaAlValidar = false;
		int ICont = 0;
		int IdRetencion = 0;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		
		public frmRetencionesVenta(Infragistics.Win.UltraWinGrid.UltraGrid uGridVentas, int idCompra, int idCliente, string factura, DateTime dtFecha, decimal dTotal, int iCont)
		{
			InitializeComponent();
			
			UGridVentas = uGridVentas;
			IdCompra = idCompra;
			IdCliente = idCliente;
			Factura = factura;
			DtFecha = dtFecha;
			DTotal = dTotal;
			ICont = iCont;
		}

		public frmRetencionesVenta(int idRetencion)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			IdRetencion = idRetencion;
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmRetencionesVenta()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRetencionesVenta));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVARenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencionIVARenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Base");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVARenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencionIVARenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Base");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencionVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencionVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			this.label7 = new System.Windows.Forms.Label();
			this.dtCaducidad = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.label11 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtRetenidoRenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtRetenidoIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.txtRetenido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdRetencion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtAutFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSerieFact = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			((System.ComponentModel.ISupportInitialize)(this.dtCaducidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(639, 6);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(58, 16);
			this.label7.TabIndex = 382;
			this.label7.Text = "Caducidad";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtCaducidad
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtCaducidad.Appearance = appearance1;
			this.dtCaducidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtCaducidad.DateButtons.Add(dateButton1);
			this.dtCaducidad.Format = "dd/MM/yyyy";
			this.dtCaducidad.Location = new System.Drawing.Point(703, 6);
			this.dtCaducidad.Name = "dtCaducidad";
			this.dtCaducidad.NonAutoSizeHeight = 23;
			this.dtCaducidad.Size = new System.Drawing.Size(112, 21);
			this.dtCaducidad.SpinButtonsVisible = true;
			this.dtCaducidad.TabIndex = 3;
			this.dtCaducidad.Value = ((object)(resources.GetObject("dtCaducidad.Value")));
			this.dtCaducidad.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtCaducidad_KeyPress);
			this.dtCaducidad.Validating += new System.ComponentModel.CancelEventHandler(this.dtCaducidad_Validating);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-1, 344);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(826, 8);
			this.groupBox1.TabIndex = 378;
			this.groupBox1.TabStop = false;
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance2;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(7, 240);
			this.txtNotas.MaxLength = 600;
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(408, 88);
			this.txtNotas.TabIndex = 8;
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(529, 240);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 21);
			this.txtIdAsiento.TabIndex = 376;
			this.txtIdAsiento.Visible = false;
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(168, 368);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 11;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.CausesValidation = false;
			this.btnAceptar.Image = ((System.Drawing.Image)(resources.GetObject("btnAceptar.Image")));
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 368);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(72, 23);
			this.btnAceptar.TabIndex = 9;
			this.btnAceptar.Text = "&Aceptar";
			this.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(88, 368);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 10;
			this.btnConsultar.Text = "&Cancelar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// label11
			// 
			this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(7, 216);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(37, 16);
			this.label11.TabIndex = 372;
			this.label11.Text = "Notas:";
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(569, 360);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 371;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(664, 272);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(34, 16);
			this.label5.TabIndex = 369;
			this.label5.Text = "Renta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRetenidoRenta
			// 
			this.txtRetenidoRenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetenidoRenta.Appearance = appearance3;
			this.txtRetenidoRenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetenidoRenta.Enabled = false;
			this.txtRetenidoRenta.FormatString = "#,##0.00";
			this.txtRetenidoRenta.Location = new System.Drawing.Point(720, 272);
			this.txtRetenidoRenta.Name = "txtRetenidoRenta";
			this.txtRetenidoRenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetenidoRenta.PromptChar = ' ';
			this.txtRetenidoRenta.Size = new System.Drawing.Size(96, 21);
			this.txtRetenidoRenta.TabIndex = 370;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(664, 240);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(32, 16);
			this.label4.TabIndex = 367;
			this.label4.Text = "I.V.A.";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRetenidoIVA
			// 
			this.txtRetenidoIVA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetenidoIVA.Appearance = appearance4;
			this.txtRetenidoIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetenidoIVA.Enabled = false;
			this.txtRetenidoIVA.FormatString = "#,##0.00";
			this.txtRetenidoIVA.Location = new System.Drawing.Point(720, 240);
			this.txtRetenidoIVA.Name = "txtRetenidoIVA";
			this.txtRetenidoIVA.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetenidoIVA.PromptChar = ' ';
			this.txtRetenidoIVA.Size = new System.Drawing.Size(96, 21);
			this.txtRetenidoIVA.TabIndex = 368;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(664, 304);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(29, 16);
			this.label2.TabIndex = 366;
			this.label2.Text = "Total";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(401, 38);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(71, 16);
			this.lblSubTotal.TabIndex = 363;
			this.lblSubTotal.Text = "Total Factura";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRetenido
			// 
			this.txtRetenido.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRetenido.Appearance = appearance5;
			this.txtRetenido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetenido.Enabled = false;
			this.txtRetenido.FormatString = "#,##0.00";
			this.txtRetenido.Location = new System.Drawing.Point(720, 304);
			this.txtRetenido.Name = "txtRetenido";
			this.txtRetenido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetenido.PromptChar = ' ';
			this.txtRetenido.Size = new System.Drawing.Size(96, 21);
			this.txtRetenido.TabIndex = 365;
			// 
			// txtTotalFactura
			// 
			this.txtTotalFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalFactura.Appearance = appearance6;
			this.txtTotalFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalFactura.Enabled = false;
			this.txtTotalFactura.FormatString = "#,##0.00";
			this.txtTotalFactura.Location = new System.Drawing.Point(489, 35);
			this.txtTotalFactura.Name = "txtTotalFactura";
			this.txtTotalFactura.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalFactura.PromptChar = ' ';
			this.txtTotalFactura.Size = new System.Drawing.Size(80, 21);
			this.txtTotalFactura.TabIndex = 6;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(545, 240);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 362;
			this.txtEstado.Visible = false;
			// 
			// txtIdRetencion
			// 
			this.txtIdRetencion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdRetencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdRetencion.Enabled = false;
			this.txtIdRetencion.Location = new System.Drawing.Point(497, 240);
			this.txtIdRetencion.Name = "txtIdRetencion";
			this.txtIdRetencion.PromptChar = ' ';
			this.txtIdRetencion.Size = new System.Drawing.Size(16, 21);
			this.txtIdRetencion.TabIndex = 361;
			this.txtIdRetencion.Visible = false;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(513, 240);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 21);
			this.txtIdCompra.TabIndex = 360;
			this.txtIdCompra.Visible = false;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(209, 38);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 359;
			this.label1.Text = "Factura";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFactura
			// 
			this.txtFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFactura.Appearance = appearance7;
			this.txtFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFactura.Enabled = false;
			this.txtFactura.Location = new System.Drawing.Point(265, 35);
			this.txtFactura.Name = "txtFactura";
			this.txtFactura.Size = new System.Drawing.Size(112, 21);
			this.txtFactura.TabIndex = 5;
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
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 74;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance9;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 103;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 102;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 381;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance10;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 101;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance11;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 102;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance13;
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
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(7, 72);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(810, 136);
			this.ultraGrid1.TabIndex = 7;
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.Double);
			ultraDataColumn5.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(System.Double);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn9.DataType = typeof(System.Double);
			ultraDataColumn9.DefaultValue = 0;
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
			// txtAutFactura
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAutFactura.Appearance = appearance18;
			this.txtAutFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutFactura.Location = new System.Drawing.Point(391, 6);
			this.txtAutFactura.MaxLength = 37;
			this.txtAutFactura.Name = "txtAutFactura";
			this.txtAutFactura.Size = new System.Drawing.Size(232, 21);
			this.txtAutFactura.TabIndex = 2;
			this.txtAutFactura.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAutFactura_KeyPress);
			this.txtAutFactura.Validating += new System.ComponentModel.CancelEventHandler(this.txtAutFactura_Validating);
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(311, 6);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(67, 16);
			this.label31.TabIndex = 355;
			this.label31.Text = "Autorización";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(7, 6);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(31, 16);
			this.label3.TabIndex = 354;
			this.label3.Text = "Serie";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(143, 6);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 353;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance19;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(199, 6);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(96, 21);
			this.txtNumero.TabIndex = 1;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			// 
			// txtSerieFact
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerieFact.Appearance = appearance20;
			this.txtSerieFact.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerieFact.Location = new System.Drawing.Point(55, 6);
			this.txtSerieFact.MaxLength = 6;
			this.txtSerieFact.Name = "txtSerieFact";
			this.txtSerieFact.Size = new System.Drawing.Size(64, 21);
			this.txtSerieFact.TabIndex = 0;
			this.txtSerieFact.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerieFact_KeyPress);
			this.txtSerieFact.Validating += new System.ComponentModel.CancelEventHandler(this.txtSerieFact_Validating);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(7, 38);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 350;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance21;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton2);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(55, 36);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 4;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
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
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance22.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance22;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Width = 161;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 93;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 118;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Width = 99;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance23;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance25.FontData.BoldAsString = "True";
			appearance25.FontData.Name = "Arial";
			appearance25.FontData.SizeInPoints = 10F;
			appearance25.ForeColor = System.Drawing.Color.White;
			appearance25.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance25;
			appearance26.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance26.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance28;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(232, 144);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(584, 50);
			this.ultraGrid2.TabIndex = 383;
			this.ultraGrid2.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.Double);
			ultraDataColumn13.DefaultValue = 0;
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.ImageList = this.imageList1;
			this.btnAsiento.Location = new System.Drawing.Point(248, 368);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 384;
			this.btnAsiento.Text = "Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// frmRetencionesVenta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(824, 400);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtCaducidad);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtRetenidoRenta);
			this.Controls.Add(this.txtRetenidoIVA);
			this.Controls.Add(this.txtRetenido);
			this.Controls.Add(this.txtTotalFactura);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdRetencion);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtFactura);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtAutFactura);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtSerieFact);
			this.Controls.Add(this.dtFecha);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmRetencionesVenta";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Retención";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRetencionesVenta_KeyDown);
			this.Load += new System.EventHandler(this.frmRetencionesVenta_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtCaducidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenidoIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetenido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieFact)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void Suma()
		{
			this.txtRetenidoIVA.Value = 0.00m;
			this.txtRetenidoRenta.Value = 0.00m;
			this.txtRetenido.Value = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (!(bool) dr.Cells["IVARenta"].Value) this.txtRetenidoIVA.Value = Convert.ToDecimal(this.txtRetenidoIVA.Value) + Convert.ToDecimal(dr.Cells["Valor"].Value);
				if ((bool) dr.Cells["IVARenta"].Value) this.txtRetenidoRenta.Value = Convert.ToDecimal(this.txtRetenidoRenta.Value) + Convert.ToDecimal(dr.Cells["Valor"].Value);
			}

			this.txtRetenido.Value = Convert.ToDecimal(this.txtRetenidoIVA.Value) + Convert.ToDecimal(this.txtRetenidoRenta.Value);
		}

		private void frmRetencionesVenta_Load(object sender, System.EventArgs e)
		{
			#region Factura De Venta
			string stTipo = "040602";
					
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			this.dtFecha.CalendarInfo.MinDate = DtFecha;
			this.dtFecha.CalendarInfo.MaxDate = DtFecha.AddDays(5);

			if (IdRetencion == 0)
			{
				this.txtIdCompra.Value = IdCompra;				
				this.txtFactura.Text = Factura;
				this.txtTotalFactura.Value = DTotal;
				string sFactura = "";

				if (ICont > 1)
				{
					this.dtFecha.Value = DtFecha;

					#region Parametros por default
					this.txtIdRetencion.Value = 0;				
					this.txtEstado.Value = 0;					
					#endregion Parametros por default 				

					#region Variables
					int idRetIVARenta = 0;
					bool bIVARenta = false;
					decimal dPorc = 0.00m;
					string sCodigo = "";
					string sDetalle = "";
					decimal dBase = 0.00m;	
					decimal dRetenido = 0.00m;
					decimal dImpuesto = 0.00m;
					decimal dSubtotal = 0.00m;
					decimal dDescuento = 0.00m;
					decimal dIVA = 0.00m;
					#endregion Variables

					#region Retorna Valores de Retencion
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in UGridVentas.Rows.All)
					{
						if ((bool)dr.Cells["Sel"].Value)
						{														
							//string sSQLRM = string.Format("Exec CalculaRetencion {0}, {1}, {2}, {3}, '{4}', {5}", (int)dr.Cells["idCompra"].Value, IdCliente, 0, DtFecha.ToString("yyyyMMdd"), 1);
							string sSQLRM = string.Format("Exec CalculaRetencion {0}, {1}", (int)dr.Cells["idCompra"].Value, IdCliente);
							SqlDataReader drRetencionMultiple = Funcion.rEscalarSQL(cdsSeteoF, sSQLRM, true);
						
							while (drRetencionMultiple.Read())
							{						
								bIVARenta = drRetencionMultiple.GetBoolean(2);
								idRetIVARenta = drRetencionMultiple.GetInt32(3);
								dPorc = drRetencionMultiple.GetDecimal(4);
								sCodigo = drRetencionMultiple.GetString(5);
								sDetalle = drRetencionMultiple.GetString(6);
								dBase = drRetencionMultiple.GetDecimal(7);	
								dRetenido = drRetencionMultiple.GetDecimal(8);
						
								#region Crea Detalle de retenciones
								int iContI = 0;
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in this.ultraGrid1.Rows.All)
								{
									if ((int)dr2.Cells["idRetencionIVARenta"].Value == idRetIVARenta && (bool) dr2.Cells["IVARenta"].Value == bIVARenta)	
									{				
										dr2.Cells["Base"].Value = Math.Round(Convert.ToDecimal(dr2.Cells["Base"].Value), 2) + Math.Round(dBase, 2);
										dr2.Cells["Valor"].Value = Math.Round(Convert.ToDecimal(dr2.Cells["Valor"].Value), 2) + Math.Round(dRetenido, 2);
										iContI ++;
									}
								}

								if (idRetIVARenta > 0 && iContI == 0)
								{
									this.ultraGrid1.Rows.Band.AddNew();
									this.ultraGrid1.ActiveRow.Cells["idDetalleRetencion"].Value = 0;
									this.ultraGrid1.ActiveRow.Cells["idRetencion"].Value = 0;
									this.ultraGrid1.ActiveRow.Cells["IVARENTA"].Value = bIVARenta;
									this.ultraGrid1.ActiveRow.Cells["idRetencionIVARenta"].Value = idRetIVARenta;
									this.ultraGrid1.ActiveRow.Cells["Porcentaje"].Value = dPorc;
									this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = sCodigo;
									this.ultraGrid1.ActiveRow.Cells["Detalle"].Value = sDetalle;
									this.ultraGrid1.ActiveRow.Cells["Base"].Value = Math.Round(dBase, 2);
									this.ultraGrid1.ActiveRow.Cells["Valor"].Value = dRetenido;	
								}
								#endregion Crea Detalle de retenciones

								#region Ingresa Valores a Retenciones de Venta
								int iContRV = 0;
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drRetVenta in this.ultraGrid2.Rows.All)
								{
									if ((int)drRetVenta.Cells["idCompra"].Value == (int)dr.Cells["idCompra"].Value)	
									{				
										drRetVenta.Cells["Retenido"].Value = Math.Round(Convert.ToDecimal(drRetVenta.Cells["Retenido"].Value), 2) + Math.Round(dRetenido, 2);
										iContRV ++;
									}
								}

								if (iContRV == 0)
								{
									this.ultraGrid2.Rows.Band.AddNew();
									this.ultraGrid2.ActiveRow.Cells["idRetencionVenta"].Value = 0;
									this.ultraGrid2.ActiveRow.Cells["idRetencion"].Value = 0;
									this.ultraGrid2.ActiveRow.Cells["idCompra"].Value = (int)dr.Cells["idCompra"].Value;
									this.ultraGrid2.ActiveRow.Cells["Retenido"].Value = dRetenido;
									this.ultraGrid2.ActiveRow.Cells["Estado"].Value = 0;								
								}
								#endregion Ingresa Valores a Retenciones de Venta						

								Suma();
							}
							drRetencionMultiple.Close();
				
							if (sFactura.Length == 0) sFactura = dr.Cells["Numero"].Value.ToString();
							else sFactura = sFactura + " - " + dr.Cells["Numero"].Value.ToString();												
						}
					}
					#endregion Retorna Valores de Retencion

					this.ultraGrid1.Rows.Band.AddNew();

					this.txtNotas.Text = sFactura;
				}
				else if (ICont == 1)
				{
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Retenciones Where Estado <> 2 And ClienteProveedor = 0 And idCompra = {0}", IdCompra)) == 0)
					{
						this.dtFecha.Value = DtFecha;

						#region Parametros por default
						this.txtIdRetencion.Value = 0;				
						this.txtEstado.Value = 0;					
						#endregion Parametros por default 				

						string sSQL = string.Format("Exec CalculaRetencion {0}, {1}", IdCompra, IdCliente);
						this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

						Suma();
					}
				}
			}
			else if (IdRetencion > 0)
			{
				string sSQLCRE = string.Format("Exec ConsultaRetencion {0}, 0", IdRetencion);
				SqlDataReader drConsulta = Funcion.rEscalarSQL(cdsSeteoF, sSQLCRE, true);
				drConsulta.Read();
				if (drConsulta.HasRows)
				{
					this.txtIdRetencion.Value = drConsulta.GetInt32(0);
					this.txtIdCompra.Value = drConsulta.GetInt32(1);
					this.txtIdAsiento.Value = drConsulta.GetInt32(2);
					this.txtEstado.Value = drConsulta.GetInt32(4);
					this.txtSerieFact.Text = drConsulta.GetString(5);
					this.txtNumero.Text = drConsulta.GetString(6);
					this.txtAutFactura.Text = drConsulta.GetString(7);
					this.dtCaducidad.Value = drConsulta.GetDateTime(8);
					this.dtFecha.Value = drConsulta.GetDateTime(9);
					this.txtNotas.Text = drConsulta.GetString(10);
					this.txtRetenidoIVA.Value = drConsulta.GetDecimal(11);
					this.txtRetenidoRenta.Value = drConsulta.GetDecimal(12);
					this.txtRetenido.Value = drConsulta.GetDecimal(13);
					if ((int)this.txtEstado.Value == 2)
					{
						this.lblEstado.Text = "ANULADO";
						this.btnAnular.Text = "ELIMINAR";
					}
				}
				drConsulta.Close();

				string sSQLCRD = string.Format("Exec ConsultaDetalleRetencion {0}", (int)this.txtIdRetencion.Value);
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLCRD);

				if ((int)this.txtEstado.Value != 2)
				{					
					if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;					
					if (miAcceso.Anular) this.btnAnular.Enabled = true;
				}
				else
				{
					if (miAcceso.Eliminar) this.btnAnular.Enabled = true;
				}
			}
			#endregion Factura De Venta
			
			this.txtNumero.Focus();
		}

		private void txtSerieFact_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtNumero.Focus();
		}

		private void txtSerieFact_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtSerieFact.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();
				e.Cancel = true;
				return;
			}

			if (this.txtSerieFact.Text.Length < 6)
			{
				MessageBox.Show("El número de serie es de 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtAutFactura.Focus();
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			char Pad = '0';
			this.txtNumero.Text = this.txtNumero.Text.ToString().PadLeft(9, Pad);
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtNumero.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				e.Cancel = true;
				return;
			}

			char Pad = '0';
			this.txtNumero.Text = this.txtNumero.Text.ToString().PadLeft(9, Pad);

			string sSQL = string.Format("Select COUNT(*) From Retenciones Where Autorizacion = '{0}' And Serie = '{1}' And Numero = '{2}' and idRetencion <> {3}", this.txtAutFactura.Text.ToString(), this.txtSerieFact.Text.ToString(), this.txtNumero.Text.ToString(), (int)this.txtIdRetencion.Value);
			if (Funcion.iEscalarSQL(cdsSeteoF, sSQL) > 0)
			{
				MessageBox.Show("Ya existe este numero de documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void txtAutFactura_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.dtCaducidad.Focus();
		}

		private void txtAutFactura_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.txtAutFactura.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAutFactura.Focus();
				e.Cancel = true;
				return;
			}

			if (this.txtAutFactura.Text.Length < 3 && this.txtAutFactura.Text.Length > 37)
			{
				MessageBox.Show("Longitud de autorización debe estar entre 3 y 37 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAutFactura.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void dtCaducidad_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFecha.Focus();
		}

		private void dtCaducidad_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.dtCaducidad.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha de caducidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtCaducidad.Focus();
				e.Cancel = true;
				return;
			}
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNotas.Focus();
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha del Documento", true, "VENTA", cdsSeteoF)) e.Cancel = true;			
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnAceptar.Focus();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value != 2)
			{
				if (DialogResult.No == MessageBox.Show(string.Format("¿Esta seguro de anular esta retención?."), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			
				string sSQLAnul = string.Format("Exec AnulaRetencion {0}, {1}, 0, 1", (int)this.txtIdRetencion.Value, 0);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAnul, true);

				MessageBox.Show("Retención Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
				this.txtEstado.Value = 2;
				this.lblEstado.Text = "ANULADO";
				this.btnAnular.Text = "Eliminar";
			}
			else
			{
				if (DialogResult.No == MessageBox.Show(string.Format("¿Esta seguro de Eliminar esta retención?."), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			
				string sSQLAnul = string.Format("Exec AnulaRetencion {0}, {1}, 1, 1", (int)this.txtIdRetencion.Value, (int)this.txtIdRetencion.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAnul, true);

				MessageBox.Show("Retención Eliminada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
				this.txtEstado.Value = 2;
				this.lblEstado.Text = "";

				this.btnConsultar_Click(sender, e);
			}
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			#region Validacion Serie
			if (this.txtSerieFact.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();
				return;
			}

			if (this.txtSerieFact.Text.Length < 6)
			{
				MessageBox.Show("El número de serie es de 6 digitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtSerieFact.Focus();
				return;
			}
			#endregion Validacion Serie

			#region Validacion Numero
			if (this.txtNumero.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número del documento.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumero.Focus();
				return;
			}
			#endregion Validacion Numero

			#region Validacion Autorizacion
			if (this.txtAutFactura.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el número de autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAutFactura.Focus();
				return;
			}

			if (this.txtAutFactura.Text.Length < 3 && this.txtAutFactura.Text.Length > 37)
			{
				MessageBox.Show("Longitud de autorización debe estar entre 3 y 37 caracteres.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAutFactura.Focus();
				return;
			}
			#endregion Validacion Autorizacion

			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha del Documento", true, "VENTA", cdsSeteoF)) return;			

			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha de caducidad", false, "", cdsSeteoF)) return;			
			
			#region Filas Vacias
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idRetencionIVARenta"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}
			}
			#endregion Filas Vacias

			this.DialogResult = DialogResult.OK;
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!bEliminaAlValidar) e.Cancel = true;				
			else e.DisplayPromptMsg = false;
		}

		private void frmRetencionesVenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
		
		}
	}
}
