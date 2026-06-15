using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmComprasEstado.
	/// </summary>
	public class frmComprasEstado : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescuento;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNuevoPrecio;
		private System.Windows.Forms.Label lblNuevoPrecio;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblEstadoCobranzas;
		private System.Windows.Forms.Label lblEstadoDT;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservacionesCobranzas;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservacionesDT;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNuevoCosto;
		private System.Windows.Forms.Label label4;		
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargen;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecioOferta;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservacionesCompras;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNuevoPrecioConIVA;
		private System.Windows.Forms.Label label8;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGarantia;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label5;

		public frmComprasEstado()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();			
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idDetalleIngresoRetiro = 0;
		string sEstadoProductoCOB = string.Empty;
		string sObservacionesCobranzas = string.Empty;
		string sEstadoProductoDT = string.Empty;
		int idDTEstadoDT = 0;
		string sObservacionesDT = string.Empty;
		int iEstadoCompras = 0;
		decimal dNuevoCosto = 0.00m;
		decimal dPrecioOferta = 0.00m;
		decimal dDescuento = 0.00m;
		decimal dNuevoPrecio = 0.00m;
		decimal dMargen = 0.00m;
		string sObservacionesCompras = string.Empty;
		bool bModifica = false;		
		
		public frmComprasEstado(int IdDetalleIngresoRetiro, string SEstadoProductoCOB, string SObservacionesCobranzas, string SEstadoProductoDT, int IdDTEstadoDT, string SObservacionesDT, int IEstadoCompras, 
			decimal DNuevoCosto, decimal DPrecioOferta, decimal DDescuento, decimal DNuevoPrecio, decimal DMargen, string SObservacionesCompras, bool BModifica)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			
			idDetalleIngresoRetiro = IdDetalleIngresoRetiro;
			sEstadoProductoCOB = SEstadoProductoCOB;
			sObservacionesCobranzas = SObservacionesCobranzas;
			sEstadoProductoDT = SEstadoProductoDT;
			idDTEstadoDT = IdDTEstadoDT;
			sObservacionesDT = SObservacionesDT;
			iEstadoCompras = IEstadoCompras;
			dNuevoCosto = DNuevoCosto;
			dPrecioOferta = DPrecioOferta;
			dDescuento = DDescuento;
			dNuevoPrecio = DNuevoPrecio;
			dMargen = DMargen;
			sObservacionesCompras = SObservacionesCompras;
			bModifica = BModifica;

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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Accesorio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasAutorizado");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Accesorio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasAutorizado");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			this.label7 = new System.Windows.Forms.Label();
			this.txtObservacionesCobranzas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDescuento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNuevoPrecio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblNuevoPrecio = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.txtObservacionesDT = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblEstadoCobranzas = new System.Windows.Forms.Label();
			this.txtMargen = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.lblEstadoDT = new System.Windows.Forms.Label();
			this.txtNuevoCosto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtPrecioOferta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtObservacionesCompras = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtNuevoPrecioConIVA = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.txtGarantia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesCobranzas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNuevoPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesDT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNuevoCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioOferta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesCompras)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNuevoPrecioConIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGarantia)).BeginInit();
			this.SuspendLayout();
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(9, 8);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(67, 16);
			this.label7.TabIndex = 680;
			this.label7.Text = "COBRANZA";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservacionesCobranzas
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservacionesCobranzas.Appearance = appearance1;
			this.txtObservacionesCobranzas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservacionesCobranzas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservacionesCobranzas.Enabled = false;
			this.txtObservacionesCobranzas.Location = new System.Drawing.Point(9, 32);
			this.txtObservacionesCobranzas.Multiline = true;
			this.txtObservacionesCobranzas.Name = "txtObservacionesCobranzas";
			this.txtObservacionesCobranzas.Size = new System.Drawing.Size(552, 66);
			this.txtObservacionesCobranzas.TabIndex = 679;
			// 
			// txtDescuento
			// 
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescuento.Appearance = appearance2;
			this.txtDescuento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescuento.Enabled = false;
			this.txtDescuento.FormatString = "#,##0.00";
			this.txtDescuento.Location = new System.Drawing.Point(113, 400);
			this.txtDescuento.MaxValue = 100;
			this.txtDescuento.MinValue = 0;
			this.txtDescuento.Name = "txtDescuento";
			this.txtDescuento.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescuento.PromptChar = ' ';
			this.txtDescuento.Size = new System.Drawing.Size(96, 22);
			this.txtDescuento.TabIndex = 684;
			this.txtDescuento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescuento_KeyDown);
			this.txtDescuento.Click += new System.EventHandler(this.txtDescuento_Click);
			this.txtDescuento.ValueChanged += new System.EventHandler(this.txtDescuento_ValueChanged);
			this.txtDescuento.Enter += new System.EventHandler(this.txtDescuento_Enter);
			// 
			// txtNuevoPrecio
			// 
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNuevoPrecio.Appearance = appearance3;
			this.txtNuevoPrecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNuevoPrecio.Enabled = false;
			this.txtNuevoPrecio.FormatString = "#,##0.00";
			this.txtNuevoPrecio.Location = new System.Drawing.Point(113, 424);
			this.txtNuevoPrecio.MaxValue = 10000;
			this.txtNuevoPrecio.MinValue = -10000;
			this.txtNuevoPrecio.Name = "txtNuevoPrecio";
			this.txtNuevoPrecio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNuevoPrecio.PromptChar = ' ';
			this.txtNuevoPrecio.Size = new System.Drawing.Size(96, 22);
			this.txtNuevoPrecio.TabIndex = 686;
			this.txtNuevoPrecio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNuevoPrecio_KeyDown);
			this.txtNuevoPrecio.Click += new System.EventHandler(this.txtNuevoPrecio_Click);
			this.txtNuevoPrecio.ValueChanged += new System.EventHandler(this.txtNuevoPrecio_ValueChanged);
			this.txtNuevoPrecio.Enter += new System.EventHandler(this.txtNuevoPrecio_Enter);
			// 
			// lblNuevoPrecio
			// 
			this.lblNuevoPrecio.AutoSize = true;
			this.lblNuevoPrecio.BackColor = System.Drawing.Color.Transparent;
			this.lblNuevoPrecio.Location = new System.Drawing.Point(9, 424);
			this.lblNuevoPrecio.Name = "lblNuevoPrecio";
			this.lblNuevoPrecio.Size = new System.Drawing.Size(85, 17);
			this.lblNuevoPrecio.TabIndex = 687;
			this.lblNuevoPrecio.Text = "Precio Sin I.V.A.";
			this.lblNuevoPrecio.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.BackColor = System.Drawing.Color.Transparent;
			this.lblDescuento.Location = new System.Drawing.Point(9, 400);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(47, 17);
			this.lblDescuento.TabIndex = 685;
			this.lblDescuento.Text = "% Dscto";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(88, 538);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(75, 24);
			this.btnCancelar.TabIndex = 689;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Enabled = false;
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(8, 538);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(75, 24);
			this.btnAceptar.TabIndex = 688;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 528);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(570, 8);
			this.groupBox1.TabIndex = 690;
			this.groupBox1.TabStop = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(9, 112);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(153, 16);
			this.label1.TabIndex = 692;
			this.label1.Text = "DEPARTAMENTO TÉCNICO";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservacionesDT
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservacionesDT.Appearance = appearance4;
			this.txtObservacionesDT.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservacionesDT.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservacionesDT.Enabled = false;
			this.txtObservacionesDT.Location = new System.Drawing.Point(8, 136);
			this.txtObservacionesDT.Multiline = true;
			this.txtObservacionesDT.Name = "txtObservacionesDT";
			this.txtObservacionesDT.Size = new System.Drawing.Size(552, 66);
			this.txtObservacionesDT.TabIndex = 691;
			// 
			// lblEstadoCobranzas
			// 
			this.lblEstadoCobranzas.AutoSize = true;
			this.lblEstadoCobranzas.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstadoCobranzas.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstadoCobranzas.Location = new System.Drawing.Point(272, 6);
			this.lblEstadoCobranzas.Name = "lblEstadoCobranzas";
			this.lblEstadoCobranzas.Size = new System.Drawing.Size(0, 20);
			this.lblEstadoCobranzas.TabIndex = 693;
			this.lblEstadoCobranzas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMargen
			// 
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargen.Appearance = appearance5;
			this.txtMargen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargen.Enabled = false;
			this.txtMargen.FormatString = "#,##0.00";
			this.txtMargen.Location = new System.Drawing.Point(113, 472);
			this.txtMargen.MaxValue = 10000;
			this.txtMargen.MinValue = -10000;
			this.txtMargen.Name = "txtMargen";
			this.txtMargen.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargen.PromptChar = ' ';
			this.txtMargen.Size = new System.Drawing.Size(96, 22);
			this.txtMargen.TabIndex = 694;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(9, 472);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(41, 17);
			this.label2.TabIndex = 695;
			this.label2.Text = "Margen";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstadoDT
			// 
			this.lblEstadoDT.AutoSize = true;
			this.lblEstadoDT.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstadoDT.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstadoDT.Location = new System.Drawing.Point(272, 110);
			this.lblEstadoDT.Name = "lblEstadoDT";
			this.lblEstadoDT.Size = new System.Drawing.Size(0, 20);
			this.lblEstadoDT.TabIndex = 696;
			this.lblEstadoDT.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNuevoCosto
			// 
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNuevoCosto.Appearance = appearance6;
			this.txtNuevoCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNuevoCosto.Enabled = false;
			this.txtNuevoCosto.FormatString = "#,##0.00";
			this.txtNuevoCosto.Location = new System.Drawing.Point(113, 352);
			this.txtNuevoCosto.MaxValue = 10000;
			this.txtNuevoCosto.MinValue = -10000;
			this.txtNuevoCosto.Name = "txtNuevoCosto";
			this.txtNuevoCosto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNuevoCosto.PromptChar = ' ';
			this.txtNuevoCosto.Size = new System.Drawing.Size(96, 22);
			this.txtNuevoCosto.TabIndex = 697;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(9, 352);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(67, 17);
			this.label3.TabIndex = 698;
			this.label3.Text = "Nuevo Costo";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPrecioOferta
			// 
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecioOferta.Appearance = appearance7;
			this.txtPrecioOferta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecioOferta.Enabled = false;
			this.txtPrecioOferta.FormatString = "#,##0.00";
			this.txtPrecioOferta.Location = new System.Drawing.Point(113, 376);
			this.txtPrecioOferta.MaxValue = 10000;
			this.txtPrecioOferta.MinValue = -10000;
			this.txtPrecioOferta.Name = "txtPrecioOferta";
			this.txtPrecioOferta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecioOferta.PromptChar = ' ';
			this.txtPrecioOferta.Size = new System.Drawing.Size(96, 22);
			this.txtPrecioOferta.TabIndex = 699;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(9, 376);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(91, 17);
			this.label4.TabIndex = 700;
			this.label4.Text = "Precio Oferta(TC)";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservacionesCompras
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservacionesCompras.Appearance = appearance8;
			this.txtObservacionesCompras.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservacionesCompras.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservacionesCompras.Enabled = false;
			this.txtObservacionesCompras.Location = new System.Drawing.Point(217, 376);
			this.txtObservacionesCompras.Multiline = true;
			this.txtObservacionesCompras.Name = "txtObservacionesCompras";
			this.txtObservacionesCompras.Size = new System.Drawing.Size(344, 120);
			this.txtObservacionesCompras.TabIndex = 701;
			this.txtObservacionesCompras.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtObservacionesCompras_KeyDown);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(216, 352);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(34, 16);
			this.label5.TabIndex = 702;
			this.label5.Text = "Notas";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// ultraDataSource2
			// 
			ultraDataColumn2.DataType = typeof(System.Double);
			ultraDataColumn2.DefaultValue = 0;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource2;
			appearance9.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance9;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 220;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.CellAppearance = appearance10;
			ultraGridColumn2.Format = "#,##0.00";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 70;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 100;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 400;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 208);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(552, 128);
			this.uGridInformacion.TabIndex = 703;
			// 
			// txtNuevoPrecioConIVA
			// 
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNuevoPrecioConIVA.Appearance = appearance16;
			this.txtNuevoPrecioConIVA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNuevoPrecioConIVA.Enabled = false;
			this.txtNuevoPrecioConIVA.FormatString = "#,##0.00";
			this.txtNuevoPrecioConIVA.Location = new System.Drawing.Point(113, 448);
			this.txtNuevoPrecioConIVA.MaxValue = 10000;
			this.txtNuevoPrecioConIVA.MinValue = -10000;
			this.txtNuevoPrecioConIVA.Name = "txtNuevoPrecioConIVA";
			this.txtNuevoPrecioConIVA.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNuevoPrecioConIVA.PromptChar = ' ';
			this.txtNuevoPrecioConIVA.Size = new System.Drawing.Size(96, 22);
			this.txtNuevoPrecioConIVA.TabIndex = 704;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(9, 448);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(89, 17);
			this.label6.TabIndex = 705;
			this.label6.Text = "Precio Con I.V.A.";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(168, 499);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(35, 17);
			this.label8.TabIndex = 708;
			this.label8.Text = "Meses";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtGarantia
			// 
			this.txtGarantia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGarantia.Appearance = appearance17;
			this.txtGarantia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGarantia.Enabled = false;
			this.txtGarantia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtGarantia.FormatString = "#0";
			this.txtGarantia.Location = new System.Drawing.Point(113, 496);
			this.txtGarantia.Name = "txtGarantia";
			this.txtGarantia.PromptChar = ' ';
			this.txtGarantia.Size = new System.Drawing.Size(48, 22);
			this.txtGarantia.TabIndex = 707;
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label14.AutoSize = true;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(9, 499);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(46, 17);
			this.label14.TabIndex = 706;
			this.label14.Text = "Garantia";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmComprasEstado
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(570, 568);
			this.ControlBox = false;
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtGarantia);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtNuevoPrecioConIVA);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtObservacionesCompras);
			this.Controls.Add(this.txtPrecioOferta);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtNuevoCosto);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblEstadoDT);
			this.Controls.Add(this.txtMargen);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblEstadoCobranzas);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtObservacionesDT);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.txtDescuento);
			this.Controls.Add(this.txtNuevoPrecio);
			this.Controls.Add(this.lblNuevoPrecio);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtObservacionesCobranzas);
			this.Controls.Add(this.uGridInformacion);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.Name = "frmComprasEstado";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Asignación de Precio Articulos de Retiro";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmComprasEstado_KeyDown);
			this.Load += new System.EventHandler(this.frmComprasEstado_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesCobranzas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNuevoPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesDT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNuevoCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecioOferta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservacionesCompras)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNuevoPrecioConIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGarantia)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void txtDescuento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtObservacionesCompras.Focus();
		}

		private void frmComprasEstado_Load(object sender, System.EventArgs e)
		{
			this.lblEstadoCobranzas.Text = sEstadoProductoCOB;
			this.txtObservacionesCobranzas.Text = sObservacionesCobranzas;
			this.lblEstadoDT.Text = sEstadoProductoDT;
			this.txtObservacionesDT.Text = sObservacionesDT;
			this.txtNuevoCosto.Value = dNuevoCosto;
			this.txtPrecioOferta.Value = dPrecioOferta;			
			this.txtObservacionesCompras.Text = sObservacionesCompras;
			
			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DTRMConsultaAccesoriosCompras {0}", idDetalleIngresoRetiro));

			if (iEstadoCompras == 1)
			{
				decimal dDescuento = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Descuento From DTEstados Where idDTEstado = {0}", idDTEstadoDT));
				this.txtDescuento.Value = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Descuento From DTEstados Where idDTEstado = {0}", idDTEstadoDT));
				this.txtNuevoPrecio.Value = dPrecioOferta - (dPrecioOferta * Convert.ToDecimal(this.txtDescuento.Value))/100.00m;	
				this.txtMargen.Value = ((Convert.ToDecimal(this.txtNuevoPrecio.Value) - Convert.ToDecimal(this.txtNuevoCosto.Value)) / Convert.ToDecimal(this.txtNuevoPrecio.Value)) * 100.00m; 				
			}
			else 
			{
				this.txtDescuento.Value = dDescuento;
				this.txtNuevoPrecio.Value = dPrecioOferta;
				this.txtMargen.Value = dMargen;
			}

			if (iEstadoCompras == 1 || bModifica)
			{
				this.txtObservacionesCompras.Enabled = true;
				this.txtDescuento.Enabled = true;
				this.btnAceptar.Enabled = true;
			}
		}

		private void txtDescuento_Click(object sender, System.EventArgs e)
		{
			this.txtDescuento.SelectAll();
		}

		private void txtDescuento_Enter(object sender, System.EventArgs e)
		{
			this.txtDescuento.SelectAll();
		}

		private void txtNuevoPrecio_Enter(object sender, System.EventArgs e)
		{
			this.txtNuevoPrecio.SelectAll();
		}

		private void txtNuevoPrecio_Click(object sender, System.EventArgs e)
		{
			this.txtNuevoPrecio.SelectAll();
		}

		private void txtNuevoPrecio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar.Focus();
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (Convert.ToDecimal(this.txtDescuento.Value) == 0 && Convert.ToDecimal(this.txtNuevoPrecio.Value) == 0)
			{
				MessageBox.Show("Escriba valor de descuento o valor del nuevo precio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDescuento.Focus();
				return;
			}

			decimal dNPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
				DateTime.Today.ToString("yyyyMMdd")));

			#region Codigo Sugerir Descuento
//			decimal dNuevocosto = (decimal)this.txtNuevoCosto.Value;
//			decimal dNPrecioOderta = (decimal)this.txtPrecioOferta.Value;
//			decimal dNDescuento = Convert.ToDecimal(this.txtDescuento.Value);
//			decimal dNnuevoPrecio ;
//			decimal dNnuevoPrecioIva; 
//			decimal iCont1 =0;
//			for(iCont1=dNDescuento  ; iCont1 < dNDescuento + 5; iCont1 = iCont1 +  00.01m)
//			{
//				dNnuevoPrecio = dNPrecioOderta - (dNPrecioOderta * Convert.ToDecimal(iCont1 ))/100.00m;	
//				dNnuevoPrecioIva = Convert.ToDecimal(dNnuevoPrecio) + Convert.ToDecimal(dNnuevoPrecio) * dNPorcIva/100.00m;
//				decimal num1 = Math.Round(Convert.ToDecimal(dNnuevoPrecioIva),2);
//				int p_ente1 = (int) num1;
//				decimal p_dec1 = num1 - p_ente1;
//				if( p_dec1 == 0.00m)
//				{
//					break;
//				}
//			}
			#endregion Codigo Sugerir Descuento
     
			#region Valida el iva sin decimales 
			if (Convert.ToDecimal(this.txtDescuento.Value) == 0 && Convert.ToDecimal(this.txtNuevoPrecio.Value) == 0)
			{
				MessageBox.Show("Escriba valor de descuento o valor del nuevo precio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDescuento.Focus();
				return;
			}
//			decimal num = Math.Round(Convert.ToDecimal( this.txtNuevoPrecioConIVA.Value),2);
//			int p_ente = (int) num;
//			decimal p_dec = num - p_ente;
//			if( p_dec > 0.00m)
//			{
//				MessageBox.Show(string.Format("El valor total con I.V.A. no debe ir con decimales.\n\nPor favor pruebe con el siguiente descuento {0} ", iCont1));
//				this.txtDescuento.Focus();
//				return;
//			}
			#endregion Valida el iva sin decimales			
			
//			num = Math.Round(Convert.ToDecimal(this.txtNuevoPrecioConIVA.Value), 2);
//			p_ente = (int) num;
//			p_dec = num - p_ente;
//			
//			if (p_dec > 0.00m)
//			{
//				MessageBox.Show("El valor total con I.V.A. no debe ir con decimales", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.txtDescuento.Focus();
//				return;
//			}	
			
			this.DialogResult = DialogResult.OK;
		}

		private void frmComprasEstado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void txtDescuento_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.txtDescuento.Value == System.DBNull.Value) return;

			decimal dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
				DateTime.Today.ToString("yyyyMMdd")));

			this.txtNuevoPrecio.Value = dPrecioOferta - (dPrecioOferta * Convert.ToDecimal(this.txtDescuento.Value))/100.00m;	
			this.txtNuevoPrecioConIVA.Value = Convert.ToDecimal(this.txtNuevoPrecio.Value) + Convert.ToDecimal(this.txtNuevoPrecio.Value) * dPorcIva/100.00m;

			if (Convert.ToDecimal(this.txtNuevoPrecio.Value) > 0.00m)
				this.txtMargen.Value = ((Convert.ToDecimal(this.txtNuevoPrecio.Value) - Convert.ToDecimal(this.txtNuevoCosto.Value)) / Convert.ToDecimal(this.txtNuevoPrecio.Value)) * 100.00m; 				
			else this.txtMargen.Value = 0.00m;

			if (Convert.ToDecimal(this.txtNuevoPrecioConIVA.Value) >= 100.00m) 
				this.txtGarantia.Value = 3; 
			else this.txtGarantia.Value = 0;
		}

		private void txtObservacionesCompras_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAceptar.Focus();
		}

		private void txtNuevoPrecio_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
