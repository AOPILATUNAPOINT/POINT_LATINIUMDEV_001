using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmsolicitudvaciones.
	/// </summary>
	public class frmsolicitudvaciones : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaRetorna;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCedula;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaSalida;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtD_Solictados;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtD_restantes;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtT_dias;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidPersonal;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtobservacion;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidsolicitudvacaciones;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmsolicitudvaciones()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmsolicitudvaciones));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idsolicitudvacaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Perido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("N_dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("D_restantes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Linea_Producto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio_Contado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ApellidoPaterno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ApellidoMaterno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaNacimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sexo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto_Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total_Deuda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Score");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Calificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo_Vivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiempo_Vivienda_Meses");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa_Trabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiempo_Trabajo_Meses");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Negocio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiempo_Negocio_Meses");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			this.label2 = new System.Windows.Forms.Label();
			this.btnSalir = new System.Windows.Forms.Button();
			this.label14 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtobservacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.dtFechaRetorna = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCedula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFechaSalida = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtD_Solictados = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtT_dias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtD_restantes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidPersonal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidsolicitudvacaciones = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtobservacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRetorna)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtD_Solictados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtT_dias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtD_restantes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidsolicitudvacaciones)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(200, 104);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(81, 16);
			this.label2.TabIndex = 1061;
			this.label2.Text = "Dias Restantes";
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(416, 272);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1056;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(320, 10);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(44, 17);
			this.label14.TabIndex = 1053;
			this.label14.Text = "Número";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(8, 104);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(88, 16);
			this.label12.TabIndex = 1028;
			this.label12.Text = "Dias Solicitados ";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(200, 74);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 1024;
			this.label8.Text = "Hasta";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(8, 74);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(37, 16);
			this.label5.TabIndex = 1020;
			this.label5.Text = "Desde";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(8, 152);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(68, 16);
			this.label9.TabIndex = 1019;
			this.label9.Text = "Observación";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(200, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 1016;
			this.label3.Text = "Nombre";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(40, 16);
			this.label1.TabIndex = 1015;
			this.label1.Text = "Cedula";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 1014;
			this.label6.Text = "Fecha";
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(376, 7);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(152, 23);
			this.txtNumero.TabIndex = 1052;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(280, 152);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 1048;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(System.UInt32);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
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
																																 ultraDataColumn14});
			// 
			// txtobservacion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtobservacion.Appearance = appearance2;
			this.txtobservacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtobservacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtobservacion.Enabled = false;
			this.txtobservacion.Location = new System.Drawing.Point(56, 184);
			this.txtobservacion.Multiline = true;
			this.txtobservacion.Name = "txtobservacion";
			this.txtobservacion.Size = new System.Drawing.Size(384, 63);
			this.txtobservacion.TabIndex = 1041;
			this.txtobservacion.ValueChanged += new System.EventHandler(this.txtobservacion_ValueChanged);
			this.txtobservacion.Click += new System.EventHandler(this.txtobservacion_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(8, 256);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(520, 8);
			this.groupBox2.TabIndex = 1037;
			this.groupBox2.TabStop = false;
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 272);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 24);
			this.btnNuevo.TabIndex = 1036;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(88, 272);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 1035;
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
			this.btnEditar.Location = new System.Drawing.Point(168, 272);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 24);
			this.btnEditar.TabIndex = 1034;
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
			this.btnGuardar.Location = new System.Drawing.Point(248, 272);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 1033;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(336, 272);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 24);
			this.btnCancelar.TabIndex = 1032;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// dtFechaRetorna
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaRetorna.Appearance = appearance3;
			this.dtFechaRetorna.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaRetorna.DateButtons.Add(dateButton1);
			this.dtFechaRetorna.Enabled = false;
			this.dtFechaRetorna.Format = "dd/MM/yyyy";
			this.dtFechaRetorna.Location = new System.Drawing.Point(256, 72);
			this.dtFechaRetorna.Name = "dtFechaRetorna";
			this.dtFechaRetorna.NonAutoSizeHeight = 23;
			this.dtFechaRetorna.Size = new System.Drawing.Size(112, 21);
			this.dtFechaRetorna.SpinButtonsVisible = true;
			this.dtFechaRetorna.TabIndex = 1023;
			this.dtFechaRetorna.Value = ((object)(resources.GetObject("dtFechaRetorna.Value")));
			this.dtFechaRetorna.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaRetorna_KeyDown);
			this.dtFechaRetorna.ValueChanged += new System.EventHandler(this.dtFechaRetorna_ValueChanged);
			// 
			// txtNombre
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance4;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(256, 38);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(272, 21);
			this.txtNombre.TabIndex = 1018;
			this.txtNombre.ValueChanged += new System.EventHandler(this.txtNombre_ValueChanged);
			// 
			// txtCedula
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCedula.Appearance = appearance5;
			this.txtCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCedula.Enabled = false;
			this.txtCedula.Location = new System.Drawing.Point(56, 38);
			this.txtCedula.MaxLength = 10;
			this.txtCedula.Name = "txtCedula";
			this.txtCedula.Size = new System.Drawing.Size(136, 21);
			this.txtCedula.TabIndex = 1017;
			this.txtCedula.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCedula_KeyDown);
			this.txtCedula.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCedula_KeyPress);
			this.txtCedula.Validating += new System.ComponentModel.CancelEventHandler(this.txtCedula_Validating);
			this.txtCedula.Validated += new System.EventHandler(this.txtCedula_Validated);
			this.txtCedula.ValueChanged += new System.EventHandler(this.txtCedula_ValueChanged);
			// 
			// dtFechaSalida
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaSalida.Appearance = appearance6;
			this.dtFechaSalida.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaSalida.DateButtons.Add(dateButton2);
			this.dtFechaSalida.Enabled = false;
			this.dtFechaSalida.Format = "dd/MM/yyyy";
			this.dtFechaSalida.Location = new System.Drawing.Point(56, 72);
			this.dtFechaSalida.Name = "dtFechaSalida";
			this.dtFechaSalida.NonAutoSizeHeight = 23;
			this.dtFechaSalida.Size = new System.Drawing.Size(136, 21);
			this.dtFechaSalida.SpinButtonsVisible = true;
			this.dtFechaSalida.TabIndex = 1013;
			this.dtFechaSalida.Value = ((object)(resources.GetObject("dtFechaSalida.Value")));
			this.dtFechaSalida.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaSalida_KeyDown);
			this.dtFechaSalida.ValueChanged += new System.EventHandler(this.dtFechaSalida_ValueChanged);
			// 
			// dtFecha
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance7;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton3);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(136, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1012;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// txtD_Solictados
			// 
			this.txtD_Solictados.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtD_Solictados.Enabled = false;
			this.txtD_Solictados.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtD_Solictados.Location = new System.Drawing.Point(56, 120);
			this.txtD_Solictados.Name = "txtD_Solictados";
			this.txtD_Solictados.PromptChar = ' ';
			this.txtD_Solictados.Size = new System.Drawing.Size(112, 28);
			this.txtD_Solictados.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtD_Solictados.TabIndex = 1062;
			// 
			// txtT_dias
			// 
			this.txtT_dias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtT_dias.Enabled = false;
			this.txtT_dias.Location = new System.Drawing.Point(480, 72);
			this.txtT_dias.Name = "txtT_dias";
			this.txtT_dias.PromptChar = ' ';
			this.txtT_dias.Size = new System.Drawing.Size(48, 21);
			this.txtT_dias.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtT_dias.TabIndex = 1063;
			this.txtT_dias.ValueChanged += new System.EventHandler(this.txttT_dias_ValueChanged);
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
			// ultraDataSource9
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.DateTime);
			ultraDataColumn27.DataType = typeof(System.DateTime);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn44.DataType = typeof(bool);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44});
			// 
			// txtD_restantes
			// 
			this.txtD_restantes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtD_restantes.Enabled = false;
			this.txtD_restantes.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtD_restantes.Location = new System.Drawing.Point(200, 120);
			this.txtD_restantes.Name = "txtD_restantes";
			this.txtD_restantes.PromptChar = ' ';
			this.txtD_restantes.Size = new System.Drawing.Size(104, 28);
			this.txtD_restantes.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtD_restantes.TabIndex = 1064;
			this.txtD_restantes.ValueChanged += new System.EventHandler(this.txtD_restantes_ValueChanged);
			// 
			// txtidPersonal
			// 
			this.txtidPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidPersonal.Enabled = false;
			this.txtidPersonal.Location = new System.Drawing.Point(448, 152);
			this.txtidPersonal.Name = "txtidPersonal";
			this.txtidPersonal.PromptChar = ' ';
			this.txtidPersonal.Size = new System.Drawing.Size(56, 21);
			this.txtidPersonal.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidPersonal.TabIndex = 1067;
			this.txtidPersonal.Visible = false;
			this.txtidPersonal.ValueChanged += new System.EventHandler(this.txtidPersonal_ValueChanged);
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(440, 120);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(56, 21);
			this.txtEstado.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEstado.TabIndex = 1068;
			this.txtEstado.Visible = false;
			// 
			// txtidsolicitudvacaciones
			// 
			this.txtidsolicitudvacaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidsolicitudvacaciones.Enabled = false;
			this.txtidsolicitudvacaciones.Location = new System.Drawing.Point(448, 176);
			this.txtidsolicitudvacaciones.Name = "txtidsolicitudvacaciones";
			this.txtidsolicitudvacaciones.PromptChar = ' ';
			this.txtidsolicitudvacaciones.Size = new System.Drawing.Size(56, 21);
			this.txtidsolicitudvacaciones.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidsolicitudvacaciones.TabIndex = 1069;
			this.txtidsolicitudvacaciones.Visible = false;
			this.txtidsolicitudvacaciones.ValueChanged += new System.EventHandler(this.ultraNumericEditor1_ValueChanged);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(384, 74);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(87, 16);
			this.label4.TabIndex = 1070;
			this.label4.Text = "Dias Pendientes";
			this.label4.Visible = false;
			this.label4.Click += new System.EventHandler(this.label4_Click);
			// 
			// frmsolicitudvaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(536, 302);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtidsolicitudvacaciones);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtidPersonal);
			this.Controls.Add(this.txtD_restantes);
			this.Controls.Add(this.txtT_dias);
			this.Controls.Add(this.txtD_Solictados);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtobservacion);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.dtFechaRetorna);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtCedula);
			this.Controls.Add(this.dtFechaSalida);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmsolicitudvaciones";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Solicitud De Vacaciones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmsolicitudvaciones_KeyDown);
			this.Load += new System.EventHandler(this.frmsolicitudvaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtobservacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRetorna)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtD_Solictados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtT_dias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtD_restantes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidsolicitudvacaciones)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		bool bEliminaAlValidar = false;
		private bool bActualiza = true;

		int Dresiduo ;
//		int idBloqueaTransacciones = 0;
//		int Grupo = 0;
//		int Cont = 0;
//		int idPersonal = 0;
//		decimal VSolicitado = 0.00m;
//		decimal dAcreditado = 0.00m;
//		string sNotasAnulacion="";
		
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			#region  conexion 
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			#endregion conexion 

		}

		private void txtCedula_Validated(object sender, System.EventArgs e)
		{
			ValidaCedula();
		}

		private void txtCedula_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			ValidaCedula();
		}

		private void frmsolicitudvaciones_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0525");

			if (!Funcion.Permiso("1423", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Solicitar Vacaciones ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			
			this.dtFechaSalida.Value = DateTime.Now.AddDays(+1);
			this.dtFechaRetorna.Value = DateTime.Now.AddDays(+2 );

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
		}
		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);

			this.dtFechaSalida.CalendarInfo.MinDate = DateTime.Now.AddDays(+1);
			this.dtFechaRetorna.CalendarInfo.MinDate = DateTime.Now.AddDays(+1);
			this.dtFechaSalida.Value = DateTime.Now.AddDays(+ 2);
			this.dtFechaRetorna.Value = DateTime.Now.AddDays(+ 2);
			this.lblEstado.Text = "PENDIENTE";
			this.txtEstado.Value = 1;

			#region controles
			this.txtCedula.Enabled = true;
			this.dtFechaSalida.Enabled = true;	
			this.dtFechaRetorna.Enabled = true;	
			//			this.txtValorS.Enabled = true;
			this.txtobservacion.Enabled = true;
			//this.btnEmpresa.Enabled = true;
			this.txtobservacion.Enabled = true;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;	
			#endregion controles
			
			bNuevo = true;
			bEdicion = true;

			this.txtCedula.Focus();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			
			bNuevo = false;
			bEdicion = false;

			#region Vacia Campos

			this.dtFecha.Value = System.DBNull.Value; 
			this.txtNumero.Text ="";
			this.txtCedula.Text ="";
			this.txtNombre.Text ="";
			this.dtFechaSalida.Value = System.DBNull.Value; 
			this.dtFechaRetorna.Value = System.DBNull.Value; 
			this.txtD_Solictados.Value = 0;
			this.txtD_restantes.Value = 0;
			this.txtEstado.Value = 0;
			this.txtidPersonal.Value = 0;
			this.txtT_dias.Value = 0;
			this.txtidsolicitudvacaciones.Value = 0;
			this.txtobservacion.Text ="";

	#endregion Vacia Campos

			#region Bloquea Controles

			this.txtCedula.Enabled = false;
			this.dtFechaSalida.Enabled = false;	
			this.dtFechaRetorna.Enabled = false;	
			this.txtobservacion.Enabled = false;

			#endregion Bloquea Controles

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
		
		}

		private void dtFechaSalida_ValueChanged(object sender, System.EventArgs e)
		{
			ValidaDias();
		}

		private void dtFechaRetorna_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFechaSalida.Value != System.DBNull.Value)
			{
				if (this.dtFechaRetorna.Value != System.DBNull.Value)
				{
							if (this.txtCedula.Text.ToString().Length > 0)
								{
									this.txtD_Solictados.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Periodovacaciones '{0}','{1}'",
						Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFechaRetorna.Value).ToString("yyyyMMdd")));	
		
								this.txtT_dias.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ConsultaDiasVacaciones '{0}',2", this.txtCedula.Text));
							
								
								Dresiduo = Convert.ToInt32 (this.txtT_dias.Value) - Convert.ToInt32 (this.txtD_Solictados.Value);
								if (Dresiduo<= 0)
								{
									MessageBox.Show("Usuario no cumple con el periodo para  solicitar vacaciones " , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									this.txtCedula.Text = "";
									this.txtCedula.Focus();
								}
								else 

								this.txtD_restantes.Value=Dresiduo;
								this.dtFechaSalida.Enabled=false;
								this.dtFechaRetorna.Enabled=false;
								}
				}
			}
		}

		private void txttT_dias_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCedula_ValueChanged(object sender, System.EventArgs e)
		{
			//ValidaCEdula();
		}

		private void txtidDeposito_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtidPersonal_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtD_restantes_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (Dresiduo<= 0){
				MessageBox.Show("Usuario no cumple con el periodo para  solicitar vacaciones " , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCedula.Focus();
				return;
			}
				this.txtD_restantes.Value=Dresiduo;

			#region Validacion			
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaSalida, "Seleccione la Fecha De Salida", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaRetorna, "Seleccione la Fecha De Retorno", false, "", cdsSeteoF)) return;
			if (!Validacion.vbIdentificacion(this.txtCedula, 5, cdsSeteoF)) return;
			if (Convert.ToDateTime(this.dtFechaSalida.Value)!=Convert.ToDateTime(this.dtFechaRetorna.Value))
			{
				if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtFechaSalida.Value), Convert.ToDateTime(this.dtFechaRetorna.Value), this.dtFechaSalida, "La fecha de salida no puede ser posterior a la fecha de retorna", false)) return;
			}
			if (!Validacion.vbTexto(this.txtobservacion, 10, 100, "Observacion")) return;		
			#endregion Validacion	

			#region Valida estados

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Solicitudvacaciones Where idPersonal={0} and estado=1 and   Fecha >=  DATEADD(DAY, -7, GETDATE())", (int)this.txtidPersonal.Value)) > 0)
			{
				MessageBox.Show("No puede Grabar mientras  existan solicitides Pendientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Solicitudvacaciones Where idPersonal={0} and estado=2 and   Fecha >=  DATEADD(DAY, -7, GETDATE())", (int)this.txtidPersonal.Value)) > 0)
			{
				MessageBox.Show("No puede Grabar mientras  existan solicitides Aprobadas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			#endregion Valida estados


			#region guardar
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

					#region Numeracion
					if (bNuevo)
					{
						string ssSQL = string.Format("Exec NumeracionLocales 73, 0, 0");
						oCmdActualiza.CommandText = ssSQL;
						this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();						
					}
					#endregion Numeracion
 
					
					string  sSQLDet = string.Format("Exec GrabaSocilitudvacaciones {0}, '{1}', {2}, '{3}', '{4}', {5}, '{6}', '{7}', {8},{9}" ,
						(int)this.txtidsolicitudvacaciones.Value,
						this.txtNumero.Text,
						(int) this.txtD_Solictados.Value, 
						Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd"), 
						Convert.ToDateTime(this.dtFechaRetorna.Value).ToString("yyyyMMdd"),
						(int) this.txtT_dias.Value, 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
						this.txtobservacion.Text,
						(int)this.txtEstado.Value,
						(int)this.txtidPersonal.Value);

					oCmdActualiza.CommandText = sSQLDet;
					this.txtidsolicitudvacaciones.Value = (int)oCmdActualiza.ExecuteScalar();
          
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);						
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
			#endregion guardar	


			this.txtCedula.Enabled = false;
			this.dtFechaSalida.Enabled = false;
			this.dtFechaRetorna.Enabled = false;
			this.txtobservacion.Enabled = false;

			 this.btnNuevo.Enabled = true;
		   this.btnConsultar.Enabled = true;
			 this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
			
			bNuevo = false;
			bEdicion = false;

		}

		private void ultraNumericEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmconsultasolicitudvacaciones Busqueda = new frmconsultasolicitudvacaciones())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
				this.Consulta((int)Busqueda.uGridCreditos.ActiveRow.Cells["idsolicitudvacaciones"].Value);
				}
			}
		}

		private void label7_Click(object sender, System.EventArgs e)
		{
		}
		private void btnAnular_Click(object sender, System.EventArgs e)
		{
		
		}
		private void Consulta (int Id)
		{			
			DateTime dDesde=DateTime.Today ;
			DateTime dHasta=DateTime.Today ;

			string sSQL = string.Format("Exec ConsultaIndividualSolicitudVacaciones {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidsolicitudvacaciones.Value = Id;
				if (dr.GetValue(1) != System.DBNull.Value) this.dtFecha.Value = dr.GetDateTime(1);
				if (dr.GetValue(2) != System.DBNull.Value) this.txtidPersonal.Value = dr.GetInt32(2);
				if (dr.GetValue(3) != System.DBNull.Value) this.txtNombre.Text = dr.GetString(3); 
				if (dr.GetValue(4) != System.DBNull.Value) this.txtCedula.Text= dr.GetString(4); 
				if (dr.GetValue(5) != System.DBNull.Value) dDesde= dr.GetDateTime(5);
				if (dr.GetValue(6) != System.DBNull.Value) dHasta= dr.GetDateTime(6);
				if (dr.GetValue(7) != System.DBNull.Value) this.txtobservacion.Text = dr.GetString(7); 
				if (dr.GetValue(8) != System.DBNull.Value) this.txtEstado.Value = dr.GetInt32(8);
				if (dr.GetValue(9) != System.DBNull.Value) this.txtNumero.Text = dr.GetString(9); 
				if (dr.GetValue(10) != System.DBNull.Value) this.lblEstado.Text = dr.GetString(10);
			
			}
			dr.Close();

			this.dtFechaSalida.Value = dDesde;
			this.dtFechaRetorna.Value = dHasta;

		
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

//			if (miAcceso.Nuevo) 
				this.btnNuevo.Enabled = true;
//			if (miAcceso.Buscar) 
				this.btnConsultar.Enabled = true;
//			if (miAcceso.Editar)
				if ((int)this.txtEstado.Value == 1) 
				{
					this.btnEditar.Enabled = true;
				}
					

			if ((int)this.txtEstado.Value > 1) // PENDIENTE - AUTORIZADO
			{
				this.btnEditar.Enabled = false;
			}		
//			if ((int)this.txtEstado.Value == 1 ) this.lblEstado.Text = "POR AUTORIZAR";
//			if ((int)this.txtEstado.Value == 2 ) this.lblEstado.Text = "AUTORIZADO";
//			if ((int)this.txtEstado.Value == 3 ) this.lblEstado.Text = "RECHAZADO";
			
		}

		private void txtNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void frmsolicitudvaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();	
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{

			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El documento esta Aprobado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El documento esta Rechazado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("El documento esta Autorizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
      
			if ((int) this.txtEstado.Value == 1)
			{
//				this.txtCedula.Enabled = true;
				this.dtFechaRetorna.Enabled = true;
				this.dtFechaSalida.Enabled = true;
				this.txtobservacion.Enabled = true;
      }

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

		}

		private void label4_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();	
		}

		private void txtCedula_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{

		}

		private void txtCedula_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtFechaSalida.Focus();
		}
		private void ValidaCedula()
		{
			#region valida_cedula
			if(!bEdicion) return;
			if (!Validacion.vbTexto(this.txtCedula, 3, 10, "Cedula del Personal")) 
			{ 
				this.txtCedula.Focus();
				return;
			}

			bool bEmpleado = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec FacturacionValidaEmpleado '{0}'", this.txtCedula.Text.ToString()), true);
			if (bEmpleado)
			{
				string sSQLoc = string.Format(" Select (ApellidoPaterno + ' ' + ApellidoMaterno + ' ' + PrimerNombre + ' ' + SegundoNombre) as Nombre  from Nomina where  NIdentificacion= '{0}' ", this.txtCedula.Value);	
				this.txtNombre.Text = Funcion.sEscalarSQL(cdsSeteoF, sSQLoc);

				string sSQPer = string.Format("Select idPersonal from Nomina where  NIdentificacion = '{0}' ", this.txtCedula.Value);	
				this.txtidPersonal.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQPer);
			}

			if(this.txtNombre.Text.Equals(""))
			{
				MessageBox.Show("Empleado  Inactivo" , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCedula.Text = "";
				this.txtCedula.Focus();
			}

			if(this.txtNombre.Text.Equals(""))
			{
				MessageBox.Show("Empleado  Inactivo" , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCedula.Text = "";
				this.txtCedula.Focus();
			}

			
			if (this.txtCedula.Text.Length > 0)
			{
				this.txtD_restantes.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ConsultaDiasVacaciones '{0}',2", this.txtCedula.Text));
				this.txtCedula.Enabled=false;
			}
			#endregion valida_cedula
		}
		private void ValidaDias()
		{
			#region valida Dias 
			if (this.dtFechaSalida.Value != System.DBNull.Value)
			{
				if (this.dtFechaRetorna.Value != System.DBNull.Value)
				{
					if (this.txtCedula.Text.ToString().Length > 0)
					{
						this.txtD_Solictados.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Periodovacaciones '{0}','{1}'",
							Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFechaRetorna.Value).ToString("yyyyMMdd")));	
		
						this.txtT_dias.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ConsultaDiasVacaciones '{0}',2", this.txtCedula.Text));
							
								
						Dresiduo = Convert.ToInt32 (this.txtT_dias.Value) - Convert.ToInt32 (this.txtD_Solictados.Value);
						this.txtD_restantes.Value=Dresiduo;

					}
				}
			}
	
			#endregion valida Dias  
		}

		private void dtFechaSalida_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtFechaRetorna.Focus();
		}

		private void dtFechaRetorna_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtobservacion.Focus();
		}

		private void txtobservacion_Click(object sender, System.EventArgs e)
		{
			//ValidaDias();
		}

		private void txtobservacion_ValueChanged(object sender, System.EventArgs e)
		{
		ValidaDias();
		}
	}
}
