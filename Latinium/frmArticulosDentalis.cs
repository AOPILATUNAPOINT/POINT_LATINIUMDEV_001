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
	/// Descripción breve de frmArticulosDentalis.
	/// </summary>
	public class frmArticulosDentalis : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private DevExpress.XtraEditors.SimpleButton btGrupo;
		private DevExpress.XtraEditors.SimpleButton simpleButton1;
		private DevExpress.XtraEditors.SimpleButton simpleButton3;
		private DevExpress.XtraEditors.SimpleButton simpleButton4;
		private System.Windows.Forms.Label label7;
		private DevExpress.XtraEditors.SimpleButton simpleButton5;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdArticulo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDenominacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProveedor1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProveedor2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProcionCantidad;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProveedor3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbUnidadMedida;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTiempoCaducidad;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFrecuenciaAbas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEspecialidad;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtprecio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbUbicacionAlmacen;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmArticulosDentalis()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idArticulo = 0;

		public frmArticulosDentalis(int IdArticulo)
		{
			idArticulo = IdArticulo;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmArticulosDentalis));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidadMedida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Medida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidadMedida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Medida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTiempoCaducidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tiempo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTiempoCaducidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUbicaAlmacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ubicacion");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUbicaAlmacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFrecuencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Frecuencia");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFrecuencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Frecuencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEspecialidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Especialidad");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEspecialidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Especialidad");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btGrupo = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton3 = new DevExpress.XtraEditors.SimpleButton();
			this.simpleButton4 = new DevExpress.XtraEditors.SimpleButton();
			this.txtDenominacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtProcionCantidad = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.simpleButton5 = new DevExpress.XtraEditors.SimpleButton();
			this.txtProveedor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtProveedor2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtProveedor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtIdArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbUnidadMedida = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label10 = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTiempoCaducidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbUbicacionAlmacen = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbFrecuenciaAbas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEspecialidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtprecio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDenominacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProcionCantidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUnidadMedida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiempoCaducidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUbicacionAlmacen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFrecuenciaAbas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEspecialidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtprecio)).BeginInit();
			this.SuspendLayout();
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
			// btGrupo
			// 
			this.btGrupo.CausesValidation = false;
			this.btGrupo.Image = ((System.Drawing.Image)(resources.GetObject("btGrupo.Image")));
			this.btGrupo.Location = new System.Drawing.Point(168, 48);
			this.btGrupo.Name = "btGrupo";
			this.btGrupo.Size = new System.Drawing.Size(24, 22);
			this.btGrupo.TabIndex = 68;
			this.btGrupo.ToolTip = "Grupo de artículos";
			this.btGrupo.Click += new System.EventHandler(this.btGrupo_Click);
			// 
			// simpleButton1
			// 
			this.simpleButton1.CausesValidation = false;
			this.simpleButton1.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton1.Image")));
			this.simpleButton1.Location = new System.Drawing.Point(168, 111);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(24, 22);
			this.simpleButton1.TabIndex = 71;
			this.simpleButton1.ToolTip = "Grupo de artículos";
			this.simpleButton1.Visible = false;
			this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// simpleButton3
			// 
			this.simpleButton3.CausesValidation = false;
			this.simpleButton3.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton3.Image")));
			this.simpleButton3.Location = new System.Drawing.Point(168, 175);
			this.simpleButton3.Name = "simpleButton3";
			this.simpleButton3.Size = new System.Drawing.Size(24, 22);
			this.simpleButton3.TabIndex = 77;
			this.simpleButton3.ToolTip = "Grupo de artículos";
			this.simpleButton3.Click += new System.EventHandler(this.simpleButton3_Click);
			// 
			// simpleButton4
			// 
			this.simpleButton4.CausesValidation = false;
			this.simpleButton4.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton4.Image")));
			this.simpleButton4.Location = new System.Drawing.Point(168, 207);
			this.simpleButton4.Name = "simpleButton4";
			this.simpleButton4.Size = new System.Drawing.Size(24, 22);
			this.simpleButton4.TabIndex = 80;
			this.simpleButton4.ToolTip = "Grupo de artículos";
			this.simpleButton4.Click += new System.EventHandler(this.simpleButton4_Click);
			// 
			// txtDenominacion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDenominacion.Appearance = appearance1;
			this.txtDenominacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDenominacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDenominacion.Location = new System.Drawing.Point(208, 14);
			this.txtDenominacion.Name = "txtDenominacion";
			this.txtDenominacion.Size = new System.Drawing.Size(256, 21);
			this.txtDenominacion.TabIndex = 143;
			this.txtDenominacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtModelo_KeyDown);
			// 
			// txtProcionCantidad
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProcionCantidad.Appearance = appearance2;
			this.txtProcionCantidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProcionCantidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProcionCantidad.Location = new System.Drawing.Point(208, 78);
			this.txtProcionCantidad.Name = "txtProcionCantidad";
			this.txtProcionCantidad.Size = new System.Drawing.Size(256, 21);
			this.txtProcionCantidad.TabIndex = 146;
			this.txtProcionCantidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtProcionCantidad_KeyDown);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(488, 142);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(107, 16);
			this.label7.TabIndex = 149;
			this.label7.Text = "Descripción General";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.AcceptsReturn = true;
			this.txtDescripcion.AcceptsTab = true;
			this.txtDescripcion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance3;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(624, 112);
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(256, 76);
			this.txtDescripcion.TabIndex = 150;
			this.txtDescripcion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescripcion_KeyDown);
			// 
			// simpleButton5
			// 
			this.simpleButton5.CausesValidation = false;
			this.simpleButton5.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton5.Image")));
			this.simpleButton5.Location = new System.Drawing.Point(168, 141);
			this.simpleButton5.Name = "simpleButton5";
			this.simpleButton5.Size = new System.Drawing.Size(24, 22);
			this.simpleButton5.TabIndex = 74;
			this.simpleButton5.ToolTip = "Grupo de artículos";
			this.simpleButton5.Click += new System.EventHandler(this.simpleButton5_Click);
			// 
			// txtProveedor1
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProveedor1.Appearance = appearance4;
			this.txtProveedor1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProveedor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProveedor1.Location = new System.Drawing.Point(624, 14);
			this.txtProveedor1.Name = "txtProveedor1";
			this.txtProveedor1.Size = new System.Drawing.Size(256, 21);
			this.txtProveedor1.TabIndex = 151;
			this.txtProveedor1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtProveedor1_KeyDown);
			// 
			// txtProveedor2
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProveedor2.Appearance = appearance5;
			this.txtProveedor2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProveedor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProveedor2.Location = new System.Drawing.Point(624, 48);
			this.txtProveedor2.Name = "txtProveedor2";
			this.txtProveedor2.Size = new System.Drawing.Size(256, 21);
			this.txtProveedor2.TabIndex = 153;
			this.txtProveedor2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtProveedor2_KeyDown);
			// 
			// txtProveedor3
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProveedor3.Appearance = appearance6;
			this.txtProveedor3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtProveedor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProveedor3.Location = new System.Drawing.Point(624, 80);
			this.txtProveedor3.Name = "txtProveedor3";
			this.txtProveedor3.Size = new System.Drawing.Size(256, 21);
			this.txtProveedor3.TabIndex = 155;
			this.txtProveedor3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtProveedor3_KeyDown);
			this.txtProveedor3.ValueChanged += new System.EventHandler(this.ultraTextEditor4_ValueChanged);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(488, 50);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(66, 16);
			this.label12.TabIndex = 154;
			this.label12.Text = "Proveedor 2";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(488, 16);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(66, 16);
			this.label13.TabIndex = 152;
			this.label13.Text = "Proveedor 1";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(488, 82);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(66, 16);
			this.label14.TabIndex = 156;
			this.label14.Text = "Proveedor 3";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox4
			// 
			this.groupBox4.Location = new System.Drawing.Point(-16, 272);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1088, 8);
			this.groupBox4.TabIndex = 182;
			this.groupBox4.TabStop = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(16, 288);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 281;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(104, 288);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 285;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtIdArticulo
			// 
			this.txtIdArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdArticulo.Enabled = false;
			this.txtIdArticulo.Location = new System.Drawing.Point(856, 289);
			this.txtIdArticulo.Name = "txtIdArticulo";
			this.txtIdArticulo.PromptChar = ' ';
			this.txtIdArticulo.Size = new System.Drawing.Size(16, 21);
			this.txtIdArticulo.TabIndex = 286;
			this.txtIdArticulo.Visible = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 16);
			this.label1.TabIndex = 287;
			this.label1.Text = "Pulgadas";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(16, 80);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(62, 16);
			this.label3.TabIndex = 289;
			this.label3.Text = "Procesador";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(16, 114);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(36, 16);
			this.label4.TabIndex = 290;
			this.label4.Text = "Marca";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(16, 144);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(76, 16);
			this.label5.TabIndex = 291;
			this.label5.Text = "Memoria RAM";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(16, 178);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(60, 16);
			this.label6.TabIndex = 292;
			this.label6.Text = "Disco Duro";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(16, 210);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(36, 16);
			this.label8.TabIndex = 293;
			this.label8.Text = "Precio";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(16, 242);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(44, 16);
			this.label9.TabIndex = 294;
			this.label9.Text = "Gaming";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbUnidadMedida
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbUnidadMedida.Appearance = appearance7;
			this.cmbUnidadMedida.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbUnidadMedida.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUnidadMedida.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 255;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbUnidadMedida.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbUnidadMedida.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbUnidadMedida.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			this.cmbUnidadMedida.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			this.cmbUnidadMedida.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbUnidadMedida.DisplayMember = "Medida";
			this.cmbUnidadMedida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbUnidadMedida.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbUnidadMedida.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUnidadMedida.Location = new System.Drawing.Point(208, 48);
			this.cmbUnidadMedida.Name = "cmbUnidadMedida";
			this.cmbUnidadMedida.Size = new System.Drawing.Size(256, 21);
			this.cmbUnidadMedida.TabIndex = 297;
			this.cmbUnidadMedida.ValueMember = "idUnidadMedida";
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(16, 48);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(31, 16);
			this.label10.TabIndex = 288;
			this.label10.Text = "Color";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cmbTiempoCaducidad
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTiempoCaducidad.Appearance = appearance8;
			this.cmbTiempoCaducidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbTiempoCaducidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTiempoCaducidad.DataSource = this.ultraDataSource2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Tiempo";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 255;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbTiempoCaducidad.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTiempoCaducidad.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbTiempoCaducidad.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			this.cmbTiempoCaducidad.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			this.cmbTiempoCaducidad.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbTiempoCaducidad.DisplayMember = "tiempo";
			this.cmbTiempoCaducidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTiempoCaducidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTiempoCaducidad.Location = new System.Drawing.Point(208, 112);
			this.cmbTiempoCaducidad.Name = "cmbTiempoCaducidad";
			this.cmbTiempoCaducidad.Size = new System.Drawing.Size(256, 21);
			this.cmbTiempoCaducidad.TabIndex = 299;
			this.cmbTiempoCaducidad.ValueMember = "idTiempoCaducidad";
			// 
			// ultraDataSource3
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// cmbUbicacionAlmacen
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbUbicacionAlmacen.Appearance = appearance9;
			this.cmbUbicacionAlmacen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbUbicacionAlmacen.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUbicacionAlmacen.DataSource = this.ultraDataSource3;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 255;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbUbicacionAlmacen.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbUbicacionAlmacen.DisplayMember = "Ubicacion";
			this.cmbUbicacionAlmacen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbUbicacionAlmacen.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbUbicacionAlmacen.Location = new System.Drawing.Point(208, 144);
			this.cmbUbicacionAlmacen.Name = "cmbUbicacionAlmacen";
			this.cmbUbicacionAlmacen.Size = new System.Drawing.Size(256, 21);
			this.cmbUbicacionAlmacen.TabIndex = 300;
			this.cmbUbicacionAlmacen.ValueMember = "idUbicaAlmacen";
			// 
			// ultraDataSource4
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// cmbFrecuenciaAbas
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFrecuenciaAbas.Appearance = appearance10;
			this.cmbFrecuenciaAbas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbFrecuenciaAbas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFrecuenciaAbas.DataSource = this.ultraDataSource4;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 255;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbFrecuenciaAbas.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbFrecuenciaAbas.DisplayMember = "Frecuencia";
			this.cmbFrecuenciaAbas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFrecuenciaAbas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFrecuenciaAbas.Location = new System.Drawing.Point(208, 176);
			this.cmbFrecuenciaAbas.Name = "cmbFrecuenciaAbas";
			this.cmbFrecuenciaAbas.Size = new System.Drawing.Size(256, 21);
			this.cmbFrecuenciaAbas.TabIndex = 301;
			this.cmbFrecuenciaAbas.ValueMember = "idFrecuencia";
			// 
			// ultraDataSource5
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// cmbEspecialidad
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEspecialidad.Appearance = appearance11;
			this.cmbEspecialidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEspecialidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEspecialidad.DataSource = this.ultraDataSource5;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 255;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbEspecialidad.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbEspecialidad.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.cmbEspecialidad.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.SiblingRowsOnly;
			this.cmbEspecialidad.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			this.cmbEspecialidad.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.cmbEspecialidad.DisplayMember = "Especialidad";
			this.cmbEspecialidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEspecialidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEspecialidad.Location = new System.Drawing.Point(208, 208);
			this.cmbEspecialidad.Name = "cmbEspecialidad";
			this.cmbEspecialidad.Size = new System.Drawing.Size(256, 21);
			this.cmbEspecialidad.TabIndex = 302;
			this.cmbEspecialidad.ValueMember = "idEspecialidad";
			// 
			// txtprecio
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtprecio.Appearance = appearance12;
			this.txtprecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtprecio.FormatString = "#,##0.00";
			this.txtprecio.Location = new System.Drawing.Point(208, 240);
			this.txtprecio.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtprecio.Name = "txtprecio";
			this.txtprecio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtprecio.PromptChar = ' ';
			this.txtprecio.Size = new System.Drawing.Size(80, 21);
			this.txtprecio.TabIndex = 303;
			// 
			// frmArticulosDentalis
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(896, 318);
			this.Controls.Add(this.txtprecio);
			this.Controls.Add(this.cmbEspecialidad);
			this.Controls.Add(this.cmbFrecuenciaAbas);
			this.Controls.Add(this.cmbUbicacionAlmacen);
			this.Controls.Add(this.cmbTiempoCaducidad);
			this.Controls.Add(this.cmbUnidadMedida);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtIdArticulo);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.txtProveedor3);
			this.Controls.Add(this.txtProveedor2);
			this.Controls.Add(this.txtProveedor1);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtProcionCantidad);
			this.Controls.Add(this.txtDenominacion);
			this.Controls.Add(this.simpleButton4);
			this.Controls.Add(this.simpleButton3);
			this.Controls.Add(this.simpleButton1);
			this.Controls.Add(this.btGrupo);
			this.Controls.Add(this.simpleButton5);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label10);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmArticulosDentalis";
			this.Text = "Atributos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmArticulosDentalis_KeyDown);
			this.Load += new System.EventHandler(this.frmArticulosDentalis_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDenominacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProcionCantidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProveedor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUnidadMedida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiempoCaducidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUbicacionAlmacen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFrecuenciaAbas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEspecialidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtprecio)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmArticulosDentalis_Load(object sender, System.EventArgs e)
		{
			this.txtIdArticulo.Value = idArticulo;
			this.cmbUnidadMedida.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idUnidadMedida, Medida From UnidadMedida Order By Medida");
			this.cmbTiempoCaducidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTiempoCaducidad, tiempo from TiempoCaducidad Order By tiempo");
			this.cmbUbicacionAlmacen.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idUbicaAlmacen, Ubicacion from UbicacionAlmacenamiento Order By Ubicacion");
			this.cmbFrecuenciaAbas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFrecuencia, Frecuencia from FrecuenciaAbastecimiento Order By Frecuencia");
			this.cmbEspecialidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEspecialidad, Especialidad from EspecialidadUtilizar Order By Especialidad");
			
			string sSQL = string.Format("Exec ConsultaIndividualArticuloDentalis {0}", idArticulo);
			SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drArticulo.Read();
			if (drArticulo.HasRows)
			{	
				this.txtIdArticulo.Value = idArticulo;
				if (drArticulo.GetValue(1) != System.DBNull.Value) this.txtDenominacion.Text = drArticulo.GetString(1);
				if (drArticulo.GetValue(2) != System.DBNull.Value) this.cmbUnidadMedida.Value = drArticulo.GetInt32(2);
				if (drArticulo.GetValue(3) != System.DBNull.Value) this.txtProcionCantidad.Text = drArticulo.GetString(3);
				if (drArticulo.GetValue(4) != System.DBNull.Value) this.txtDescripcion.Text = drArticulo.GetString(4);
				if (drArticulo.GetValue(5) != System.DBNull.Value) this.cmbTiempoCaducidad.Value = drArticulo.GetInt32(5);
				if (drArticulo.GetValue(6) != System.DBNull.Value) this.cmbUbicacionAlmacen.Value = drArticulo.GetInt32(6);
				if (drArticulo.GetValue(7) != System.DBNull.Value) this.cmbFrecuenciaAbas.Value = drArticulo.GetInt32(7);
				if (drArticulo.GetValue(8) != System.DBNull.Value) this.txtProveedor1.Text = drArticulo.GetString(8);
				if (drArticulo.GetValue(9) != System.DBNull.Value) this.txtProveedor2.Text = drArticulo.GetString(9);
				if (drArticulo.GetValue(10) != System.DBNull.Value) this.txtProveedor3.Text = drArticulo.GetString(10);
				if (drArticulo.GetValue(11) != System.DBNull.Value) this.cmbEspecialidad.Value = drArticulo.GetInt32(11);
				if (drArticulo.GetValue(12) != System.DBNull.Value) this.txtprecio.Value = drArticulo.GetDecimal(12);
			}
			drArticulo.Close();
		}

		private void btGrupo_Click(object sender, System.EventArgs e)
		{
			using (ArtUnidadMedida Umedida = new ArtUnidadMedida())
			{
				if (DialogResult.OK == Umedida.ShowDialog())
				{
					this.cmbUnidadMedida.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idUnidadMedida, Medida From UnidadMedida Order By Medida");
					this.cmbUnidadMedida.Focus();
				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{

					#region Validacion

					if (!Validacion.vbComboVacio(this.cmbUnidadMedida, "Seleccione una Unidad de Medida")) return;
					if (!Validacion.vbComboVacio(this.cmbTiempoCaducidad, "Seleccione el Tiempo de Caducidad")) return;
					if (!Validacion.vbComboVacio(this.cmbUbicacionAlmacen, "Seleccione una Ubicaci{on de Almacen")) return;
					if (!Validacion.vbComboVacio(this.cmbFrecuenciaAbas, "Seleccione una Frecuencia")) return;
					if (!Validacion.vbComboVacio(this.cmbEspecialidad, "Seleccione una Especialidad")) return;

					if (!Validacion.vbTexto(this.txtDenominacion, 1, 500, "Denominacion")) return;
					if (!Validacion.vbTexto(this.txtProcionCantidad, 1, 500, "Porción Cantidad")) return;
					if (!Validacion.vbTexto(this.txtProveedor1, 1, 500, "Proveedor 1")) return;
//					if (!Validacion.vbTexto(this.txtProveedor2, 1, 500, "Proveedor 2")) return;
//					if (!Validacion.vbTexto(this.txtProveedor3, 1, 500, "Proveedor 3")) return;
					if (!Validacion.vbTexto(this.txtDescripcion, 1, 500, "Descripción")) return;
			    if (!Validacion.vbCampoDecimalVacio (this.txtprecio,"Ingrese el Precio Preferencial",1,10000)) return;	

//					if (!Validacion.vbTexto(this.txtDenominacion, 1, 500, "Denominacion")) return;
//					if (!Validacion.vbTexto(this.txtProcionCantidad, 1, 500, "Porción Cantidad")) return;
//					if (!Validacion.vbTexto(this.txtModelo, 1, 500, "Modelo")) return;
//
//					if (!Validacion.vbCampoEnteroVacio(this.txtProcionCantidad, "Ingrese Cantidad del artículo")) return;
//					if (!Validacion.vbCampoEnteroVacio(this.txtprecio, "Ingrese Precio Referencial")) return;
//					if (!Validacion.vbCampoEnteroVacio(this.txtDescripcion, "Ingrese una Descripción")) return;

					#endregion Validacion

					int IdEspecialidad  = 0;
					int	IdFrecuencia  = 0;
					int	IdTiempoCaducidad  = 0;
					int IdUnidadMedida  = 0;
					int IdUbicaAlmacen  = 0;
			
					if (this.cmbEspecialidad.ActiveRow != null) IdEspecialidad = (int)this.cmbEspecialidad.Value;
					if (this.cmbFrecuenciaAbas.ActiveRow != null) IdFrecuencia = (int)this.cmbFrecuenciaAbas.Value;
					if (this.cmbTiempoCaducidad.ActiveRow != null) IdTiempoCaducidad = (int)this.cmbTiempoCaducidad.Value;
					if (this.cmbUnidadMedida.ActiveRow != null) IdUnidadMedida = (int)this.cmbUnidadMedida.Value;
					if (this.cmbUbicacionAlmacen.ActiveRow != null) IdUbicaAlmacen = (int)this.cmbUbicacionAlmacen.Value;

					#region Graba
				
					string sSQLGuarda = string.Format("Exec ActualizarIndividualArticuloDentalis {0}, '{1}', {2}, '{3}', '{4}', {5}, {6}, {7}, '{8}', '{9}', '{10}', {11}, {12}" ,
						(int)this.txtIdArticulo.Value, 
						this.txtDenominacion.Text.ToString(),
						IdUnidadMedida, 
						this.txtProcionCantidad.Text.ToString(), 
						this.txtDescripcion.Text.ToString(),
						IdTiempoCaducidad, 
						IdUbicaAlmacen, 
						IdFrecuencia, 
						this.txtProveedor1.Text.ToString(), 
						this.txtProveedor2.Text.ToString(), 
						this.txtProveedor3.Text.ToString(), 
						IdEspecialidad, 
						Convert.ToDecimal(this.txtprecio.Value));

					Funcion.EjecutaSQL(cdsSeteoF, sSQLGuarda, true);  
//						this.txtIdArticulo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLGuarda, true);

					#endregion Graba
			
					this.DialogResult = DialogResult.OK;
			
		}
			private void txtModelo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtProveedor1.Focus();
		}

		private void ultraTextEditor4_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtProveedor1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbUnidadMedida.Focus();
		}

		private void cmbUnidadMedida_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtProveedor2.Focus();
		}

		private void txtProveedor2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtProcionCantidad.Focus();
		}

		private void txtProcionCantidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtProveedor3.Focus();
		}

		private void txtProveedor3_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTiempoCaducidad.Focus();
		}

		private void cmbTiempoCaducidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbUbicacionAlmacen.Focus();
		}

		private void cmdUbicacionAlmacen_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbFrecuenciaAbas.Focus();
		}

		private void cmbFrecuenciaAbas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDescripcion.Focus();
		}

		private void txtDescripcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbEspecialidad.Focus();
		}

		private void cmbEspecialidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtprecio.Focus();
		}

		private void txtprecio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			if (e.KeyCode == Keys.Enter) this.txtprecio.Focus();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			using (ArtTiempoCaducidad TiempoCaducidad = new ArtTiempoCaducidad())
			{
				if (DialogResult.OK == TiempoCaducidad.ShowDialog())
				{
					this.cmbTiempoCaducidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select idTiempoCaducidad, tiempo from TiempoCaducidad Order By tiempo");
					this.cmbTiempoCaducidad.Focus();
				}
			}
		}

		private void simpleButton5_Click(object sender, System.EventArgs e)
		{
			using (ArtUbicacionAlmacenamiento UbicaAlmacenamiento = new ArtUbicacionAlmacenamiento())
			{
				if (DialogResult.OK == UbicaAlmacenamiento.ShowDialog())
				{
					this.cmbUbicacionAlmacen.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select idUbicaAlmacen, Ubicacion from UbicacionAlmacenamiento Order By Ubicacion");
					this.cmbUbicacionAlmacen.Focus();
				}
			}
		}

		private void simpleButton3_Click(object sender, System.EventArgs e)
		{
			using (ArtFrecuenciaAbastecimiento FreAbast = new ArtFrecuenciaAbastecimiento())
			{
				if (DialogResult.OK == FreAbast.ShowDialog())
				{
					this.cmbFrecuenciaAbas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select idFrecuencia, Frecuencia from FrecuenciaAbastecimiento Order By Frecuencia");
					this.cmbFrecuenciaAbas.Focus();
				}
			}
		}

		private void simpleButton4_Click(object sender, System.EventArgs e)
		{
			using (ArtEspecialidadUtilizar EspUtilizar = new ArtEspecialidadUtilizar())
			{
				if (DialogResult.OK == EspUtilizar.ShowDialog())
				{
					this.cmbEspecialidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select idEspecialidad, Especialidad from EspecialidadUtilizar Order By Especialidad");
					this.cmbEspecialidad.Focus();
				}
			}

		}

		private void cmbUnidadMedida_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtIdArticulo.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idUnidadMedida, 0) From UnidadMedida Where idUnidadMedida = {0}", (int)this.cmbUnidadMedida.Value));
		}

		private void simpleButton2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmArticulosDentalis_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
	}
}
