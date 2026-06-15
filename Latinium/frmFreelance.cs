using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFreelance.
	/// </summary>
	public class frmFreelance : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCedula;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidDatosFreelance;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidDatosFreelancewEB;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtiEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombres;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidos;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEdad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNacimiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbResidencia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmail;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbViveCon;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNacionalidad;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtReferencia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbParentesco;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private System.Windows.Forms.Button btnGps;
		private System.Windows.Forms.Button btnCedula;
		private System.Windows.Forms.Button btnPlanilla;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccionRef;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chk2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chk3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chk4;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chk5;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chk6;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chk7;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chk8;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chk9;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chk10;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chk11;
		private System.Windows.Forms.Button btnDomicilio;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFreelance()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int idDatosFreelancewEB = 0;
		public frmFreelance(int IdDatosFreelancewEB)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			idDatosFreelancewEB = IdDatosFreelancewEB;
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("id");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("id");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("id");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("id");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("id");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFreelance));
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			this.label1 = new System.Windows.Forms.Label();
			this.txtCedula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNombres = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtApellidos = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEdad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbNacimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbResidencia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label8 = new System.Windows.Forms.Label();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtEmail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbViveCon = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkRuc = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbNacionalidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtiEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidDatosFreelancewEB = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidDatosFreelance = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.cmbParentesco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label12 = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnGps = new System.Windows.Forms.Button();
			this.label13 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.btnCedula = new System.Windows.Forms.Button();
			this.label17 = new System.Windows.Forms.Label();
			this.btnPlanilla = new System.Windows.Forms.Button();
			this.txtDireccionRef = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.chk2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chk3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chk4 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chk5 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chk6 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chk7 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chk9 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.chk10 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chk11 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chk8 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnDomicilio = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEdad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResidencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbViveCon)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbNacionalidad)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtiEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDatosFreelancewEB)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDatosFreelance)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccionRef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 35);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(45, 17);
			this.label1.TabIndex = 784;
			this.label1.Text = "Cedula";
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// txtCedula
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCedula.Appearance = appearance1;
			this.txtCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCedula.Enabled = false;
			this.txtCedula.Location = new System.Drawing.Point(72, 32);
			this.txtCedula.MaxLength = 10;
			this.txtCedula.Name = "txtCedula";
			this.txtCedula.Size = new System.Drawing.Size(128, 23);
			this.txtCedula.TabIndex = 785;
			this.txtCedula.ValueChanged += new System.EventHandler(this.txtCedula_ValueChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(208, 35);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(60, 17);
			this.label2.TabIndex = 786;
			this.label2.Text = "Nombres:";
			this.label2.Click += new System.EventHandler(this.label2_Click);
			// 
			// txtNombres
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombres.Appearance = appearance2;
			this.txtNombres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombres.Enabled = false;
			this.txtNombres.Location = new System.Drawing.Point(280, 32);
			this.txtNombres.MaxLength = 10;
			this.txtNombres.Name = "txtNombres";
			this.txtNombres.Size = new System.Drawing.Size(216, 23);
			this.txtNombres.TabIndex = 787;
			this.txtNombres.ValueChanged += new System.EventHandler(this.ultraTextEditor1_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(504, 35);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(61, 17);
			this.label3.TabIndex = 788;
			this.label3.Text = "Apellidos:";
			this.label3.Click += new System.EventHandler(this.label3_Click);
			// 
			// txtApellidos
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidos.Appearance = appearance3;
			this.txtApellidos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidos.Enabled = false;
			this.txtApellidos.Location = new System.Drawing.Point(568, 32);
			this.txtApellidos.MaxLength = 10;
			this.txtApellidos.Name = "txtApellidos";
			this.txtApellidos.Size = new System.Drawing.Size(216, 23);
			this.txtApellidos.TabIndex = 789;
			this.txtApellidos.ValueChanged += new System.EventHandler(this.ultraTextEditor2_ValueChanged);
			// 
			// txtEdad
			// 
			this.txtEdad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEdad.Enabled = false;
			this.txtEdad.Location = new System.Drawing.Point(840, 32);
			this.txtEdad.Name = "txtEdad";
			this.txtEdad.PromptChar = ' ';
			this.txtEdad.Size = new System.Drawing.Size(56, 23);
			this.txtEdad.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEdad.TabIndex = 866;
			this.txtEdad.ValueChanged += new System.EventHandler(this.txtidEgreso_ValueChanged);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(792, 35);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(38, 17);
			this.label4.TabIndex = 867;
			this.label4.Text = "Edad:";
			this.label4.Click += new System.EventHandler(this.label4_Click);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(16, 64);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(83, 17);
			this.label5.TabIndex = 869;
			this.label5.Text = "Nacionalidad:";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(448, 64);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(117, 17);
			this.label6.TabIndex = 871;
			this.label6.Text = "Ciudad Nacimiento:";
			// 
			// cmbNacimiento
			// 
			this.cmbNacimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNacimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNacimiento.DataSource = this.ultraDataSource2;
			ultraGridBand1.AddButtonCaption = "RolSucursal";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 208;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbNacimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbNacimiento.DisplayMember = "Nombre";
			this.cmbNacimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNacimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNacimiento.Enabled = false;
			this.cmbNacimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNacimiento.Location = new System.Drawing.Point(568, 61);
			this.cmbNacimiento.Name = "cmbNacimiento";
			this.cmbNacimiento.Size = new System.Drawing.Size(216, 22);
			this.cmbNacimiento.TabIndex = 870;
			this.cmbNacimiento.ValueMember = "id";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(16, 99);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(116, 17);
			this.label7.TabIndex = 873;
			this.label7.Text = "Ciudad Residencia:";
			// 
			// cmbResidencia
			// 
			this.cmbResidencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResidencia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResidencia.DataSource = this.ultraDataSource2;
			ultraGridBand2.AddButtonCaption = "RolSucursal";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 208;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbResidencia.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbResidencia.DisplayMember = "Nombre";
			this.cmbResidencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResidencia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbResidencia.Enabled = false;
			this.cmbResidencia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResidencia.Location = new System.Drawing.Point(144, 96);
			this.cmbResidencia.Name = "cmbResidencia";
			this.cmbResidencia.Size = new System.Drawing.Size(208, 22);
			this.cmbResidencia.TabIndex = 872;
			this.cmbResidencia.ValueMember = "id";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(368, 99);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(62, 17);
			this.label8.TabIndex = 874;
			this.label8.Text = "Dirección:";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDireccion
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance4;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Location = new System.Drawing.Point(440, 96);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(408, 23);
			this.txtDireccion.TabIndex = 875;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(16, 131);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(45, 17);
			this.label9.TabIndex = 876;
			this.label9.Text = "E-mail:";
			// 
			// txtEmail
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEmail.Appearance = appearance5;
			this.txtEmail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmail.Enabled = false;
			this.txtEmail.Location = new System.Drawing.Point(72, 128);
			this.txtEmail.MaxLength = 10;
			this.txtEmail.Name = "txtEmail";
			this.txtEmail.Size = new System.Drawing.Size(216, 23);
			this.txtEmail.TabIndex = 877;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(304, 128);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(67, 17);
			this.label10.TabIndex = 878;
			this.label10.Text = "N° Celular:";
			// 
			// txtTelefono
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance6;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Location = new System.Drawing.Point(384, 128);
			this.txtTelefono.MaxLength = 10;
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(144, 23);
			this.txtTelefono.TabIndex = 879;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(536, 131);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(57, 17);
			this.label11.TabIndex = 881;
			this.label11.Text = "Vive con:";
			// 
			// cmbViveCon
			// 
			this.cmbViveCon.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbViveCon.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbViveCon.DataSource = this.ultraDataSource2;
			ultraGridBand3.AddButtonCaption = "RolSucursal";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 208;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbViveCon.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbViveCon.DisplayMember = "Nombre";
			this.cmbViveCon.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbViveCon.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbViveCon.Enabled = false;
			this.cmbViveCon.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbViveCon.Location = new System.Drawing.Point(608, 128);
			this.cmbViveCon.Name = "cmbViveCon";
			this.cmbViveCon.Size = new System.Drawing.Size(136, 22);
			this.cmbViveCon.TabIndex = 880;
			this.cmbViveCon.ValueMember = "id";
			// 
			// chkRuc
			// 
			this.chkRuc.BackColor = System.Drawing.Color.Transparent;
			this.chkRuc.Enabled = false;
			this.chkRuc.Location = new System.Drawing.Point(840, 128);
			this.chkRuc.Name = "chkRuc";
			this.chkRuc.Size = new System.Drawing.Size(48, 21);
			this.chkRuc.TabIndex = 882;
			this.chkRuc.Text = "Ruc ";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cmbNacionalidad);
			this.groupBox1.Controls.Add(this.txtApellidos);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.txtNombres);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.txtCedula);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.txtEdad);
			this.groupBox1.Controls.Add(this.label5);
			this.groupBox1.Controls.Add(this.cmbNacimiento);
			this.groupBox1.Controls.Add(this.label6);
			this.groupBox1.Controls.Add(this.txtDireccion);
			this.groupBox1.Controls.Add(this.cmbResidencia);
			this.groupBox1.Controls.Add(this.label8);
			this.groupBox1.Controls.Add(this.label7);
			this.groupBox1.Controls.Add(this.txtEmail);
			this.groupBox1.Controls.Add(this.label9);
			this.groupBox1.Controls.Add(this.label11);
			this.groupBox1.Controls.Add(this.cmbViveCon);
			this.groupBox1.Controls.Add(this.chkRuc);
			this.groupBox1.Controls.Add(this.label10);
			this.groupBox1.Controls.Add(this.txtTelefono);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(904, 168);
			this.groupBox1.TabIndex = 883;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Datos Personales";
			// 
			// cmbNacionalidad
			// 
			this.cmbNacionalidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNacionalidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNacionalidad.DataSource = this.ultraDataSource2;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbNacionalidad.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbNacionalidad.DisplayMember = "Nombre";
			this.cmbNacionalidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNacionalidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNacionalidad.Enabled = false;
			this.cmbNacionalidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNacionalidad.Location = new System.Drawing.Point(144, 64);
			this.cmbNacionalidad.Name = "cmbNacionalidad";
			this.cmbNacionalidad.Size = new System.Drawing.Size(208, 22);
			this.cmbNacionalidad.TabIndex = 883;
			this.cmbNacionalidad.ValueMember = "id";
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.txtiEstado);
			this.groupBox2.Controls.Add(this.txtidDatosFreelancewEB);
			this.groupBox2.Controls.Add(this.txtidDatosFreelance);
			this.groupBox2.Controls.Add(this.lblEstado);
			this.groupBox2.Controls.Add(this.txtReferencia);
			this.groupBox2.Controls.Add(this.label14);
			this.groupBox2.Controls.Add(this.label16);
			this.groupBox2.Controls.Add(this.cmbParentesco);
			this.groupBox2.Controls.Add(this.label12);
			this.groupBox2.Controls.Add(this.txtCelular);
			this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(8, 184);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(864, 104);
			this.groupBox2.TabIndex = 884;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos Referencia";
			// 
			// txtiEstado
			// 
			this.txtiEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtiEstado.Enabled = false;
			this.txtiEstado.Location = new System.Drawing.Point(672, 32);
			this.txtiEstado.MinValue = 0;
			this.txtiEstado.Name = "txtiEstado";
			this.txtiEstado.PromptChar = ' ';
			this.txtiEstado.Size = new System.Drawing.Size(80, 23);
			this.txtiEstado.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtiEstado.TabIndex = 1029;
			this.txtiEstado.Visible = false;
			// 
			// txtidDatosFreelancewEB
			// 
			this.txtidDatosFreelancewEB.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidDatosFreelancewEB.Enabled = false;
			this.txtidDatosFreelancewEB.Location = new System.Drawing.Point(768, 64);
			this.txtidDatosFreelancewEB.MinValue = 0;
			this.txtidDatosFreelancewEB.Name = "txtidDatosFreelancewEB";
			this.txtidDatosFreelancewEB.PromptChar = ' ';
			this.txtidDatosFreelancewEB.Size = new System.Drawing.Size(80, 23);
			this.txtidDatosFreelancewEB.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtidDatosFreelancewEB.TabIndex = 1028;
			this.txtidDatosFreelancewEB.Visible = false;
			// 
			// txtidDatosFreelance
			// 
			this.txtidDatosFreelance.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidDatosFreelance.Enabled = false;
			this.txtidDatosFreelance.Location = new System.Drawing.Point(768, 32);
			this.txtidDatosFreelance.MinValue = 0;
			this.txtidDatosFreelance.Name = "txtidDatosFreelance";
			this.txtidDatosFreelance.PromptChar = ' ';
			this.txtidDatosFreelance.Size = new System.Drawing.Size(80, 23);
			this.txtidDatosFreelance.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtidDatosFreelance.TabIndex = 1027;
			this.txtidDatosFreelance.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(376, 66);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 19);
			this.lblEstado.TabIndex = 1026;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtReferencia
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtReferencia.Appearance = appearance7;
			this.txtReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtReferencia.Enabled = false;
			this.txtReferencia.Location = new System.Drawing.Point(144, 32);
			this.txtReferencia.MaxLength = 10;
			this.txtReferencia.Name = "txtReferencia";
			this.txtReferencia.Size = new System.Drawing.Size(216, 23);
			this.txtReferencia.TabIndex = 787;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(8, 35);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(126, 17);
			this.label14.TabIndex = 786;
			this.label14.Text = "Nombres Referencia:";
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(8, 67);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(69, 17);
			this.label16.TabIndex = 869;
			this.label16.Text = "Parentezco";
			// 
			// cmbParentesco
			// 
			this.cmbParentesco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbParentesco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbParentesco.DisplayMember = "Nombre";
			this.cmbParentesco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbParentesco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbParentesco.Enabled = false;
			this.cmbParentesco.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbParentesco.Location = new System.Drawing.Point(144, 64);
			this.cmbParentesco.Name = "cmbParentesco";
			this.cmbParentesco.Size = new System.Drawing.Size(216, 22);
			this.cmbParentesco.TabIndex = 868;
			this.cmbParentesco.ValueMember = "id";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(376, 35);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(67, 17);
			this.label12.TabIndex = 883;
			this.label12.Text = "N° Celular:";
			// 
			// txtCelular
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance8;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Location = new System.Drawing.Point(456, 32);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(144, 23);
			this.txtCelular.TabIndex = 884;
			// 
			// btnGps
			// 
			this.btnGps.Image = ((System.Drawing.Image)(resources.GetObject("btnGps.Image")));
			this.btnGps.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGps.Location = new System.Drawing.Point(128, 341);
			this.btnGps.Name = "btnGps";
			this.btnGps.Size = new System.Drawing.Size(24, 23);
			this.btnGps.TabIndex = 885;
			this.btnGps.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGps.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(8, 344);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(105, 16);
			this.label13.TabIndex = 886;
			this.label13.Text = "Ubicación domicilio:";
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(160, 344);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(120, 16);
			this.label15.TabIndex = 888;
			this.label15.Text = "Cédula y Pap. votación";
			// 
			// btnCedula
			// 
			this.btnCedula.Image = ((System.Drawing.Image)(resources.GetObject("btnCedula.Image")));
			this.btnCedula.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCedula.Location = new System.Drawing.Point(288, 341);
			this.btnCedula.Name = "btnCedula";
			this.btnCedula.Size = new System.Drawing.Size(24, 23);
			this.btnCedula.TabIndex = 887;
			this.btnCedula.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCedula.Click += new System.EventHandler(this.btnCedula_Click);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(320, 344);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(135, 16);
			this.label17.TabIndex = 890;
			this.label17.Text = "Planilla de servicio Basico";
			// 
			// btnPlanilla
			// 
			this.btnPlanilla.Image = ((System.Drawing.Image)(resources.GetObject("btnPlanilla.Image")));
			this.btnPlanilla.Location = new System.Drawing.Point(464, 341);
			this.btnPlanilla.Name = "btnPlanilla";
			this.btnPlanilla.Size = new System.Drawing.Size(32, 23);
			this.btnPlanilla.TabIndex = 889;
			this.btnPlanilla.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnPlanilla.Click += new System.EventHandler(this.btnPlanilla_Click);
			// 
			// txtDireccionRef
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccionRef.Appearance = appearance9;
			this.txtDireccionRef.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccionRef.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccionRef.Enabled = false;
			this.txtDireccionRef.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDireccionRef.Location = new System.Drawing.Point(64, 304);
			this.txtDireccionRef.Name = "txtDireccionRef";
			this.txtDireccionRef.Size = new System.Drawing.Size(800, 19);
			this.txtDireccionRef.TabIndex = 892;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 305);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(55, 16);
			this.label18.TabIndex = 891;
			this.label18.Text = "Dirección:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chk2
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.chk2.Appearance = appearance10;
			this.chk2.BackColor = System.Drawing.Color.Transparent;
			this.chk2.Location = new System.Drawing.Point(920, 16);
			this.chk2.Name = "chk2";
			this.chk2.Size = new System.Drawing.Size(128, 21);
			this.chk2.TabIndex = 893;
			this.chk2.Text = "Datos Personales";
			// 
			// chk3
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.chk3.Appearance = appearance11;
			this.chk3.BackColor = System.Drawing.Color.Transparent;
			this.chk3.Location = new System.Drawing.Point(920, 40);
			this.chk3.Name = "chk3";
			this.chk3.Size = new System.Drawing.Size(136, 21);
			this.chk3.TabIndex = 894;
			this.chk3.Text = "Datos Nacimiento";
			// 
			// chk4
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.chk4.Appearance = appearance12;
			this.chk4.BackColor = System.Drawing.Color.Transparent;
			this.chk4.Location = new System.Drawing.Point(920, 64);
			this.chk4.Name = "chk4";
			this.chk4.Size = new System.Drawing.Size(136, 21);
			this.chk4.TabIndex = 895;
			this.chk4.Text = "Datos Residencia";
			// 
			// chk5
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.chk5.Appearance = appearance13;
			this.chk5.BackColor = System.Drawing.Color.Transparent;
			this.chk5.Location = new System.Drawing.Point(920, 88);
			this.chk5.Name = "chk5";
			this.chk5.Size = new System.Drawing.Size(136, 21);
			this.chk5.TabIndex = 896;
			this.chk5.Text = "Datos Contacto";
			// 
			// chk6
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.chk6.Appearance = appearance14;
			this.chk6.BackColor = System.Drawing.Color.Transparent;
			this.chk6.Location = new System.Drawing.Point(920, 112);
			this.chk6.Name = "chk6";
			this.chk6.Size = new System.Drawing.Size(136, 21);
			this.chk6.TabIndex = 897;
			this.chk6.Text = "Datos Referencia";
			// 
			// chk7
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.chk7.Appearance = appearance15;
			this.chk7.BackColor = System.Drawing.Color.Transparent;
			this.chk7.Location = new System.Drawing.Point(920, 136);
			this.chk7.Name = "chk7";
			this.chk7.Size = new System.Drawing.Size(136, 21);
			this.chk7.TabIndex = 898;
			this.chk7.Text = "Ubicación domiciliaria";
			// 
			// chk9
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.chk9.Appearance = appearance16;
			this.chk9.BackColor = System.Drawing.Color.Transparent;
			this.chk9.Location = new System.Drawing.Point(920, 184);
			this.chk9.Name = "chk9";
			this.chk9.Size = new System.Drawing.Size(136, 21);
			this.chk9.TabIndex = 900;
			this.chk9.Text = "Servicio Basico";
			this.chk9.CheckedChanged += new System.EventHandler(this.chk9_CheckedChanged);
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
			// chk10
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.chk10.Appearance = appearance17;
			this.chk10.BackColor = System.Drawing.Color.Transparent;
			this.chk10.Location = new System.Drawing.Point(920, 208);
			this.chk10.Name = "chk10";
			this.chk10.Size = new System.Drawing.Size(136, 21);
			this.chk10.TabIndex = 901;
			this.chk10.Text = "Verificacion Terrena";
			// 
			// chk11
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.chk11.Appearance = appearance18;
			this.chk11.BackColor = System.Drawing.Color.Transparent;
			this.chk11.Location = new System.Drawing.Point(920, 232);
			this.chk11.Name = "chk11";
			this.chk11.Size = new System.Drawing.Size(136, 21);
			this.chk11.TabIndex = 902;
			this.chk11.Text = "Contrato";
			// 
			// chk8
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.chk8.Appearance = appearance19;
			this.chk8.BackColor = System.Drawing.Color.Transparent;
			this.chk8.Location = new System.Drawing.Point(920, 160);
			this.chk8.Name = "chk8";
			this.chk8.Size = new System.Drawing.Size(136, 21);
			this.chk8.TabIndex = 903;
			this.chk8.Text = "Cedula/Papeleta";
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(920, 304);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(84, 24);
			this.btnGuardar.TabIndex = 904;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click_1);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(920, 336);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(88, 23);
			this.btnAnular.TabIndex = 905;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Visible = false;
			// 
			// btnDomicilio
			// 
			this.btnDomicilio.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.btnDomicilio.Enabled = false;
			this.btnDomicilio.Image = ((System.Drawing.Image)(resources.GetObject("btnDomicilio.Image")));
			this.btnDomicilio.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDomicilio.Location = new System.Drawing.Point(1064, 232);
			this.btnDomicilio.Name = "btnDomicilio";
			this.btnDomicilio.Size = new System.Drawing.Size(25, 24);
			this.btnDomicilio.TabIndex = 1008;
			this.btnDomicilio.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDomicilio.Click += new System.EventHandler(this.btnDomicilio_Click);
			// 
			// frmFreelance
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1096, 374);
			this.Controls.Add(this.btnDomicilio);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.chk8);
			this.Controls.Add(this.chk11);
			this.Controls.Add(this.chk10);
			this.Controls.Add(this.chk9);
			this.Controls.Add(this.chk7);
			this.Controls.Add(this.chk6);
			this.Controls.Add(this.chk5);
			this.Controls.Add(this.chk4);
			this.Controls.Add(this.chk3);
			this.Controls.Add(this.chk2);
			this.Controls.Add(this.txtDireccionRef);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.btnPlanilla);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.btnCedula);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.btnGps);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmFreelance";
			this.Text = "Datos";
			this.Load += new System.EventHandler(this.frmFreelance_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEdad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResidencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbViveCon)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbNacionalidad)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtiEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDatosFreelancewEB)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDatosFreelance)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccionRef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void txtCedula_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label4_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraTextEditor2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void ultraTextEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtidEgreso_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			string latitude = sLatitud;
			string longitude = sLongitud;
	
			string url = "https://www.google.com/maps?q="+latitude+","+longitude+"&hl=es;z=10";
    

			try
			{
				Process.Start(url);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error opening browser: " + ex.Message);
			}
		}

		private void btnGuardar_Click_1(object sender, System.EventArgs e)
		{
			// Mensajes correspondientes a cada CheckBox
			string[] messages = new string[]
		{
			"Datos Personales", "Datos Nacimiento", "Datos Residencia", "Datos Contacto",
			"Datos Referencia", "Ubicación domiciliaria", "Cedula/Papeleta", 
			"Servicio Basico", "Verificacion Terrena", "Contrato"
		};

			// Capturar el estado de cada CheckBox en variables booleanas
			bool chk2Checked = this.chk2.Checked;
			bool chk3Checked = this.chk3.Checked;
			bool chk4Checked = this.chk4.Checked;
			bool chk5Checked = this.chk5.Checked;
			bool chk6Checked = this.chk6.Checked;
			bool chk7Checked = this.chk7.Checked;
			bool chk8Checked = this.chk8.Checked;
			bool chk9Checked = this.chk9.Checked;
			bool chk10Checked = this.chk10.Checked;
			bool chk11Checked = this.chk11.Checked;


			// Crear un arreglo con el estado de los CheckBox
			bool[] checkBoxStates = new bool[]
		{
			chk2Checked, chk3Checked, chk4Checked, chk5Checked,
			chk6Checked, chk7Checked, chk8Checked, chk9Checked,
			chk10Checked, chk11Checked
		};

			// Verificar si al menos un CheckBox está seleccionado
			bool atLeastOneChecked = false;
			for (int i = 0; i < checkBoxStates.Length; i++)
			{
				if (checkBoxStates[i])
				{
					atLeastOneChecked = true;
					break;
				}
			}

			if (!atLeastOneChecked)
			{
				MessageBox.Show("Por favor, seleccione al menos un CheckBox para continuar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			// Verificar si los CheckBox están seleccionados en orden
			for (int i = 0; i < checkBoxStates.Length; i++)
			{
				// Si el CheckBox actual está seleccionado, verificar los anteriores
				if (checkBoxStates[i])
				{
					
					for (int j = 0; j < i; j++)
					{
						if (!checkBoxStates[j])
						{
							MessageBox.Show("Por favor, seleccione " + messages[j] + " antes de seleccionar " + messages[i] + ".", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
							return;
						}
					}
				}
			}

			// Si todos los CheckBox están seleccionados en orden, proceder con la acción de guardar
			// Tu lógica de guardar aquí
		}

		string sLatitud = "";
		string sLongitud ="";
		string sUrlSB ="";
		string sUrlCedula ="";
		private void frmFreelance_Load(object sender, System.EventArgs e)
		{
			this.cmbNacionalidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "exec [ListasFreelance] 0");
			this.cmbNacimiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "exec [ListasFreelance] 1");
			this.cmbViveCon.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "exec [ListasFreelance] 2");
			this.cmbParentesco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "exec [ListasFreelance]3");

			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaReporteFreelance {0}",idDatosFreelancewEB); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			if (dr.HasRows)
			{
				if (dr.GetValue(0) != System.DBNull.Value) this.txtCedula.Text = dr.GetString (0);
				if (dr.GetValue(1) != System.DBNull.Value) this.txtNombres.Text = dr.GetString (1);
				if (dr.GetValue(2) != System.DBNull.Value) this.txtApellidos.Text = dr.GetString (2);
				if (dr.GetValue(3) != System.DBNull.Value)  this.txtEdad.Value = dr.GetInt32(3);
				if (dr.GetValue(4) != System.DBNull.Value)  this.cmbNacionalidad.Value = dr.GetInt32(4);
				if (dr.GetValue(5) != System.DBNull.Value) this.cmbNacimiento.Value = dr.GetInt32(5);
				if (dr.GetValue(6) != System.DBNull.Value) this.cmbResidencia.Value = dr.GetInt32(6);
				if (dr.GetValue(7) != System.DBNull.Value) this.chkRuc.Checked = dr.GetBoolean(7);
				if (dr.GetValue(8) != System.DBNull.Value) this.cmbResidencia.Value = dr.GetInt32(8);
				if (dr.GetValue(9) != System.DBNull.Value) this.txtDireccion.Value = dr.GetString(9);
				if (dr.GetValue(10) != System.DBNull.Value) this.txtEmail.Text = dr.GetString(10);
				if (dr.GetValue(11) != System.DBNull.Value) this.txtTelefono.Text = dr.GetString(11);
				if (dr.GetValue(12) != System.DBNull.Value) this.cmbViveCon.Value = dr.GetInt32(12);
				if (dr.GetValue(13) != System.DBNull.Value) this.txtReferencia.Text = dr.GetString(13);
				if (dr.GetValue(14) != System.DBNull.Value) this.txtCelular.Text = dr.GetString(14);
				if (dr.GetValue(15) != System.DBNull.Value) this.cmbParentesco.Value = dr.GetInt32(15);
				if (dr.GetValue(16) != System.DBNull.Value) this.txtidDatosFreelancewEB.Value = dr.GetInt32(16);
				if (dr.GetValue(17) != System.DBNull.Value) this.txtiEstado.Value = dr.GetInt32(17);
				if (dr.GetValue(18) != System.DBNull.Value) this.txtDireccionRef.Text = dr.GetString(18);
				if (dr.GetValue(19) != System.DBNull.Value) sLatitud = dr.GetString(19);
				if (dr.GetValue(20) != System.DBNull.Value) sLongitud = dr.GetString(20);
				if (dr.GetValue(21) != System.DBNull.Value) sUrlSB = dr.GetString(21);
				if (dr.GetValue(22) != System.DBNull.Value) sUrlCedula = dr.GetString(22);
				if (dr.GetValue(23) != System.DBNull.Value) this.lblEstado.Text = dr.GetString(23);
				if (dr.GetValue(24) != System.DBNull.Value) this.txtidDatosFreelance.Value = dr.GetInt32(24);


				
				
				

				
				
				

				/*if (dr.GetValue(7) != System.DBNull.Value) this.txtDireccion.Text = dr.GetString (7);
				if (dr.GetValue(8) != System.DBNull.Value) this.txtEmail.Text  = dr.GetString (8);
				if (dr.GetValue(9) != System.DBNull.Value) this.txtTelefono.Text = dr.GetString (9);
				if (dr.GetValue(10) != System.DBNull.Value) this.cmbViveCon.Value = dr.GetInt32(10);
				if (dr.GetValue(11) != System.DBNull.Value) this.txtReferencia.Text = dr.GetString (11);
				if (dr.GetValue(12) != System.DBNull.Value) this.txtCelular.Text = dr.GetString (12);
				if (dr.GetValue(13) != System.DBNull.Value) this.txtiEstado.Value = dr.GetInt32(13);
				if (dr.GetValue(14) != System.DBNull.Value)
				if (dr.GetValue(15) != System.DBNull.Value)
				if (dr.GetValue(16) != System.DBNull.Value)
				if (dr.GetValue(17) != System.DBNull.Value)
				if (dr.GetValue(18) != System.DBNull.Value)
				if (dr.GetValue(19) != System.DBNull.Value)
				if (dr.GetValue(20) != System.DBNull.Value)
				if (dr.GetValue(21) != System.DBNull.Value)
				if (dr.GetValue(22) != System.DBNull.Value)
				if (dr.GetValue(23) != System.DBNull.Value)





				this.txtIdBodega.Value= dr.GetInt32 (0);
				this.txtNumero.Value = dr.GetInt32 (1);
				this.txtCodigo.Text = dr.GetString (2);
				this.txtNombre.Text = dr.GetString (3);
				this.dtFecha.Value= dr.GetDateTime (4);
				if (dr.GetValue(5) != System.DBNull.Value) this.dtcierre.Value = dr.GetDateTime (5);
				if (dr.GetValue(6) != System.DBNull.Value)this.txtcprovincia.Text = dr.GetString (6);
				this.cmbCiudad.Value = dr.GetInt32(7);
				if (dr.GetValue(8) != System.DBNull.Value)this.txtdireccion.Text=dr.GetString(8); 
				if (dr.GetValue(9) != System.DBNull.Value)this.txtExtencion.Text= dr.GetString(9);
				if (dr.GetValue(10) != System.DBNull.Value)this.txtEMail.Text = dr.GetString (10);
				if (dr.GetValue(11) != System.DBNull.Value)this.cmbCentroDeCosto.Value = dr.GetInt32(11);
				this.chkActivo.Checked = dr.GetBoolean(12);
				this.chkMatriz.Checked = dr.GetBoolean(13);
				this.chkCaja.Checked= dr.GetBoolean(14);
				this.chkInventario.Checked= dr.GetBoolean(15);
				this.chkTransferencia.Checked= dr.GetBoolean(16);
				this.chkTransferenciaMultiple.Checked= dr.GetBoolean(17);
				this.chkCompra.Checked= dr.GetBoolean(18);
				this.chkNCP.Checked= dr.GetBoolean(19);
				this.chkDT.Checked= dr.GetBoolean(20);
				this.chkDestinoT.Checked= dr.GetBoolean(21);
				this.chkCompras.Checked= dr.GetBoolean(22);
				this.txtOrdenSpf.Value= dr.GetInt32 (23);
				this.chkFactura.Checked= dr.GetBoolean(24);
				this.cmbCategorias.Value= dr.GetInt32 (25);
				this.cmbSectores.Value= dr.GetInt32 (26);	
				this.chkCredito.Checked= dr.GetBoolean(27);
				this.chkPrincipal.Checked= dr.GetBoolean(28);
				if (dr.GetValue(29) != System.DBNull.Value)this.txtSerie.Text = dr.GetString(29);
				if (dr.GetValue(30) != System.DBNull.Value)this.txtEquipo.Text = dr.GetString (30);
				this.chkFS.Checked= dr.GetBoolean(31);
				this.chkPromociones.Checked= dr.GetBoolean(32);
				this.chkCarpas.Checked= dr.GetBoolean(33);
				this.chkImprimePrecios.Checked= dr.GetBoolean(34);
				this.PTC.Checked= dr.GetBoolean(35);
				this.chkPromocionesCombos.Checked= dr.GetBoolean(36);
				this.txtCreditoGrande.Text = dr.GetString (37);
				this.txtCreditoMediano.Text = dr.GetString (38);
				this.txtCreditoPequeño.Text = dr.GetString (39);
				if (dr.GetValue(40) != System.DBNull.Value)this.txtContado.Text = dr.GetString (40);
				if (dr.GetValue(41) != System.DBNull.Value)this.chkNAut.Checked= dr.GetBoolean(41);
				if (dr.GetValue(42) != System.DBNull.Value)this.cmbRegion.Value = dr.GetInt32(42);
				if (dr.GetValue(43) != System.DBNull.Value)	this.chkGRAut.Checked= dr.GetBoolean(43);
				if (dr.GetValue(44) != System.DBNull.Value)this.chkBuffer2.Checked= dr.GetBoolean(44);*/
			}
			dr.Close();
			#endregion Consulta Individual

			
			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		
	
		}

		private void btnCedula_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process.Start("chrome.exe", sUrlSB);
		}

		private void btnPlanilla_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process.Start("chrome.exe", sUrlCedula);
		}

		private void chk9_CheckedChanged(object sender, System.EventArgs e)
		{
			this.btnDomicilio.Enabled = this.chk9.Checked;
		}

		private void btnDomicilio_Click(object sender, System.EventArgs e)
		{
			#region Cobrador


			using (frmTerrenaFreelance Gestion = new frmTerrenaFreelance((int) this.txtidDatosFreelancewEB.Value, (string) this.txtCedula.Text))
			{
				if (DialogResult.OK == Gestion.ShowDialog( ))
				{

	
				}
			}
			#endregion Cobrador
		}
	
	}
}
