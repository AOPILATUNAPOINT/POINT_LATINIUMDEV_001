using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmMatriz.
	/// </summary>
	public class frmMatriz : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblDepartamento;
		private System.Windows.Forms.Label lblCorreo;
		private System.Windows.Forms.Label lblExtension;
		private System.Windows.Forms.Label lblPersonal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtExtension;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPersonal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidExtensiones;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridContacto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtemail;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private System.Windows.Forms.Label lblip;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIp;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private System.Windows.Forms.Label lblCelular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtdireccion;
		private System.Windows.Forms.Label lblNotas;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet opdAlmacen;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet ultraOptionSet1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbConsultaDept;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbconsultaB;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmMatriz()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmMatriz));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idExtension");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Area");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Personal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ext");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("celular");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMatriz");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IP");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idExtension");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Area");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Personal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Email");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ext");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMatriz");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			this.lblDepartamento = new System.Windows.Forms.Label();
			this.lblPersonal = new System.Windows.Forms.Label();
			this.lblCorreo = new System.Windows.Forms.Label();
			this.lblExtension = new System.Windows.Forms.Label();
			this.txtExtension = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtemail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPersonal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidExtensiones = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.uGridContacto = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblip = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtIp = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbConsultaDept = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.opdAlmacen = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCelular = new System.Windows.Forms.Label();
			this.txtdireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblNotas = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraOptionSet1 = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cmbconsultaB = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.txtExtension)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtemail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidExtensiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridContacto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIp)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbConsultaDept)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.opdAlmacen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraOptionSet1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbconsultaB)).BeginInit();
			this.SuspendLayout();
			// 
			// lblDepartamento
			// 
			this.lblDepartamento.AutoSize = true;
			this.lblDepartamento.BackColor = System.Drawing.Color.Transparent;
			this.lblDepartamento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDepartamento.Location = new System.Drawing.Point(24, 50);
			this.lblDepartamento.Name = "lblDepartamento";
			this.lblDepartamento.Size = new System.Drawing.Size(76, 17);
			this.lblDepartamento.TabIndex = 369;
			this.lblDepartamento.Text = "Departamento";
			this.lblDepartamento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblPersonal
			// 
			this.lblPersonal.AutoSize = true;
			this.lblPersonal.BackColor = System.Drawing.Color.Transparent;
			this.lblPersonal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPersonal.Location = new System.Drawing.Point(24, 80);
			this.lblPersonal.Name = "lblPersonal";
			this.lblPersonal.Size = new System.Drawing.Size(47, 17);
			this.lblPersonal.TabIndex = 371;
			this.lblPersonal.Text = "Personal";
			this.lblPersonal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblPersonal.Click += new System.EventHandler(this.lblPersonal_Click);
			// 
			// lblCorreo
			// 
			this.lblCorreo.AutoSize = true;
			this.lblCorreo.BackColor = System.Drawing.Color.Transparent;
			this.lblCorreo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCorreo.Location = new System.Drawing.Point(24, 144);
			this.lblCorreo.Name = "lblCorreo";
			this.lblCorreo.Size = new System.Drawing.Size(37, 17);
			this.lblCorreo.TabIndex = 372;
			this.lblCorreo.Text = "Correo";
			this.lblCorreo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblExtension
			// 
			this.lblExtension.AutoSize = true;
			this.lblExtension.BackColor = System.Drawing.Color.Transparent;
			this.lblExtension.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblExtension.Location = new System.Drawing.Point(24, 114);
			this.lblExtension.Name = "lblExtension";
			this.lblExtension.Size = new System.Drawing.Size(23, 17);
			this.lblExtension.TabIndex = 375;
			this.lblExtension.Text = "Ext.";
			this.lblExtension.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtExtension
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtExtension.Appearance = appearance1;
			this.txtExtension.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtExtension.Enabled = false;
			this.txtExtension.Location = new System.Drawing.Point(112, 112);
			this.txtExtension.Name = "txtExtension";
			this.txtExtension.PromptChar = ' ';
			this.txtExtension.Size = new System.Drawing.Size(216, 21);
			this.txtExtension.TabIndex = 376;
			this.txtExtension.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtExtension_KeyDown);
			// 
			// txtemail
			// 
			this.txtemail.AcceptsTab = true;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtemail.Appearance = appearance2;
			this.txtemail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtemail.Enabled = false;
			this.txtemail.Location = new System.Drawing.Point(112, 142);
			this.txtemail.MaxLength = 10000;
			this.txtemail.Name = "txtemail";
			this.txtemail.Size = new System.Drawing.Size(216, 21);
			this.txtemail.TabIndex = 377;
			this.txtemail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtemail_KeyDown);
			// 
			// txtPersonal
			// 
			this.txtPersonal.AcceptsTab = true;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPersonal.Appearance = appearance3;
			this.txtPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPersonal.Enabled = false;
			this.txtPersonal.Location = new System.Drawing.Point(112, 80);
			this.txtPersonal.MaxLength = 100000;
			this.txtPersonal.Name = "txtPersonal";
			this.txtPersonal.Size = new System.Drawing.Size(216, 21);
			this.txtPersonal.TabIndex = 378;
			this.txtPersonal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPersonal_KeyDown);
			// 
			// txtidExtensiones
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidExtensiones.Appearance = appearance4;
			this.txtidExtensiones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidExtensiones.Enabled = false;
			this.txtidExtensiones.Location = new System.Drawing.Point(776, 8);
			this.txtidExtensiones.Name = "txtidExtensiones";
			this.txtidExtensiones.PromptChar = ' ';
			this.txtidExtensiones.Size = new System.Drawing.Size(16, 21);
			this.txtidExtensiones.TabIndex = 379;
			this.txtidExtensiones.Visible = false;
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
			this.btnEditar.Location = new System.Drawing.Point(200, 216);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(72, 23);
			this.btnEditar.TabIndex = 383;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(288, 216);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 382;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(376, 216);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 384;
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
			this.btnSalir.Location = new System.Drawing.Point(472, 216);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 385;
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
			this.btnNuevo.Location = new System.Drawing.Point(112, 216);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 380;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(464, 264);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 23);
			this.btnConsultar.TabIndex = 381;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// uGridContacto
			// 
			this.uGridContacto.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridContacto.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridContacto.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			this.uGridContacto.DisplayLayout.Appearance = appearance5;
			this.uGridContacto.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "RetDetalle";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 53;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 90;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 103;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.Header.Caption = "E-Mail";
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Width = 87;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance6;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 43;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance7;
			ultraGridColumn6.Header.Caption = "Celular";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 54;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 118;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 41;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 77;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance8;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 192;
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
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			this.uGridContacto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridContacto.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridContacto.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridContacto.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridContacto.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridContacto.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridContacto.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridContacto.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.uGridContacto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridContacto.Location = new System.Drawing.Point(0, 304);
			this.uGridContacto.Name = "uGridContacto";
			this.uGridContacto.Size = new System.Drawing.Size(808, 192);
			this.uGridContacto.TabIndex = 433;
			this.uGridContacto.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridContacto_BeforeSelectChange);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
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
																																 ultraDataColumn11});
			// 
			// cmbDepartamento
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamento.Appearance = appearance13;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDepartamento.Enabled = false;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(112, 48);
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(216, 21);
			this.cmbDepartamento.TabIndex = 436;
			this.cmbDepartamento.ValueMember = "idDepartamento";
			this.cmbDepartamento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDepartamento_KeyDown);
			// 
			// lblip
			// 
			this.lblip.AutoSize = true;
			this.lblip.BackColor = System.Drawing.Color.Transparent;
			this.lblip.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblip.Location = new System.Drawing.Point(384, 48);
			this.lblip.Name = "lblip";
			this.lblip.Size = new System.Drawing.Size(15, 17);
			this.lblip.TabIndex = 437;
			this.lblip.Text = "IP";
			this.lblip.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// txtIp
			// 
			this.txtIp.AcceptsTab = true;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIp.Appearance = appearance14;
			this.txtIp.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtIp.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIp.Enabled = false;
			this.txtIp.Location = new System.Drawing.Point(424, 46);
			this.txtIp.MaxLength = 1000;
			this.txtIp.Name = "txtIp";
			this.txtIp.Size = new System.Drawing.Size(216, 21);
			this.txtIp.TabIndex = 438;
			this.txtIp.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtIp_KeyDown);
			this.txtIp.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIp_KeyPress);
			// 
			// cmbConsultaDept
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbConsultaDept.Appearance = appearance15;
			this.cmbConsultaDept.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbConsultaDept.DisplayMember = "Descripcion";
			this.cmbConsultaDept.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbConsultaDept.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbConsultaDept.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbConsultaDept.Location = new System.Drawing.Point(216, 264);
			this.cmbConsultaDept.Name = "cmbConsultaDept";
			this.cmbConsultaDept.Size = new System.Drawing.Size(216, 21);
			this.cmbConsultaDept.TabIndex = 439;
			this.cmbConsultaDept.ValueMember = "idDepartamento";
			// 
			// opdAlmacen
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.opdAlmacen.Appearance = appearance16;
			this.opdAlmacen.CheckedIndex = 0;
			this.opdAlmacen.Enabled = false;
			this.opdAlmacen.ItemAppearance = appearance17;
			this.opdAlmacen.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Matriz";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Almacen";
			this.opdAlmacen.Items.Add(valueListItem1);
			this.opdAlmacen.Items.Add(valueListItem2);
			this.opdAlmacen.ItemSpacingVertical = 10;
			this.opdAlmacen.Location = new System.Drawing.Point(112, 8);
			this.opdAlmacen.Name = "opdAlmacen";
			this.opdAlmacen.Size = new System.Drawing.Size(152, 26);
			this.opdAlmacen.TabIndex = 612;
			this.opdAlmacen.Text = "Matriz";
			this.opdAlmacen.ValueChanged += new System.EventHandler(this.opdAlmacen_ValueChanged);
			// 
			// txtCelular
			// 
			this.txtCelular.AcceptsTab = true;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance18;
			this.txtCelular.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Location = new System.Drawing.Point(112, 176);
			this.txtCelular.MaxLength = 100000;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(216, 21);
			this.txtCelular.TabIndex = 614;
			this.txtCelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCelular_KeyDown);
			this.txtCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCelular_KeyPress);
			// 
			// lblCelular
			// 
			this.lblCelular.AutoSize = true;
			this.lblCelular.BackColor = System.Drawing.Color.Transparent;
			this.lblCelular.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCelular.Location = new System.Drawing.Point(24, 176);
			this.lblCelular.Name = "lblCelular";
			this.lblCelular.Size = new System.Drawing.Size(38, 17);
			this.lblCelular.TabIndex = 613;
			this.lblCelular.Text = "Celular";
			this.lblCelular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtdireccion
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdireccion.Appearance = appearance19;
			this.txtdireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtdireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdireccion.Enabled = false;
			this.txtdireccion.Location = new System.Drawing.Point(424, 80);
			this.txtdireccion.Multiline = true;
			this.txtdireccion.Name = "txtdireccion";
			this.txtdireccion.Size = new System.Drawing.Size(216, 112);
			this.txtdireccion.TabIndex = 615;
			// 
			// lblNotas
			// 
			this.lblNotas.AutoSize = true;
			this.lblNotas.BackColor = System.Drawing.Color.Transparent;
			this.lblNotas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblNotas.Location = new System.Drawing.Point(384, 112);
			this.lblNotas.Name = "lblNotas";
			this.lblNotas.Size = new System.Drawing.Size(33, 17);
			this.lblNotas.TabIndex = 616;
			this.lblNotas.Text = "Notas";
			this.lblNotas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance20;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbBodega.Location = new System.Drawing.Point(112, 48);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(216, 21);
			this.cmbBodega.TabIndex = 621;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.Visible = false;
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.groupBox1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.groupBox1.Location = new System.Drawing.Point(-8, 248);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(808, 8);
			this.groupBox1.TabIndex = 622;
			this.groupBox1.TabStop = false;
			// 
			// ultraOptionSet1
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraOptionSet1.Appearance = appearance21;
			this.ultraOptionSet1.CheckedIndex = 0;
			this.ultraOptionSet1.ItemAppearance = appearance22;
			this.ultraOptionSet1.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Matriz";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Almacen";
			this.ultraOptionSet1.Items.Add(valueListItem3);
			this.ultraOptionSet1.Items.Add(valueListItem4);
			this.ultraOptionSet1.ItemSpacingVertical = 10;
			this.ultraOptionSet1.Location = new System.Drawing.Point(40, 264);
			this.ultraOptionSet1.Name = "ultraOptionSet1";
			this.ultraOptionSet1.Size = new System.Drawing.Size(152, 26);
			this.ultraOptionSet1.TabIndex = 623;
			this.ultraOptionSet1.Text = "Matriz";
			this.ultraOptionSet1.ValueChanged += new System.EventHandler(this.ultraOptionSet1_ValueChanged);
			// 
			// cmbconsultaB
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbconsultaB.Appearance = appearance23;
			this.cmbconsultaB.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbconsultaB.DisplayMember = "Nombre";
			this.cmbconsultaB.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbconsultaB.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbconsultaB.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbconsultaB.Location = new System.Drawing.Point(216, 264);
			this.cmbconsultaB.Name = "cmbconsultaB";
			this.cmbconsultaB.Size = new System.Drawing.Size(216, 21);
			this.cmbconsultaB.TabIndex = 625;
			this.cmbconsultaB.ValueMember = "Bodega";
			this.cmbconsultaB.Visible = false;
			// 
			// frmMatriz
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(808, 494);
			this.Controls.Add(this.cmbconsultaB);
			this.Controls.Add(this.ultraOptionSet1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblNotas);
			this.Controls.Add(this.lblCelular);
			this.Controls.Add(this.lblip);
			this.Controls.Add(this.lblExtension);
			this.Controls.Add(this.lblCorreo);
			this.Controls.Add(this.lblPersonal);
			this.Controls.Add(this.lblDepartamento);
			this.Controls.Add(this.txtdireccion);
			this.Controls.Add(this.txtCelular);
			this.Controls.Add(this.opdAlmacen);
			this.Controls.Add(this.cmbConsultaDept);
			this.Controls.Add(this.txtIp);
			this.Controls.Add(this.cmbDepartamento);
			this.Controls.Add(this.uGridContacto);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.txtidExtensiones);
			this.Controls.Add(this.txtPersonal);
			this.Controls.Add(this.txtemail);
			this.Controls.Add(this.txtExtension);
			this.KeyPreview = true;
			this.Name = "frmMatriz";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Directorio";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmMatriz_KeyDown);
			this.Load += new System.EventHandler(this.frmMatriz_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtExtension)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtemail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidExtensiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridContacto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIp)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbConsultaDept)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.opdAlmacen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraOptionSet1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbconsultaB)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
    
		bool bNuevo = false;
		bool bEdicion = false;
		string sDireccion ="";
		int iBodega =0;

		private void frmMatriz_Load(object sender, System.EventArgs e)
		{
			this.cmbDepartamento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " select idDepartamento,Descripcion from RolDepartamento ");
		  this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre From Bodega Where Activo = 1 And Factura = 1 Order By Nombre");
			this.cmbConsultaDept.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " select idDepartamento,Descripcion from RolDepartamento ");
			this.cmbconsultaB.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre From Bodega Where Activo = 1 And Factura = 1 Order By Nombre");
			this.uGridContacto.DisplayLayout.Bands[0].Columns["Email"].Hidden = false;
			this.uGridContacto.DisplayLayout.Bands[0].Columns["celular"].Hidden = true;
			this.uGridContacto.DisplayLayout.Bands[0].Columns["Observaciones"].Hidden = true;
			this.uGridContacto.DisplayLayout.Bands[0].Columns["Ip"].Hidden = true;
			this.uGridContacto.DisplayLayout.Bands[0].Columns["Direccion"].Hidden = true;
      
			this.btnConsultar_Click(sender, e);

		}

		

		private void txtPersonal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtExtension.Focus();
		}

		private void txtemail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCelular.Focus();
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			#region Habilita controles
			this.cmbDepartamento.Enabled = true;
			this.txtPersonal.Enabled = true;
			this.txtemail.Enabled = true;
			this.txtExtension.Enabled = true;
		
			this.opdAlmacen.Enabled = true;
			#endregion Habilita controles

		
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;

      this.uGridContacto.Enabled = false;		
			this.ultraOptionSet1.Enabled= false;
			this.cmbconsultaB.Enabled = false;
			this.cmbConsultaDept.Enabled = false;

			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;	
	
	
			bNuevo = true;
			bEdicion = true;
			
			this.cmbDepartamento.Focus();
			
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;
			#region vaciar  campos
		  this.opdAlmacen.Value = 1;
			this.uGridContacto.Enabled = false;	
			this.cmbBodega.Visible = false;
			this.cmbDepartamento.Visible = true;
			this.txtidExtensiones.Value =0;
			this.cmbDepartamento.Value = System.DBNull.Value;
			this.txtPersonal.Text="";
			this.txtExtension.Value=0;
      this.txtemail.Text ="";
			this.txtIp.Text = "";
			this.txtidExtensiones.Value = 0;
			iBodega = 0;
			sDireccion="";
		
			this.txtCelular.Text = "";
			this.txtdireccion.Text = "";
		
			#endregion vaciar  campos

      #region bloquear campos
			
			this.cmbDepartamento.Enabled = false;
			this.txtPersonal.Enabled =  false;
			this.txtExtension.Enabled = false;;
			this.txtemail.Enabled = false;
			this.txtIp.Enabled = false;
			this.txtIp.Enabled = false;
			this.txtCelular.Enabled = false;
			this.txtdireccion.Enabled = false;
			this.opdAlmacen.Enabled = false;

			#endregion bloquear campos

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.uGridContacto.Enabled = true;	
			this.ultraOptionSet1.Enabled= true;
			this.cmbconsultaB.Enabled = true;
			this.cmbConsultaDept.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
		}
    
		public void almacen ()
		{
		 this.txtIp.Enabled = true;
     this.txtCelular.Enabled = true;
     this.txtdireccion.Enabled = true;
		 this.cmbBodega.Enabled= true;
     this.cmbBodega.Visible = true;
		 this.cmbDepartamento.Visible = false;
			
		}

		private void cmbDepartamento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.txtPersonal.Focus();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			int iConsulta =0;
			if((int)this.ultraOptionSet1.Value == 1)
			{
				if (this.cmbConsultaDept.ActiveRow != null) iConsulta = (int)this.cmbConsultaDept.Value;
				
			}
			else
			{
			if (this.cmbconsultaB.ActiveRow != null) iConsulta = (int)this.cmbconsultaB.Value;
			}
			string sSQL = string.Format("Exec ConsultaGeneralExtensiones {0},{1}",(int)this.ultraOptionSet1.Value, iConsulta);
			this.uGridContacto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			#region Habilita Campos
			this.opdAlmacen.Enabled = true;
			this.cmbDepartamento.Enabled= true;
			this.txtPersonal.Enabled = true;
			this.txtExtension.Enabled = true;
			this.txtemail.Enabled = true;
      
			if((int)this.opdAlmacen.Value == 2)
			{
				
				this.txtCelular.Enabled = true;
				this.txtdireccion.Enabled = true;
				this.txtIp.Enabled = true;
				this.cmbBodega.Enabled= true;
			}
			else
			{
			this.cmbDepartamento.Enabled= true;
			}
			
			#endregion Habilita campos

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.ultraOptionSet1.Enabled= false;
			this.uGridContacto.Enabled = false;		
			this.cmbconsultaB.Enabled = false;
			this.cmbConsultaDept.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			
			bNuevo = false;
			bEdicion = true;

			this.cmbDepartamento.Focus();
		}

		private void txtExtension_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		if(e.KeyCode == Keys.Enter) this.txtemail.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		
		}
     
		private void btnGuardar_Click(object sender, System.EventArgs e)
		{

			#region Validacion
			if( (int)this.opdAlmacen.Value == 2)
			{
				if (!Validacion.vbValidaTelefono(this.txtCelular, 2, false)) return;
				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione el Local ")) return;
				iBodega = (int) this.cmbBodega.Value;
					

				sDireccion = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select direccion from Bodega where Bodega = {0} ",
					(int)this.cmbBodega.Value ));		
			}
			else
			{
			if (!Validacion.vbComboVacio(this.cmbDepartamento, "Seleccione la ciudad el Departamento ")) return;
				
				iBodega = (int) this.cmbDepartamento.Value;
			}
			
			if (!Validacion.vbTexto(this.txtPersonal, 2, 100, "Personal")) return;

			if (this.txtemail.Text.ToString().Trim().Length > 0)
			{
				if (!Validacion.bEMail(this.txtemail.Text.ToString())) return;
			}
			if (!Validacion.vbCampoEnteroVacio(this.txtExtension,"Ingrese la extension",1,10000)) return;

		

     #endregion Validacion

			#region Guarda Matriz
			string sSQL = string.Format("Exec GuardaExtensiones {0},{1},'{2}','{3}','{4}','{5}','{6}','{7}','{8}',{9}",
				(int)this.txtidExtensiones.Value, iBodega, 	this.txtPersonal.Text.ToString(),
				this.txtemail.Text.ToString(), (int)this.txtExtension.Value, this.txtIp.Text.ToString(), this.txtCelular.Text.ToString(),
				sDireccion, this.txtdireccion.Text.ToString(), (int)this.opdAlmacen.Value);
			this.txtidExtensiones.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQL);
			#endregion Guarda Matriz


			this.cmbDepartamento.Enabled = false;
			this.txtPersonal.Enabled =  false;
			this.txtExtension.Enabled = false;;
			this.txtemail.Enabled = false;
			this.txtIp.Enabled = false;
			this.txtCelular.Enabled = false;
			this.txtdireccion.Enabled = false;
			this.opdAlmacen.Enabled = false;
			this.cmbBodega.Enabled = false;
     this.uGridContacto.Enabled = true;		
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.ultraOptionSet1.Enabled = true;



			this.uGridContacto.Enabled = true;		
			this.ultraOptionSet1.Enabled= true;
			this.cmbconsultaB.Enabled = true;
			this.cmbConsultaDept.Enabled = true;

      this.btnConsultar_Click(sender, e);
			bNuevo = false;
			bEdicion = false;

		}

		private void opdAlmacen_ValueChanged(object sender, System.EventArgs e)
		{
			if( !bEdicion)return;
			if( (int)this.opdAlmacen.Value == 2)
			{
				almacen ();
				this.cmbDepartamento.Value = 0;
			
			}
			else
			{
				this.txtIp.Enabled = false;
				this.txtCelular.Enabled = false;
				this.txtdireccion.Enabled = false;
				this.cmbBodega.Visible = false;
				this.cmbDepartamento.Visible = true;

				this.txtIp.Text = "";
				this.txtCelular.Text = "";
				this.txtdireccion.Text = "";
				this.cmbBodega.Value = 0;
				
				
			}
			
		}

		private void ultraOptionSet1_ValueChanged(object sender, System.EventArgs e)
		{
//			this.btnCancelar_Click(sender, e);
		
			if( (int)this.ultraOptionSet1.Value == 2)
			{
				
        diseñoGrid();
//				this.uGridContacto.DisplayLayout.Bands[0].Columns["Email"].Hidden = true;
//				this.uGridContacto.DisplayLayout.Bands[0].Columns["celular"].Hidden = false;
//				this.uGridContacto.DisplayLayout.Bands[0].Columns["Observaciones"].Hidden = true;
//				this.uGridContacto.DisplayLayout.Bands[0].Columns["IP"].Hidden = false;
//				this.uGridContacto.DisplayLayout.Bands[0].Columns["Direccion"].Hidden = false;

				this.cmbConsultaDept.Visible = false;
				this.cmbconsultaB.Visible = true;
					this.cmbBodega.Visible = true;
				this.cmbBodega.Enabled = false;
				this.cmbDepartamento.Visible = false;
				

			
			}
			else
			{
				 diseñoGrid();
//				this.uGridContacto.DisplayLayout.Bands[0].Columns["Email"].Hidden = false;
//				this.uGridContacto.DisplayLayout.Bands[0].Columns["celular"].Hidden = true;
//				this.uGridContacto.DisplayLayout.Bands[0].Columns["Observaciones"].Hidden = true;
//				this.uGridContacto.DisplayLayout.Bands[0].Columns["IP"].Hidden = true;
//				this.uGridContacto.DisplayLayout.Bands[0].Columns["Direccion"].Hidden = true;
				this.cmbConsultaDept.Visible = true;
				this.cmbconsultaB.Visible = false;	
				this.cmbDepartamento.Visible = true;
				this.cmbBodega.Visible = false;
			
			}

			this.btnConsultar_Click(sender, e);
		Vaciarcampos();
		}
		public void Vaciarcampos()
												{
													this.txtidExtensiones.Value =0;
													this.cmbDepartamento.Value = System.DBNull.Value;
			                    this.cmbBodega.Value =  System.DBNull.Value;
													this.txtPersonal.Text="";
													this.txtExtension.Value=0;
													this.txtemail.Text ="";
													this.txtIp.Text = "";
													this.txtidExtensiones.Value = 0;
													iBodega = 0;
													sDireccion="";
													this.txtCelular.Text = "";
													this.txtdireccion.Text = "";
												}
		public void diseñoGrid ()
		{
//			this.uGridContacto.DisplayLayout.Bands[0].Columns["Area"].Width = 100;
//			this.uGridContacto.DisplayLayout.Bands[0].Columns["Personal"].Width = 120;
//			this.uGridContacto.DisplayLayout.Bands[0].Columns["Ext"].Width = 35;
//			this.uGridContacto.DisplayLayout.Bands[0].Columns["celular"].Width = 75;
//			this.uGridContacto.DisplayLayout.Bands[0].Columns["IP"].Width = 80;

		 }

		private void txtCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtIp_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtCelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtIp.Focus();
		}

		private void txtIp_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtdireccion.Focus();
		}

		private void uGridContacto_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			#region Capturar datos del datagrid a los cajas de texto
			if(!bEdicion)
			{
				if (this.uGridContacto.ActiveRow.Cells["idExtension"].Value == DBNull.Value) return;

				this.txtidExtensiones.Value = (int)this.uGridContacto.ActiveRow.Cells["idExtension"].Value;
				this.opdAlmacen.Value = (int)this.uGridContacto.ActiveRow.Cells["idMatriz"].Value;

				if ((int) this.opdAlmacen.Value == 2)
				{
         this.cmbBodega.Value = (int)this.uGridContacto.ActiveRow.Cells["idDepartamento"].Value;
					this.cmbBodega.Visible = true;
					this.cmbDepartamento.Visible = false;
					this.cmbBodega.Enabled = false;
				}
				else
				{
					this.cmbDepartamento.Value = (int)this.uGridContacto.ActiveRow.Cells["idDepartamento"].Value;
					this.cmbBodega.Visible = false;
					this.cmbDepartamento.Visible = true;
					this.cmbDepartamento.Enabled = false;
				}
				this.txtPersonal.Text = this.uGridContacto.ActiveRow.Cells["Personal"].Value.ToString();
        this.txtemail.Text = this.uGridContacto.ActiveRow.Cells["Email"].Value.ToString();
				this.txtExtension.Value = (int)this.uGridContacto.ActiveRow.Cells["Ext"].Value;
				this.txtIp.Text = this.uGridContacto.ActiveRow.Cells["IP"].Value.ToString();
				this.txtCelular.Text = this.uGridContacto.ActiveRow.Cells["celular"].Value.ToString();
				this.txtdireccion.Text = this.uGridContacto.ActiveRow.Cells["Observaciones"].Value.ToString();
			}
			this.btnNuevo.Enabled = true;
			this.btnEditar.Enabled = true;			
			this.btnCancelar.Enabled = true;
		 #endregion Capturar datos del datagrid a los cajas de texto
		}

		private void lblPersonal_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmMatriz_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidExtensiones.Value > 0)
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

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtPersonal.Focus();
		}
		
		class validar 
		{
			public static void Solonumeros (KeyPressEventArgs Nu)
			{
				 
					
				if(char.IsDigit(Nu.KeyChar)||Nu.KeyChar == 46 )
				{
			
					Nu.Handled = false;
			
				}

				else if (char.IsControl(Nu.KeyChar))
				{
			
					Nu.Handled = false;			
				} 
				else
				{
					Nu.Handled= true;
				}
			}
	
	
		}
	}
}
