using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimientoPerTecnicoCrear.
	/// </summary>
	public class frmRequerimientoPerTecnicoCrear : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private System.Windows.Forms.Label lblAsignado;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button button2;
		public System.Windows.Forms.Label label2;
		public System.Windows.Forms.Label label12;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.Button button4;
		private System.Windows.Forms.Button button1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreCompleto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtUsuario;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbUsuario;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRol;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoConfiguracion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamentoCon;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoUsuarioCon;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoRolCon;
		private System.Windows.Forms.Label label4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoPerTecnicoCrear()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreCompleto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreCompleto");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoPerTecnicoCrear));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoConfiguracion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Obligatorio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TamanoMinimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TamanoMaximo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Excepcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mensaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Color1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Color2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Color3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoDepartamentoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.cmbUsuario = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoUsuarioCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbRol = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoRolCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtNombreCompleto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtUsuario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.button4 = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.udsRequerimientoConfiguracion = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoUsuarioCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRol)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolCon)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreCompleto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoConfiguracion)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDepartamento
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamento.Appearance = appearance1;
			this.cmbDepartamento.CausesValidation = false;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento.DataSource = this.udsRequerimientoDepartamentoCon;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 218;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(128, 24);
			this.cmbDepartamento.MaxDropDownItems = 30;
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(200, 21);
			this.cmbDepartamento.TabIndex = 1041;
			this.cmbDepartamento.ValueMember = "idRequerimientoDepartamento";
			this.cmbDepartamento.ValueChanged += new System.EventHandler(this.cmbDepartamento_ValueChanged);
			// 
			// udsRequerimientoDepartamentoCon
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsRequerimientoDepartamentoCon.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn1,
																																								ultraDataColumn2});
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(16, 24);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(96, 23);
			this.lblAsignado.TabIndex = 1042;
			this.lblAsignado.Text = "* Departamento:";
			// 
			// cmbUsuario
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbUsuario.Appearance = appearance2;
			this.cmbUsuario.CausesValidation = false;
			this.cmbUsuario.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbUsuario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUsuario.DataSource = this.udsRequerimientoUsuarioCon;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Nombre completo";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 208;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbUsuario.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbUsuario.DisplayMember = "NombreCompleto";
			this.cmbUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUsuario.Location = new System.Drawing.Point(128, 72);
			this.cmbUsuario.MaxDropDownItems = 30;
			this.cmbUsuario.Name = "cmbUsuario";
			this.cmbUsuario.Size = new System.Drawing.Size(200, 21);
			this.cmbUsuario.TabIndex = 1043;
			this.cmbUsuario.ValueMember = "idUsuario";
			this.cmbUsuario.ValueChanged += new System.EventHandler(this.cmbUsuario_ValueChanged);
			// 
			// udsRequerimientoUsuarioCon
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.udsRequerimientoUsuarioCon.Band.Columns.AddRange(new object[] {
																																					 ultraDataColumn3,
																																					 ultraDataColumn4});
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(96, 23);
			this.label1.TabIndex = 1044;
			this.label1.Text = "* Usuario:";
			// 
			// cmbRol
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRol.Appearance = appearance3;
			this.cmbRol.CausesValidation = false;
			this.cmbRol.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRol.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRol.DataSource = this.udsRequerimientoRolCon;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.Caption = "Descripción";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 218;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbRol.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbRol.DisplayMember = "Descripcion";
			this.cmbRol.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRol.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRol.Location = new System.Drawing.Point(128, 208);
			this.cmbRol.MaxDropDownItems = 30;
			this.cmbRol.Name = "cmbRol";
			this.cmbRol.Size = new System.Drawing.Size(200, 21);
			this.cmbRol.TabIndex = 1056;
			this.cmbRol.ValueMember = "idRequerimientoRol";
			this.cmbRol.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.idRol_InitializeLayout);
			// 
			// udsRequerimientoRolCon
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.udsRequerimientoRolCon.Band.Columns.AddRange(new object[] {
																																			 ultraDataColumn5,
																																			 ultraDataColumn6});
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(16, 208);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(96, 23);
			this.label3.TabIndex = 1057;
			this.label3.Text = "* Rol:";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.lblAsignado);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.cmbDepartamento);
			this.groupBox1.Controls.Add(this.cmbRol);
			this.groupBox1.Controls.Add(this.cmbUsuario);
			this.groupBox1.Controls.Add(this.txtNombreCompleto);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.txtUsuario);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Location = new System.Drawing.Point(16, 16);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(352, 248);
			this.groupBox1.TabIndex = 1058;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información usuario";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 168);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(75, 16);
			this.label2.TabIndex = 1055;
			this.label2.Text = "* Usuario cod:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(16, 120);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(104, 16);
			this.label12.TabIndex = 1053;
			this.label12.Text = "* Nombre completo:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombreCompleto
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreCompleto.Appearance = appearance4;
			this.txtNombreCompleto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombreCompleto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreCompleto.Location = new System.Drawing.Point(128, 120);
			this.txtNombreCompleto.Name = "txtNombreCompleto";
			this.txtNombreCompleto.Size = new System.Drawing.Size(200, 21);
			this.txtNombreCompleto.TabIndex = 1052;
			this.txtNombreCompleto.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// txtUsuario
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUsuario.Appearance = appearance5;
			this.txtUsuario.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtUsuario.Location = new System.Drawing.Point(128, 168);
			this.txtUsuario.Name = "txtUsuario";
			this.txtUsuario.Size = new System.Drawing.Size(200, 21);
			this.txtUsuario.TabIndex = 1054;
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(16, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 23);
			this.label4.TabIndex = 1042;
			this.label4.Text = "* Departamento:";
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(-80, 148);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 1059;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(397, 148);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1060;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.button2.CausesValidation = false;
			this.button2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button2.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.ImageIndex = 62;
			this.button2.Location = new System.Drawing.Point(477, 0);
			this.button2.Name = "button2";
			this.button2.TabIndex = 1062;
			this.button2.Text = "Salir";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button3.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button3.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button3.Location = new System.Drawing.Point(16, 280);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(76, 23);
			this.button3.TabIndex = 1063;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// button4
			// 
			this.button4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.button4.CausesValidation = false;
			this.button4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button4.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button4.ImageIndex = 62;
			this.button4.Location = new System.Drawing.Point(488, 280);
			this.button4.Name = "button4";
			this.button4.TabIndex = 1064;
			this.button4.Text = "Salir";
			this.button4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.CausesValidation = false;
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(104, 280);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(80, 23);
			this.button1.TabIndex = 1065;
			this.button1.Text = "Salir";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// udsRequerimientoConfiguracion
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(bool);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(bool);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			this.udsRequerimientoConfiguracion.Band.Columns.AddRange(new object[] {
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
																																							ultraDataColumn17,
																																							ultraDataColumn18});
			// 
			// frmRequerimientoPerTecnicoCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(392, 318);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.button4);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoPerTecnicoCrear";
			this.Text = "Configuracion usuarios";
			this.Load += new System.EventHandler(this.frmRequerimientoPerTecnicoCrear_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoUsuarioCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRol)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolCon)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNombreCompleto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoConfiguracion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoPerTecnicoCrear_Load(object sender, System.EventArgs e)
		{
			//Cargar departamento
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
			//Cargar Usuario
			RequerimientoUsuarioCon();
			//Cargar Rol
			RequerimientoRolCon();
		}

		private void RequerimientoRolCon()
		{
			string squery = String.Format("Exec RequerimientoRolCon");
			this.cmbRol.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		private void RequerimientoUsuarioCon()
		{
			int idRequerimientoDepartamento = 0;

			try
			{
				idRequerimientoDepartamento = (int)this.cmbDepartamento.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			string squery = String.Format("Exec RequerimientoUsuarioCon {0}", idRequerimientoDepartamento);
			this.cmbUsuario.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void idRol_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbDepartamento_ValueChanged(object sender, System.EventArgs e)
		{
			RequerimientoUsuarioCon();
			LimpiarDatosUsuario();
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			RequerimientoPerTecnicoNuevo();

		}

		private void cmbUsuario_ValueChanged(object sender, System.EventArgs e)
		{
			RequerimientoUsuarioModel();
		}

		//		private void RequerimientoUsuarioModel()
		//		{
		//			int idUsuario = 0;
		//				try
		//				{
		//					idUsuario = (int)this.cmbUsuario.Value;
		//				}
		//				catch(Exception ex)
		//				{
		//					string sms = ex.Message;
		//				}
		//
		//			if(idUsuario > 0)
		//			{
		//				string squery = String.Format("Exec RequerimientoUsuarioModel {0}, '{1}'", idUsuario, "NombreCompleto");
		//				this.txtNombreCompleto.Value = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
		//				squery = String.Format("Exec RequerimientoUsuarioModel {0}, '{1}'", idUsuario, "Nombre");
		//				this.txtUsuario.Value = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
		//			}
		//		}

		private void RequerimientoUsuarioModel()
		{
			LimpiarDatosUsuario();
			int idUsuario = 0;
			try
			{
				idUsuario = (int)this.cmbUsuario.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			if(idUsuario > 0)
			{
				string squery = String.Format("Exec RequerimientoUsuarioModel {0}, '{1}'", idUsuario, "NombreCompleto");
				this.txtNombreCompleto.Value = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
				squery = String.Format("Exec RequerimientoUsuarioModel {0}, '{1}'", idUsuario, "Nombre");
				this.txtUsuario.Value = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, squery);
			}
		}

		private void LimpiarDatosUsuario()
		{
			this.txtNombreCompleto.Value = string.Empty;
			this.txtUsuario.Value = string.Empty;
		}

		//RequerimientoPerTecnicoNuevo
		private void RequerimientoPerTecnicoNuevo()
		{
			#region Validación
			if (!Validacion.vbComboVacio(this.cmbDepartamento, "Seleccione el departamento ")) return;
			if (!Validacion.vbComboVacio(this.cmbUsuario, "Seleccione el usuario ")) return;
			if (!Validacion.vbTexto(this.txtNombreCompleto, 10, 100, "Ingrese el nombre completo")) return;
			if (!Validacion.vbTexto(this.txtUsuario, 5, 100, "Ingrese código del usuario")) return;
			if (!Validacion.vbComboVacio(this.cmbRol, "Seleccione el rol ")) return;
			int idUsuario = (int)this.cmbUsuario.Value;
			int idRequerimientoDepartamento = (int)this.cmbDepartamento.Value;
			// Llamar al método y obtener el mensaje
			if (RequerimientoPerTecnicoVal(idUsuario, idRequerimientoDepartamento))return;
			#endregion Validación

			string connectionString = frmRequerimiento.sconexionPoint; // Cambia esto con tu cadena de conexión real

			try
			{
				using (SqlConnection connection = new SqlConnection(connectionString))
				{
					string squery = "Exec RequerimientoPerTecnicoNuevo @idRequerimientoPerTecnico, @NombreCompleto, @idPersonal, @idUsuario, @Usuario, @idRequerimientoRol, @Estado, @idRequerimientoDepartamento";

					using (SqlCommand cmd = new SqlCommand(squery, connection))
					{
						// Agregar parámetros de manera explícita
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoPerTecnico", System.Data.SqlDbType.Int)).Value = 0;
						cmd.Parameters.Add(new SqlParameter("@NombreCompleto", System.Data.SqlDbType.VarChar, 100)).Value = this.txtNombreCompleto.Text;
						cmd.Parameters.Add(new SqlParameter("@idPersonal", System.Data.SqlDbType.Int)).Value = (int)this.cmbUsuario.Value;
						cmd.Parameters.Add(new SqlParameter("@idUsuario", System.Data.SqlDbType.Int)).Value = (int)this.cmbUsuario.Value;
						cmd.Parameters.Add(new SqlParameter("@Usuario", System.Data.SqlDbType.VarChar, 100)).Value = this.txtUsuario.Text;
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoRol", System.Data.SqlDbType.Int)).Value = (int)this.cmbRol.Value;
						cmd.Parameters.Add(new SqlParameter("@Estado", System.Data.SqlDbType.Bit)).Value = 1;
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoDepartamento", System.Data.SqlDbType.Int)).Value = (int)this.cmbDepartamento.Value;

						connection.Open();
						cmd.ExecuteNonQuery();
					}
				}

				// Mostrar mensaje de éxito
				MessageBox.Show("Usuario registrado correctamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.DialogResult = DialogResult.OK;
			}
			catch (Exception ex)
			{
				// Manejo de errores
				MessageBox.Show("Ocurrió un error al registrar el usuario: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		public bool RequerimientoPerTecnicoVal(int idUsuario, int idRequerimientoDepartamento)
		{
			bool existe = false;
			string squery = String.Format("Exec RequerimientoPerTecnicoVal {0}, {1}, 0", idUsuario, idRequerimientoDepartamento);
			existe = Funcion.bEjecucion(frmRequerimiento.sconexionPoint, squery, false);

			if (existe)
			{
				// Mostrar mensaje de alerta si ya existe
				MessageBox.Show("No se puede ingresar el usuario, ya existe.", "Alerta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
			return existe;
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

	}
}
