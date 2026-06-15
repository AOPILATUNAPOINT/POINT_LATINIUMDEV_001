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
	/// Descripción breve de frmRequerimientoMensajeRolCrear.
	/// </summary>
	public class frmRequerimientoMensajeRolCrear : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnNvRol;
		private System.Windows.Forms.Button btnNvEstado;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblAsignado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRol;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoRolCon;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamentoCon;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMensaje;
		private System.Windows.Forms.Label txtMensaje;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoMensajeObtenerActivos;
		private System.Windows.Forms.Button button2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRequerimientoMensajeRolCrear()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoMensajeRolCrear));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoMensaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoMensaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoRequerimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.btnNvRol = new System.Windows.Forms.Button();
			this.btnNvEstado = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbMensaje = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoMensajeObtenerActivos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtMensaje = new System.Windows.Forms.Label();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbRol = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.udsRequerimientoRolCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsRequerimientoDepartamentoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsEstado = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.button2 = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbMensaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoMensajeObtenerActivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRol)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// btnNvRol
			// 
			this.btnNvRol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNvRol.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNvRol.CausesValidation = false;
			this.btnNvRol.Enabled = false;
			this.btnNvRol.Image = ((System.Drawing.Image)(resources.GetObject("btnNvRol.Image")));
			this.btnNvRol.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNvRol.Location = new System.Drawing.Point(384, 112);
			this.btnNvRol.Name = "btnNvRol";
			this.btnNvRol.Size = new System.Drawing.Size(80, 24);
			this.btnNvRol.TabIndex = 1083;
			this.btnNvRol.Text = "&Nuevo rol";
			this.btnNvRol.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNvRol.Click += new System.EventHandler(this.btnNvRol_Click);
			// 
			// btnNvEstado
			// 
			this.btnNvEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNvEstado.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNvEstado.CausesValidation = false;
			this.btnNvEstado.Image = ((System.Drawing.Image)(resources.GetObject("btnNvEstado.Image")));
			this.btnNvEstado.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNvEstado.Location = new System.Drawing.Point(384, 80);
			this.btnNvEstado.Name = "btnNvEstado";
			this.btnNvEstado.Size = new System.Drawing.Size(80, 24);
			this.btnNvEstado.TabIndex = 1082;
			this.btnNvEstado.Text = "&Nuevo est.";
			this.btnNvEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNvEstado.Click += new System.EventHandler(this.btnNvEstado_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(384, 48);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 24);
			this.btnNuevo.TabIndex = 1081;
			this.btnNuevo.Text = "&Nuevo dep.";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.CausesValidation = false;
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(384, 176);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(80, 23);
			this.button1.TabIndex = 1080;
			this.button1.Text = "Salir";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button3.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button3.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button3.Location = new System.Drawing.Point(384, 16);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(76, 23);
			this.button3.TabIndex = 1079;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cmbMensaje);
			this.groupBox1.Controls.Add(this.txtMensaje);
			this.groupBox1.Controls.Add(this.cmbEstado);
			this.groupBox1.Controls.Add(this.lblEstado);
			this.groupBox1.Controls.Add(this.lblAsignado);
			this.groupBox1.Controls.Add(this.cmbDepartamento);
			this.groupBox1.Controls.Add(this.cmbRol);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(352, 192);
			this.groupBox1.TabIndex = 1078;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información usuario";
			// 
			// cmbMensaje
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMensaje.Appearance = appearance1;
			this.cmbMensaje.CausesValidation = false;
			this.cmbMensaje.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMensaje.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMensaje.DataSource = this.udsRequerimientoMensajeObtenerActivos;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 218;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbMensaje.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbMensaje.DisplayMember = "Descripcion";
			this.cmbMensaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMensaje.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMensaje.Location = new System.Drawing.Point(128, 120);
			this.cmbMensaje.MaxDropDownItems = 30;
			this.cmbMensaje.Name = "cmbMensaje";
			this.cmbMensaje.Size = new System.Drawing.Size(200, 21);
			this.cmbMensaje.TabIndex = 1060;
			this.cmbMensaje.ValueMember = "idRequerimientoMensaje";
			// 
			// udsRequerimientoMensajeObtenerActivos
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(bool);
			this.udsRequerimientoMensajeObtenerActivos.Band.Columns.AddRange(new object[] {
																																											ultraDataColumn1,
																																											ultraDataColumn2,
																																											ultraDataColumn3,
																																											ultraDataColumn4});
			// 
			// txtMensaje
			// 
			this.txtMensaje.Location = new System.Drawing.Point(16, 120);
			this.txtMensaje.Name = "txtMensaje";
			this.txtMensaje.Size = new System.Drawing.Size(96, 23);
			this.txtMensaje.TabIndex = 1061;
			this.txtMensaje.Text = "* Mensaje:";
			// 
			// cmbEstado
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEstado.Appearance = appearance2;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEstado.DisplayMember = "Descripcion";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(128, 56);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(200, 21);
			this.cmbEstado.TabIndex = 1059;
			this.cmbEstado.ValueMember = "idEstadoRequerimiento";
			// 
			// lblEstado
			// 
			this.lblEstado.Location = new System.Drawing.Point(16, 56);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(96, 23);
			this.lblEstado.TabIndex = 1058;
			this.lblEstado.Text = "* Estado:";
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(16, 24);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(96, 23);
			this.lblAsignado.TabIndex = 1042;
			this.lblAsignado.Text = "* Departamento:";
			// 
			// cmbDepartamento
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDepartamento.Appearance = appearance3;
			this.cmbDepartamento.CausesValidation = false;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.Caption = "Descripción";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 218;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(128, 24);
			this.cmbDepartamento.MaxDropDownItems = 30;
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(200, 21);
			this.cmbDepartamento.TabIndex = 1041;
			this.cmbDepartamento.ValueMember = "idRequerimientoDepartamento";
			// 
			// cmbRol
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRol.Appearance = appearance4;
			this.cmbRol.CausesValidation = false;
			this.cmbRol.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRol.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.Caption = "Descripción";
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 218;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbRol.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbRol.DisplayMember = "Descripcion";
			this.cmbRol.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRol.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRol.Location = new System.Drawing.Point(128, 88);
			this.cmbRol.MaxDropDownItems = 30;
			this.cmbRol.Name = "cmbRol";
			this.cmbRol.Size = new System.Drawing.Size(200, 21);
			this.cmbRol.TabIndex = 1056;
			this.cmbRol.ValueMember = "idRequerimientoRol";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(16, 88);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(96, 23);
			this.label3.TabIndex = 1057;
			this.label3.Text = "* Rol:";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(16, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 23);
			this.label4.TabIndex = 1042;
			this.label4.Text = "* Departamento:";
			// 
			// udsRequerimientoRolCon
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.udsRequerimientoRolCon.Band.Columns.AddRange(new object[] {
																																			 ultraDataColumn5,
																																			 ultraDataColumn6});
			// 
			// udsRequerimientoDepartamentoCon
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.udsRequerimientoDepartamentoCon.Band.Columns.AddRange(new object[] {
																																								ultraDataColumn7,
																																								ultraDataColumn8});
			// 
			// udsEstado
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.udsEstado.Band.Columns.AddRange(new object[] {
																													ultraDataColumn9,
																													ultraDataColumn10});
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button2.CausesValidation = false;
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.Location = new System.Drawing.Point(384, 144);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(80, 24);
			this.button2.TabIndex = 1084;
			this.button2.Text = "&Nuevo msg.";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// frmRequerimientoMensajeRolCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 214);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.btnNvRol);
			this.Controls.Add(this.btnNvEstado);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoMensajeRolCrear";
			this.Text = "Mensaje - Rol";
			this.Load += new System.EventHandler(this.frmRequerimientoMensajeRolCrear_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbMensaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoMensajeObtenerActivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRol)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoRolCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoMensajeRolCrear_Load(object sender, System.EventArgs e)
		{
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
			frmRequerimientoConfiguracion.RequerimientoEstadoObtenerActivos(this.cmbEstado);
			frmRequerimientoConfiguracion.RequerimientoRolCon(this.cmbRol);
			frmRequerimientoConfiguracion.RequerimientoMensajeObtenerActivos(this.cmbMensaje);
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoDepartamentoCrear frmdc = new frmRequerimientoDepartamentoCrear())
			{		
				if (DialogResult.OK == frmdc.ShowDialog())
				{
				}
				//Cargar departamento
				frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
			}
		}

		private void btnNvEstado_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoEstadoCrear frmec = new frmRequerimientoEstadoCrear())
			{		
				if (DialogResult.OK == frmec.ShowDialog())
				{
				}
				//Cargar Estados Activos
				frmRequerimientoConfiguracion.RequerimientoEstadoObtenerActivos(this.cmbEstado);
			}	
		}

		private void btnNvRol_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoRolCrear frmrrc = new frmRequerimientoRolCrear())
			{		
				if (DialogResult.OK == frmrrc.ShowDialog())
				{
				}
				//Cargar Rol
				frmRequerimientoConfiguracion.RequerimientoRolCon(this.cmbRol);
			}	
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			using (frmRequerimientoMensajeCrear frmred = new frmRequerimientoMensajeCrear())
			{		
				if (DialogResult.OK == frmred.ShowDialog())
				{
				}
				frmRequerimientoConfiguracion.RequerimientoMensajeObtenerActivos(this.cmbMensaje);
			}		
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			#region Validación
			if (!Validacion.vbComboVacio(this.cmbDepartamento, "Seleccione el departamento ")) return;
			if (!Validacion.vbComboVacio(this.cmbEstado, "Seleccione el estado ")) return;
			if (!Validacion.vbComboVacio(this.cmbRol, "Seleccione el rol ")) return;
			if (!Validacion.vbComboVacio(this.cmbMensaje, "Seleccione el mensaje ")) return;
			#endregion Validación

			// Variables para almacenar los valores seleccionados de los ComboBox
			int idRol = Convert.ToInt32(cmbRol.Value);
			int idMensaje = Convert.ToInt32(cmbMensaje.Value);
			int idEstado = Convert.ToInt32(cmbEstado.Value);

			try
			{
				// Establecer la conexión con la base de datos
				using (SqlConnection connection = new SqlConnection(frmRequerimiento.sconexionPoint))
				{
					connection.Open();

					// Crear el comando para ejecutar el procedimiento almacenado
					using (SqlCommand command = new SqlCommand("RequerimientoMensajeRolInsertar", connection))
					{
						command.CommandType = CommandType.StoredProcedure;

						// Agregar los parámetros al procedimiento almacenado
						command.Parameters.Add(new SqlParameter("@idRequerimientoRol", idRol));
						command.Parameters.Add(new SqlParameter("@idRequerimientoMensaje", idMensaje));
						command.Parameters.Add(new SqlParameter("@idEstadoRequerimiento", idEstado));

						// Ejecutar el procedimiento almacenado y obtener el ID del nuevo registro
						int newId = Convert.ToInt32(command.ExecuteScalar());

						// Mostrar el ID del nuevo registro (puedes ajustarlo según lo que desees hacer)
						MessageBox.Show("Registro insertado con ID: " + newId, "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.DialogResult = DialogResult.OK;
					}
				}
			}
			catch (Exception ex)
			{
				// Manejar posibles excepciones
				MessageBox.Show("Ocurrió un error: " +ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

	}
}
