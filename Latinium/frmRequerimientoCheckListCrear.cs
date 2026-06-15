using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data; // Para CommandType
using System.Data.SqlClient; // Para SqlConnection, SqlCommand, y SqlParameter

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimientoCheckListCrear.
	/// </summary>
	public class frmRequerimientoCheckListCrear : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblAsignado;
		public System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAreaAtencion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbAsignaA;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoDepartamentoCon;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoAreaAtencionTipoActivosObtener;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequerimientoPerTecnicoActivosCon;
		private System.Windows.Forms.Label lblAsignaA;
		int idRequerimientoDepartamento = 0;

		public frmRequerimientoCheckListCrear(int IdRequerimientoDepartamento)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			idRequerimientoDepartamento = IdRequerimientoDepartamento;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoCheckListCrear));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoAreaAtencionTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoAreaAtencionTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Departamento");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoPerTecnico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoPerTecnico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRequerimientoDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Departamento");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			this.button1 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblAsignaA = new System.Windows.Forms.Label();
			this.lblAsignado = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoDepartamentoCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbAreaAtencion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoAreaAtencionTipoActivosObtener = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbAsignaA = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsRequerimientoPerTecnicoActivosCon = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAreaAtencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoAreaAtencionTipoActivosObtener)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignaA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPerTecnicoActivosCon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			this.SuspendLayout();
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
			this.button1.Location = new System.Drawing.Point(384, 80);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(80, 23);
			this.button1.TabIndex = 1068;
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
			this.button3.Size = new System.Drawing.Size(80, 23);
			this.button3.TabIndex = 1067;
			this.button3.Text = "Guardar";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.lblAsignaA);
			this.groupBox1.Controls.Add(this.lblAsignado);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.cmbDepartamento);
			this.groupBox1.Controls.Add(this.cmbAreaAtencion);
			this.groupBox1.Controls.Add(this.cmbAsignaA);
			this.groupBox1.Controls.Add(this.txtDescripcion);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Location = new System.Drawing.Point(16, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(352, 184);
			this.groupBox1.TabIndex = 1066;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Información";
			// 
			// lblAsignaA
			// 
			this.lblAsignaA.Location = new System.Drawing.Point(16, 144);
			this.lblAsignaA.Name = "lblAsignaA";
			this.lblAsignaA.Size = new System.Drawing.Size(104, 23);
			this.lblAsignaA.TabIndex = 1057;
			this.lblAsignaA.Text = "* Asigna a:";
			// 
			// lblAsignado
			// 
			this.lblAsignado.Location = new System.Drawing.Point(16, 24);
			this.lblAsignado.Name = "lblAsignado";
			this.lblAsignado.Size = new System.Drawing.Size(96, 23);
			this.lblAsignado.TabIndex = 1042;
			this.lblAsignado.Text = "* Departamento:";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(16, 104);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(89, 16);
			this.label12.TabIndex = 1053;
			this.label12.Text = "* Requerimiento:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			this.cmbDepartamento.Enabled = false;
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
			// cmbAreaAtencion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAreaAtencion.Appearance = appearance2;
			this.cmbAreaAtencion.CausesValidation = false;
			this.cmbAreaAtencion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbAreaAtencion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAreaAtencion.DataSource = this.udsRequerimientoAreaAtencionTipoActivosObtener;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Descripción";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 218;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbAreaAtencion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbAreaAtencion.DisplayMember = "Descripcion";
			this.cmbAreaAtencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAreaAtencion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAreaAtencion.Location = new System.Drawing.Point(128, 64);
			this.cmbAreaAtencion.MaxDropDownItems = 30;
			this.cmbAreaAtencion.Name = "cmbAreaAtencion";
			this.cmbAreaAtencion.Size = new System.Drawing.Size(200, 21);
			this.cmbAreaAtencion.TabIndex = 1056;
			this.cmbAreaAtencion.ValueMember = "idRequerimientoAreaAtencionTipo";
			// 
			// udsRequerimientoAreaAtencionTipoActivosObtener
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.udsRequerimientoAreaAtencionTipoActivosObtener.Band.Columns.AddRange(new object[] {
																																															 ultraDataColumn3,
																																															 ultraDataColumn4,
																																															 ultraDataColumn5});
			// 
			// cmbAsignaA
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbAsignaA.Appearance = appearance3;
			this.cmbAsignaA.CausesValidation = false;
			this.cmbAsignaA.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbAsignaA.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAsignaA.DataSource = this.udsRequerimientoPerTecnicoActivosCon;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.Caption = "Nombre completo";
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 208;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 6;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 8;
			ultraGridColumn14.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbAsignaA.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbAsignaA.DisplayMember = "Nombre";
			this.cmbAsignaA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbAsignaA.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAsignaA.Location = new System.Drawing.Point(128, 144);
			this.cmbAsignaA.MaxDropDownItems = 30;
			this.cmbAsignaA.Name = "cmbAsignaA";
			this.cmbAsignaA.Size = new System.Drawing.Size(200, 21);
			this.cmbAsignaA.TabIndex = 1043;
			this.cmbAsignaA.ValueMember = "idRequerimientoPerTecnico";
			// 
			// udsRequerimientoPerTecnicoActivosCon
			// 
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(bool);
			ultraDataColumn13.DataType = typeof(int);
			this.udsRequerimientoPerTecnicoActivosCon.Band.Columns.AddRange(new object[] {
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
			// txtDescripcion
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance4;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(128, 104);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(200, 21);
			this.txtDescripcion.TabIndex = 1052;
			this.txtDescripcion.ValueChanged += new System.EventHandler(this.txtNombreCompleto_ValueChanged);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 64);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(104, 23);
			this.label1.TabIndex = 1044;
			this.label1.Text = "* Área de atención";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(16, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 23);
			this.label4.TabIndex = 1042;
			this.label4.Text = "* Departamento:";
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
			this.btnNuevo.TabIndex = 1069;
			this.btnNuevo.Text = "&Nueva Área";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click_1);
			// 
			// frmRequerimientoCheckListCrear
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 206);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmRequerimientoCheckListCrear";
			this.Text = "Checklist";
			this.Load += new System.EventHandler(this.frmRequerimientoCheckListCrear_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoDepartamentoCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAreaAtencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoAreaAtencionTipoActivosObtener)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAsignaA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsRequerimientoPerTecnicoActivosCon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRequerimientoCheckListCrear_Load(object sender, System.EventArgs e)
		{
			//Cargar departamento
			frmRequerimientoConfiguracion.RequerimientoDepartamentoCon(this.cmbDepartamento);
			this.cmbDepartamento.Value = idRequerimientoDepartamento;
			//Cargar Usuario
			RequerimientoUsuarioCon();
			//Atención area tipo
			RequerimientoAreaAtencionTipoActivosObtener();
			MostrarAsignadoA();
			ValorDefaultUsuarioAsignado();

		}

		private void ValorDefaultUsuarioAsignado()
		{
			if(!frmSeguimiento.RequerimientoDepartamentoTabMostrar((int)this.cmbDepartamento.Value, 0, "Asignacion"))
			{
				string squery = String.Format("Select top 1 idRequerimientoPertecnico From RequerimientoPertecnico where idRequerimientoDepartamento = {0}", idRequerimientoDepartamento);
				this.cmbAsignaA.Value = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, squery, false);
			}
		}
		
		private void RequerimientoUsuarioCon()
		{
			int getidRequerimientoDepartamento = 0;

			try
			{
				getidRequerimientoDepartamento = (int)this.cmbDepartamento.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			string squery = String.Format("Exec RequerimientoPerTecnicoActivosCon {0}", getidRequerimientoDepartamento);
			this.cmbAsignaA.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		private void RequerimientoAreaAtencionTipoActivosObtener()
		{
			int getidRequerimientoDepartamento = 0;

			try
			{
				getidRequerimientoDepartamento = (int)this.cmbDepartamento.Value;
			}
			catch(Exception ex)
			{
				string sms = ex.Message;
			}

			string squery = String.Format("Exec RequerimientoAreaAtencionTipoActivosObtener {0}", getidRequerimientoDepartamento);
			this.cmbAreaAtencion.DataSource = Funcion.dsEjecucion(frmRequerimiento.sconexionPoint,squery);
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnNuevo_Click_1(object sender, System.EventArgs e)
		{
			using (frmRequerimientoAreaAtencionTipoCrear frmraatc = new frmRequerimientoAreaAtencionTipoCrear(idRequerimientoDepartamento))
			{		
				if (DialogResult.OK == frmraatc.ShowDialog())
				{
				}
				RequerimientoAreaAtencionTipoActivosObtener();
			}
		}

		private void txtNombreCompleto_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			RequerimientoAreaAtencionInsertar();
		}

		private void RequerimientoAreaAtencionInsertar()
		{
			#region Validación
			if (!Validacion.vbComboVacio(this.cmbDepartamento, "Seleccione el departamento")) return;
			if (!Validacion.vbComboVacio(this.cmbAreaAtencion, "Seleccione el área de atención")) return;
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Ingrese el requerimiento")) return;
			if (!Validacion.vbComboVacio(this.cmbAsignaA, "Seleccione al usuario asignado")) return;
			#endregion Validación

			// Declaración de variables
			int idRequerimientoAreaAtencionTipo = (int)this.cmbAreaAtencion.Value;
			bool Seleccionar = false;  // O ajusta el valor según el comportamiento deseado
			int idRequerimientoPerTecnico = (int)this.cmbAsignaA.Value;
			bool Estado = true;  // Asumido como 1 (activo)
			int getidRequerimientoDepartamento = (int)this.cmbDepartamento.Value;
			string Descripcion = this.txtDescripcion.Text;

			try
			{
				// Conexión a la base de datos
				using (SqlConnection conn = new SqlConnection(frmRequerimiento.sconexionPoint))
				{
					conn.Open();

					using (SqlCommand cmd = new SqlCommand("RequerimientoAreaAtencionInsertar", conn))
					{
						cmd.CommandType = CommandType.StoredProcedure;

						// Parámetros del procedimiento almacenado
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoAreaAtencionTipo", SqlDbType.Int)).Value = idRequerimientoAreaAtencionTipo;
						cmd.Parameters.Add(new SqlParameter("@Seleccionar", SqlDbType.Bit)).Value = Seleccionar;
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoPerTecnico", SqlDbType.Int)).Value = idRequerimientoPerTecnico;
						cmd.Parameters.Add(new SqlParameter("@Estado", SqlDbType.Bit)).Value = Estado;
						cmd.Parameters.Add(new SqlParameter("@idRequerimientoDepartamento", SqlDbType.Int)).Value = getidRequerimientoDepartamento;
						cmd.Parameters.Add(new SqlParameter("@Descripcion", SqlDbType.VarChar, 100)).Value = Descripcion;

						// Ejecución del procedimiento almacenado
						cmd.ExecuteNonQuery();
					}
				}

				// Mostrar mensaje de éxito
				MessageBox.Show("Registro insertado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.DialogResult = DialogResult.OK;

			}
			catch (Exception ex)
			{
				// Mostrar mensaje de error
				MessageBox.Show("Error al insertar el registro: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void cmbDepartamento_ValueChanged(object sender, System.EventArgs e)
		{
			MostrarAsignadoA();
		}

		private void MostrarAsignadoA()
		{
			bool Asignacion = false;
				try
				{
					Asignacion = frmSeguimiento.RequerimientoDepartamentoTabMostrar((int)this.cmbDepartamento.Value, 0, "Asignacion");
				}
				catch(Exception ex)
				{
					string sms = ex.Message;
				}

			if(Asignacion)
			{
				this.cmbAsignaA.Visible = true;
				this.lblAsignaA.Visible = true;
			}
			else
			{
				this.cmbAsignaA.Visible = false;
				this.lblAsignaA.Visible = false;
			}
		}
	}
}
