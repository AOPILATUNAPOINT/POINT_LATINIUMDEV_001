using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmUsuarios.
	/// </summary>
	public class frmUsuarios : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnGrupo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClave;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.Misc.UltraButton btEditar;
		private Infragistics.Win.Misc.UltraButton btNuevo;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		bool bNuevo = false;
		bool bModifica = false;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGrupo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConsulta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteo;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdUsuario;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkActivo;
		bool bEdicion = false;

		public frmUsuarios()
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
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Clave");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuarioBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("credito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuarioBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			this.spnGrupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.btEditar = new Infragistics.Win.Misc.UltraButton();
			this.btNuevo = new Infragistics.Win.Misc.UltraButton();
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtConsulta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtGrupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.txtIdUsuario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkTodos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkActivo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// spnGrupo
			// 
			this.spnGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnGrupo.Enabled = false;
			this.spnGrupo.Location = new System.Drawing.Point(96, 8);
			this.spnGrupo.MaskInput = "nn";
			this.spnGrupo.MaxValue = 36;
			this.spnGrupo.MinValue = 1;
			this.spnGrupo.Name = "spnGrupo";
			this.spnGrupo.PromptChar = ' ';
			this.spnGrupo.Size = new System.Drawing.Size(80, 21);
			this.spnGrupo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnGrupo.TabIndex = 3;
			this.spnGrupo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.spnGrupo_KeyPress);
			this.spnGrupo.ValueChanged += new System.EventHandler(this.spnGrupo_ValueChanged);
			// 
			// txtClave
			// 
			this.txtClave.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtClave.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClave.Enabled = false;
			this.txtClave.Location = new System.Drawing.Point(96, 72);
			this.txtClave.MaxLength = 12;
			this.txtClave.Name = "txtClave";
			this.txtClave.PasswordChar = '*';
			this.txtClave.Size = new System.Drawing.Size(208, 21);
			this.txtClave.TabIndex = 5;
			this.txtClave.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtClave_KeyPress);
			// 
			// txtNombre
			// 
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(96, 40);
			this.txtNombre.MaxLength = 20;
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(208, 21);
			this.txtNombre.TabIndex = 4;
			this.txtNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombre_KeyPress);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(36, 16);
			this.label4.TabIndex = 17;
			this.label4.Text = "Grupo";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(63, 16);
			this.label2.TabIndex = 16;
			this.label2.Text = "Contraseña";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 15;
			this.label1.Text = "Nombre";
			// 
			// btEditar
			// 
			appearance1.Image = 9;
			this.btEditar.Appearance = appearance1;
			this.btEditar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btEditar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btEditar.Enabled = false;
			this.btEditar.Location = new System.Drawing.Point(84, 114);
			this.btEditar.Name = "btEditar";
			this.btEditar.Size = new System.Drawing.Size(64, 21);
			this.btEditar.TabIndex = 8;
			this.btEditar.Text = "&Editar";
			this.btEditar.Click += new System.EventHandler(this.btEditar_Click);
			// 
			// btNuevo
			// 
			appearance2.Image = 9;
			this.btNuevo.Appearance = appearance2;
			this.btNuevo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNuevo.Location = new System.Drawing.Point(12, 114);
			this.btNuevo.Name = "btNuevo";
			this.btNuevo.Size = new System.Drawing.Size(64, 21);
			this.btNuevo.TabIndex = 2;
			this.btNuevo.Text = "&Nuevo";
			this.btNuevo.Click += new System.EventHandler(this.btNuevo_Click);
			// 
			// btSalir
			// 
			appearance3.Image = 9;
			this.btSalir.Appearance = appearance3;
			this.btSalir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSalir.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btSalir.Location = new System.Drawing.Point(300, 114);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(64, 21);
			this.btSalir.TabIndex = 9;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// btGrabar
			// 
			appearance4.Image = 9;
			this.btGrabar.Appearance = appearance4;
			this.btGrabar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrabar.Enabled = false;
			this.btGrabar.Location = new System.Drawing.Point(156, 114);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(64, 21);
			this.btGrabar.TabIndex = 6;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btCancelar
			// 
			appearance5.Image = 9;
			this.btCancelar.Appearance = appearance5;
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Enabled = false;
			this.btCancelar.Location = new System.Drawing.Point(228, 114);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(64, 21);
			this.btCancelar.TabIndex = 7;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Usuario";
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.Caption = "Grupo";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 32;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 258;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 81;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn5.Header.Caption = "Ingreso";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 65;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "A";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 32;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			ultraGridBand1.Header.Caption = "USUARIOS";
			ultraGridBand1.HeaderVisible = true;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 10F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 192);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(408, 330);
			this.ultraGrid1.TabIndex = 26;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			this.ultraGrid1.Click += new System.EventHandler(this.ultraGrid1_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn6.DataType = typeof(bool);
			ultraDataColumn6.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// txtConsulta
			// 
			this.txtConsulta.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConsulta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConsulta.Location = new System.Drawing.Point(8, 160);
			this.txtConsulta.Name = "txtConsulta";
			this.txtConsulta.Size = new System.Drawing.Size(208, 21);
			this.txtConsulta.TabIndex = 0;
			this.txtConsulta.TextChanged += new System.EventHandler(this.txtConsulta_TextChanged);
			this.txtConsulta.ValueChanged += new System.EventHandler(this.txtConsulta_ValueChanged);
			// 
			// txtGrupo
			// 
			this.txtGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGrupo.Location = new System.Drawing.Point(224, 160);
			this.txtGrupo.MaskInput = "nn";
			this.txtGrupo.MaxValue = 36;
			this.txtGrupo.MinValue = 1;
			this.txtGrupo.Name = "txtGrupo";
			this.txtGrupo.PromptChar = ' ';
			this.txtGrupo.Size = new System.Drawing.Size(88, 21);
			this.txtGrupo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtGrupo.TabIndex = 1;
			this.txtGrupo.ValueChanged += new System.EventHandler(this.txtGrupo_ValueChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 144);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(376, 8);
			this.groupBox1.TabIndex = 29;
			this.groupBox1.TabStop = false;
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
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.FillOnRequest = false;
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// txtIdUsuario
			// 
			this.txtIdUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdUsuario.Enabled = false;
			this.txtIdUsuario.Location = new System.Drawing.Point(288, 8);
			this.txtIdUsuario.MaskInput = "nn";
			this.txtIdUsuario.MinValue = 0;
			this.txtIdUsuario.Name = "txtIdUsuario";
			this.txtIdUsuario.PromptChar = ' ';
			this.txtIdUsuario.Size = new System.Drawing.Size(80, 21);
			this.txtIdUsuario.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtIdUsuario.TabIndex = 78;
			this.txtIdUsuario.Visible = false;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance13.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance13;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 70;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 22;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 63;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Bodega";
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 58;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Width = 26;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Cr";
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Width = 31;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "Almacen";
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 152;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			ultraGridBand2.Header.Caption = "LOCALES";
			ultraGridBand2.HeaderVisible = true;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance14;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 10F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance16;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(424, 32);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(288, 490);
			this.ultraGrid2.TabIndex = 183;
			this.ultraGrid2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid2_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(bool);
			ultraDataColumn12.DataType = typeof(bool);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// chkTodos
			// 
			this.chkTodos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTodos.Appearance = appearance20;
			this.chkTodos.Enabled = false;
			this.chkTodos.Location = new System.Drawing.Point(488, 0);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(176, 22);
			this.chkTodos.TabIndex = 187;
			this.chkTodos.Text = "Seleccionar todos los Locales";
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// chkActivo
			// 
			this.chkActivo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkActivo.Appearance = appearance21;
			this.chkActivo.Enabled = false;
			this.chkActivo.Location = new System.Drawing.Point(304, 8);
			this.chkActivo.Name = "chkActivo";
			this.chkActivo.Size = new System.Drawing.Size(56, 22);
			this.chkActivo.TabIndex = 188;
			this.chkActivo.Text = "Activo";
			// 
			// frmUsuarios
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(714, 528);
			this.Controls.Add(this.chkActivo);
			this.Controls.Add(this.chkTodos);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.txtIdUsuario);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtGrupo);
			this.Controls.Add(this.txtConsulta);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.btEditar);
			this.Controls.Add(this.btNuevo);
			this.Controls.Add(this.spnGrupo);
			this.Controls.Add(this.txtClave);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmUsuarios";
			this.Text = "Usuarios";
			this.Load += new System.EventHandler(this.frmUsuarios_Load);
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			
			for (int i = 35; i <= 119; i++)
			{
				this.txtNombre.Text = Funcion.sEscalarSQL(cdsSeguridad, string.Format("Select Nombre From Usuario Where idUsuario = {0}", i));
				string stSelect = string.Format("SELECT Count(*) FROM Master.Dbo.syslogins Where LoginName = '{0}'", this.txtNombre.Text.ToString());
				int iCuentaBases = Funcion.iEscalarSQL(cdsSeguridad, stSelect, true);
				if (iCuentaBases == 0)
				{
					string sSQL = string.Format("Create Login {0} With Password = '{1}'; Exec sp_addsrvrolemember '{0}', 'sysadmin'; Create User {0} For Login {0}", this.txtNombre.Text.ToString(), "Bl45o6$9");
					Funcion.EjecutaSQL(cdsSeguridad, sSQL);
				}	
			}

			Close();
		}

		private void btNuevo_Click(object sender, System.EventArgs e)
		{
			this.spnGrupo.Enabled = true;
			this.txtNombre.Enabled = true;
			this.txtClave.Enabled = true;
			this.txtIdUsuario.Value = 0;
			this.spnGrupo.Value = 1;
			this.txtNombre.Text = "";
			this.txtClave.Text = "";
			this.btEditar.Enabled = false;
			this.btGrabar.Enabled = true;
			this.btCancelar.Enabled = true;
			this.ultraGrid1.Enabled = false;

			this.chkActivo.Checked = true;
			
			bNuevo = true;
			bEdicion = true;
			
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaBodegasUsuarios 0, 1");

			ModificaGrilla(true);
			this.chkTodos.Enabled = true;
			this.chkTodos.Checked = false;
			this.spnGrupo.Focus();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.spnGrupo.Enabled = false;
			this.txtNombre.Enabled = false;
			this.txtClave.Enabled = false;
			this.txtIdUsuario.Value = 0;
			this.spnGrupo.Value = DBNull.Value;
			this.txtNombre.Text = "";
			this.txtClave.Text = "";
			this.btEditar.Enabled = false;
			this.btGrabar.Enabled = false;
			this.btCancelar.Enabled = false;

			this.ultraGrid1.Enabled = true;
			
			ModificaGrilla(false);
			this.chkTodos.Enabled = false;
			this.chkTodos.Checked = false;
			this.chkActivo.Enabled = false;
			this.chkActivo.Checked = false;

			bNuevo = false;
			bModifica = false;
			bEdicion = false;
		}

		private void btEditar_Click(object sender, System.EventArgs e)
		{
			this.spnGrupo.Enabled = true;
			this.txtNombre.Enabled = true;
			this.txtClave.Enabled = true;

			this.btGrabar.Enabled = true;
			this.btEditar.Enabled = false;
			this.ultraGrid1.Enabled = false;

      bEdicion = true;
			bModifica = true;

			ModificaGrilla(true);
			this.chkTodos.Enabled = true;
			this.chkTodos.Checked = false;
			
			this.chkActivo.Enabled = true;

			this.spnGrupo.Focus();
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.spnGrupo.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese el número de grupo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.spnGrupo.Focus();
				return;
			}

			if (this.txtNombre.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el nombre de usuario.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtNombre.Focus();
				return;
			}

			if (bNuevo)
			{
				string sSQL = string.Format("Select Count(*) From Usuario Where Nombre = '{0}'", this.txtNombre.Text.ToString());
				int iContUsuario = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iContUsuario > 0)
				{
					MessageBox.Show(string.Format("El nombre de usuario '{0}' ya existe.", this.txtNombre.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtNombre.Focus();
					return;
				}
			}

			if (bModifica)
			{
				string sSQL = string.Format("Select Count(*) From Usuario Where Nombre = '{0}' And idUsuario <> {1}", this.txtNombre.Text.ToString(), (int)this.txtIdUsuario.Value);
				int iContUsuario = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iContUsuario > 0)
				{
					MessageBox.Show(string.Format("El nombre de usuario '{0}' ya existe.", this.txtNombre.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtNombre.Focus();
					return;
				}
			}

			if (this.txtClave.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese la clave de usuario.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtClave.Focus();
				return;
			}
			#endregion Validacion

			#region Crea usuario en la base de datos
			string stExec = string.Format("Exec UsuarioCrear {0}, '{1}', '{2}', {3}, {4}, {5}", 
				(int)this.txtIdUsuario.Value, this.txtNombre.Text.ToString(), this.txtClave.Text.ToString(), (int)this.spnGrupo.Value, (bool)this.chkActivo.Checked, MenuLatinium.stIdDB);
			this.txtIdUsuario.Value = Funcion.iEscalarSQL(cdsSeteo, stExec);
			#endregion Crea usuario en la base de datos

			#region Crea usuario en el servidor
			string stSelect = string.Format("SELECT Count(*) FROM Master.Dbo.syslogins Where LoginName = '{0}'", this.txtNombre.Text.ToString());
			int iCuentaBases = Funcion.iEscalarSQL(cdsSeguridad, stSelect, true);
			if (iCuentaBases == 0)
			{
				string sSQL = string.Format("Create Login {0} With Password = '{1}'; Exec sp_addsrvrolemember '{0}', 'sysadmin'; Create User {0} For Login {0}", this.txtNombre.Text.ToString(), "Bl45o6$9");
				Funcion.EjecutaSQL(cdsSeguridad, sSQL);
			}
			#endregion Crea usuario en el servidor			

			#region Mensaje de confirmación
			string sMensaje = "";
			if (bNuevo) sMensaje = "Usuario Registrado.";
			if (bModifica) sMensaje = "Usuario Modificado.";
			MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			#endregion Mensaje de confirmación

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				string sSQL = string.Format("Exec GuardaUsuarioBodega {0}, {1}, {2}, {3}", (int)dr.Cells["idUsuarioBodega"].Value, (int)this.txtIdUsuario.Value, (int)dr.Cells["Bodega"].Value, (bool)dr.Cells["Sel"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);		
			}

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec UsuarioConsulta {0}, ''", (int)this.spnGrupo.Value));

			ModificaGrilla(false);
			this.chkTodos.Enabled = false;
			this.chkTodos.Checked = false;
			this.chkActivo.Enabled = false;
			this.chkActivo.Checked = false;

			this.spnGrupo.Enabled = false;
			this.txtNombre.Enabled = false;
			this.txtClave.Enabled = false;
			
			this.btGrabar.Enabled = false;
			this.btEditar.Enabled = true;
			this.chkActivo.Enabled = false;

			this.ultraGrid1.Enabled = true;

			bNuevo = false;
			bEdicion = false;
			bModifica = false;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void spnGrupo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNombre.Focus();
		}

		private void txtNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtClave.Focus();
		}

		private void txtClave_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btGrabar.Focus();
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void frmUsuarios_Load(object sender, System.EventArgs e)
		{
			ModificaGrilla(false);
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec UsuarioConsulta 0, ''");
		}

		private void txtGrupo_ValueChanged(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec UsuarioConsulta {0}, '{1}'", (int)this.txtGrupo.Value, this.txtConsulta.Text.ToString()));
		}

		private void txtConsulta_TextChanged(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec UsuarioConsulta {0}, '{1}'", (int)this.txtGrupo.Value, this.txtConsulta.Text.ToString()));
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void ultraGrid1_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.ActiveRow.Cells["idUsuario"].Value == DBNull.Value) return;

			this.txtIdUsuario.Value = (int)this.ultraGrid1.ActiveRow.Cells["idUsuario"].Value;
			this.spnGrupo.Value = (int)this.ultraGrid1.ActiveRow.Cells["idGrupo"].Value;
			this.txtNombre.Text = this.ultraGrid1.ActiveRow.Cells["Nombre"].Value.ToString();
			this.txtClave.Text = this.ultraGrid1.ActiveRow.Cells["Clave"].Value.ToString();
			this.chkActivo.Checked = (bool)this.ultraGrid1.ActiveRow.Cells["Activo"].Value;

			this.btCancelar.Enabled = true;

			this.btEditar.Enabled = true;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegasUsuarios {0}, 1", (int)this.ultraGrid1.ActiveRow.Cells["idUsuario"].Value));
		}

		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if (this.chkTodos.Checked) dr.Cells["Sel"].Value = 1;
				else dr.Cells["Sel"].Value = 0;
			}
		}

		private void txtConsulta_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void spnGrupo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
