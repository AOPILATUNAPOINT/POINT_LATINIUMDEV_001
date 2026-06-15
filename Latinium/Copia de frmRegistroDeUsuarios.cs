using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmRegistroDeUsuarios.
	/// </summary>
	public class frmRegistroDeUsuarios : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btSalir;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private Infragistics.Win.Misc.UltraButton btEditar;
		private Infragistics.Win.Misc.UltraButton btNuevo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtClave;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGrupo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConsulta;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkActivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnGrupo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdUsuario;
		private C1.Data.C1DataSet cdsSeguridad;
		private C1.Data.C1DataSet cdsSeteo;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		bool bNuevo = true;
		bool bEdicion = true;
		public frmRegistroDeUsuarios()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresaUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEmpresaUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUsuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Clave");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Activo");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			this.label19 = new System.Windows.Forms.Label();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.btEditar = new Infragistics.Win.Misc.UltraButton();
			this.btNuevo = new Infragistics.Win.Misc.UltraButton();
			this.txtClave = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtGrupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConsulta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.chkActivo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.spnGrupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdUsuario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.cdsSeteo = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(8, 11);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(35, 16);
			this.label19.TabIndex = 645;
			this.label19.Text = "Grupo";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btCancelar
			// 
			appearance1.Image = 9;
			this.btCancelar.Appearance = appearance1;
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Enabled = false;
			this.btCancelar.Location = new System.Drawing.Point(232, 120);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(64, 21);
			this.btCancelar.TabIndex = 650;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btSalir
			// 
			appearance2.Image = 9;
			this.btSalir.Appearance = appearance2;
			this.btSalir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSalir.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btSalir.Location = new System.Drawing.Point(304, 120);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(64, 21);
			this.btSalir.TabIndex = 652;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// btGrabar
			// 
			appearance3.Image = 9;
			this.btGrabar.Appearance = appearance3;
			this.btGrabar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGrabar.Enabled = false;
			this.btGrabar.Location = new System.Drawing.Point(160, 120);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(64, 21);
			this.btGrabar.TabIndex = 649;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btEditar
			// 
			appearance4.Image = 9;
			this.btEditar.Appearance = appearance4;
			this.btEditar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btEditar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btEditar.Enabled = false;
			this.btEditar.Location = new System.Drawing.Point(88, 120);
			this.btEditar.Name = "btEditar";
			this.btEditar.Size = new System.Drawing.Size(64, 21);
			this.btEditar.TabIndex = 651;
			this.btEditar.Text = "&Editar";
			this.btEditar.Click += new System.EventHandler(this.btEditar_Click);
			// 
			// btNuevo
			// 
			appearance5.Image = 9;
			this.btNuevo.Appearance = appearance5;
			this.btNuevo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNuevo.Location = new System.Drawing.Point(16, 120);
			this.btNuevo.Name = "btNuevo";
			this.btNuevo.Size = new System.Drawing.Size(64, 21);
			this.btNuevo.TabIndex = 646;
			this.btNuevo.Text = "&Nuevo";
			this.btNuevo.Click += new System.EventHandler(this.btNuevo_Click);
			// 
			// txtClave
			// 
			this.txtClave.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtClave.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClave.Enabled = false;
			this.txtClave.Location = new System.Drawing.Point(96, 72);
			this.txtClave.Name = "txtClave";
			this.txtClave.PasswordChar = '*';
			this.txtClave.Size = new System.Drawing.Size(248, 22);
			this.txtClave.TabIndex = 648;
			this.txtClave.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtClave_KeyPress);
			// 
			// txtNombre
			// 
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(96, 40);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(248, 22);
			this.txtNombre.TabIndex = 647;
			this.txtNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombre_KeyPress);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 75);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(61, 16);
			this.label2.TabIndex = 654;
			this.label2.Text = "Contraseña";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 653;
			this.label1.Text = "Nombre";
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(8, 152);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(672, 8);
			this.groupBox1.TabIndex = 658;
			this.groupBox1.TabStop = false;
			// 
			// txtGrupo
			// 
			this.txtGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGrupo.Location = new System.Drawing.Point(232, 168);
			this.txtGrupo.MaskInput = "nn";
			this.txtGrupo.MaxValue = 20;
			this.txtGrupo.MinValue = 1;
			this.txtGrupo.Name = "txtGrupo";
			this.txtGrupo.PromptChar = ' ';
			this.txtGrupo.Size = new System.Drawing.Size(88, 22);
			this.txtGrupo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtGrupo.TabIndex = 656;
			// 
			// txtConsulta
			// 
			this.txtConsulta.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConsulta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConsulta.Location = new System.Drawing.Point(16, 168);
			this.txtConsulta.Name = "txtConsulta";
			this.txtConsulta.Size = new System.Drawing.Size(208, 22);
			this.txtConsulta.TabIndex = 655;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance6;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 87;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 56;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 242;
			ultraGridColumn4.Header.Caption = "...";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 25;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance7;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 10F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance9;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(392, 200);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(288, 240);
			this.ultraGrid2.TabIndex = 659;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(bool);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance13.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance13;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "Usuario";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.Caption = "Grupo";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 30;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 230;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 81;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.Caption = "Ingreso";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Width = 62;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "A";
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Width = 25;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand2.Header.Caption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance14;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance15;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 10F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 200);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(368, 240);
			this.ultraGrid1.TabIndex = 660;
			this.ultraGrid1.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ultraGrid1_BeforeSelectChange);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			this.ultraGrid1.Click += new System.EventHandler(this.ultraGrid1_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn9.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn10.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
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
			// chkActivo
			// 
			this.chkActivo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkActivo.Appearance = appearance20;
			this.chkActivo.Enabled = false;
			this.chkActivo.Location = new System.Drawing.Point(368, 8);
			this.chkActivo.Name = "chkActivo";
			this.chkActivo.Size = new System.Drawing.Size(56, 22);
			this.chkActivo.TabIndex = 662;
			this.chkActivo.Text = "Activo";
			// 
			// spnGrupo
			// 
			this.spnGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnGrupo.Enabled = false;
			this.spnGrupo.Location = new System.Drawing.Point(96, 8);
			this.spnGrupo.MaskInput = "nn";
			this.spnGrupo.MaxValue = 20;
			this.spnGrupo.MinValue = 1;
			this.spnGrupo.Name = "spnGrupo";
			this.spnGrupo.PromptChar = ' ';
			this.spnGrupo.Size = new System.Drawing.Size(80, 22);
			this.spnGrupo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnGrupo.TabIndex = 661;
			this.spnGrupo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.spnGrupo_KeyPress);
			// 
			// txtIdUsuario
			// 
			this.txtIdUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdUsuario.Enabled = false;
			this.txtIdUsuario.Location = new System.Drawing.Point(600, 8);
			this.txtIdUsuario.MaskInput = "nn";
			this.txtIdUsuario.MinValue = 0;
			this.txtIdUsuario.Name = "txtIdUsuario";
			this.txtIdUsuario.PromptChar = ' ';
			this.txtIdUsuario.Size = new System.Drawing.Size(80, 22);
			this.txtIdUsuario.TabIndex = 663;
			this.txtIdUsuario.Visible = false;
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
			// frmRegistroDeUsuarios
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(688, 454);
			this.Controls.Add(this.txtIdUsuario);
			this.Controls.Add(this.chkActivo);
			this.Controls.Add(this.spnGrupo);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtGrupo);
			this.Controls.Add(this.txtConsulta);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.btEditar);
			this.Controls.Add(this.btNuevo);
			this.Controls.Add(this.txtClave);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "frmRegistroDeUsuarios";
			this.Text = "Registro De Usuarios";
			this.Load += new System.EventHandler(this.frmRegistroDeUsuarios_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtClave)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmRegistroDeUsuarios_Load(object sender, System.EventArgs e)
		{
		
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

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;			
			bEdicion = false;

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

			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaEmpresasPorUsuario 0");

			this.ultraGrid1.Enabled = true;
			
			ModificaGrilla(false);						
		}

		private void btEditar_Click(object sender, System.EventArgs e)
		{
			this.spnGrupo.Enabled = true;
			this.txtNombre.Enabled = true;
			this.txtClave.Enabled = true;

			this.btGrabar.Enabled = true;
			this.btEditar.Enabled = false;
			this.ultraGrid1.Enabled = false;

			ModificaGrilla(true);
			
			this.chkActivo.Enabled = true;

			this.spnGrupo.Focus();

			bEdicion = true;			
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
						
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaEmpresasPorUsuario 0");

			bNuevo = true;
			bEdicion = true;
			
			ModificaGrilla(true);
			
			this.spnGrupo.Focus();
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

			if (!bNuevo)
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
			string stExec = string.Format("Exec UsuarioCrear {0}, '{1}', '{2}', {3}, {4}", (int)this.txtIdUsuario.Value, this.txtNombre.Text.ToString(), this.txtClave.Text.ToString(), (int)this.spnGrupo.Value, (bool)this.chkActivo.Checked);
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
			if (!bNuevo) sMensaje = "Usuario Modificado.";
			MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			#endregion Mensaje de confirmación

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				string sSQL = string.Format("Exec GuardaEmpresasUsuarios {0}, {1}, {2}, {3}", (int)dr.Cells["idEmpresaUsuario"].Value, (int)this.txtIdUsuario.Value, (int)dr.Cells["idEmpresa"].Value, (bool)dr.Cells["Sel"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);		
			}

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec UsuarioConsulta {0}, ''", (int)this.spnGrupo.Value));

			ModificaGrilla(false);
			
			this.spnGrupo.Enabled = false;
			this.txtNombre.Enabled = false;
			this.txtClave.Enabled = false;
			
			this.btGrabar.Enabled = false;
			this.btEditar.Enabled = true;

			this.ultraGrid1.Enabled = true;

			bNuevo = false;
			bEdicion = false;			
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

		private void ultraGrid1_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaEmpresasPorUsuario {0}", (int)this.ultraGrid1.ActiveRow.Cells["idUsuario"].Value));
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}

