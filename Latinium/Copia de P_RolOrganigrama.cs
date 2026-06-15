using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Summary description for P_RolOrganigrama.
	/// </summary>
	public class P_RolOrganigrama : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataView cdvSucursal;
		private C1.Data.C1DataSet cdsSucursal;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvCargo;
		private DevExpress.XtraEditors.SplitContainerControl splitContainerControl1;
		private DevExpress.XtraTreeList.TreeList treeList1;
		private DevExpress.XtraEditors.SimpleButton btContraer;
		private DevExpress.XtraEditors.SimpleButton btExpandir;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colNombre;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colidCargo;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSucursal;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraPictureBox picFoto;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public P_RolOrganigrama()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolDepartamento", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolSucursal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OrigenFondos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSeguro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipoEmpleador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			this.cdvSucursal = new C1.Data.C1DataView();
			this.cdsSucursal = new C1.Data.C1DataSet();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdvCargo = new C1.Data.C1DataView();
			this.splitContainerControl1 = new DevExpress.XtraEditors.SplitContainerControl();
			this.btContraer = new DevExpress.XtraEditors.SimpleButton();
			this.btExpandir = new DevExpress.XtraEditors.SimpleButton();
			this.treeList1 = new DevExpress.XtraTreeList.TreeList();
			this.colNombre = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.colidCargo = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label13 = new System.Windows.Forms.Label();
			this.cmbSucursal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label12 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.picFoto = new Infragistics.Win.UltraWinEditors.UltraPictureBox();
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).BeginInit();
			this.splitContainerControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.treeList1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			this.SuspendLayout();
			// 
			// cdvSucursal
			// 
			this.cdvSucursal.BindingContextCtrl = this;
			this.cdvSucursal.DataSet = this.cdsSucursal;
			this.cdvSucursal.Sort = "Sucursal";
			this.cdvSucursal.TableName = "";
			this.cdvSucursal.TableViewName = "RolSucursal";
			// 
			// cdsSucursal
			// 
			this.cdsSucursal.BindingContextCtrl = this;
			this.cdsSucursal.DataLibrary = "LibPersonal";
			this.cdsSucursal.DataLibraryUrl = "";
			this.cdsSucursal.DataSetDef = "dsRolTablas";
			this.cdsSucursal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSucursal.Name = "cdsSucursal";
			this.cdsSucursal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsSucursal.SchemaDef = null;
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			// 
			// cdvCargo
			// 
			this.cdvCargo.BindingContextCtrl = this;
			this.cdvCargo.DataSet = this.cdsSucursal;
			this.cdvCargo.TableName = "";
			this.cdvCargo.TableViewName = "RolCargo";
			// 
			// splitContainerControl1
			// 
			this.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.splitContainerControl1.Location = new System.Drawing.Point(0, 0);
			this.splitContainerControl1.Name = "splitContainerControl1";
			this.splitContainerControl1.Panel1.Controls.Add(this.btContraer);
			this.splitContainerControl1.Panel1.Controls.Add(this.btExpandir);
			this.splitContainerControl1.Panel1.Controls.Add(this.treeList1);
			this.splitContainerControl1.Panel1.Text = "splitContainerControl1_Panel1";
			this.splitContainerControl1.Panel2.Controls.Add(this.cmbDepartamento);
			this.splitContainerControl1.Panel2.Controls.Add(this.label13);
			this.splitContainerControl1.Panel2.Controls.Add(this.cmbSucursal);
			this.splitContainerControl1.Panel2.Controls.Add(this.label12);
			this.splitContainerControl1.Panel2.Controls.Add(this.txtNombre);
			this.splitContainerControl1.Panel2.Controls.Add(this.label1);
			this.splitContainerControl1.Panel2.Controls.Add(this.picFoto);
			this.splitContainerControl1.Panel2.Text = "splitContainerControl1_Panel2";
			this.splitContainerControl1.Size = new System.Drawing.Size(888, 390);
			this.splitContainerControl1.SplitterPosition = 597;
			this.splitContainerControl1.TabIndex = 0;
			this.splitContainerControl1.Text = "splitContainerControl1";
			// 
			// btContraer
			// 
			this.btContraer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btContraer.ImageIndex = 2;
			this.btContraer.Location = new System.Drawing.Point(344, 360);
			this.btContraer.Name = "btContraer";
			this.btContraer.Size = new System.Drawing.Size(160, 25);
			this.btContraer.TabIndex = 10;
			this.btContraer.Text = "-";
			this.btContraer.ToolTip = "Contraer plan de cuentas";
			this.btContraer.Click += new System.EventHandler(this.btContraer_Click);
			// 
			// btExpandir
			// 
			this.btExpandir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btExpandir.ImageIndex = 3;
			this.btExpandir.Location = new System.Drawing.Point(96, 360);
			this.btExpandir.Name = "btExpandir";
			this.btExpandir.Size = new System.Drawing.Size(168, 25);
			this.btExpandir.TabIndex = 9;
			this.btExpandir.Text = "+";
			this.btExpandir.ToolTip = "Expandir plan de cuentas";
			this.btExpandir.Click += new System.EventHandler(this.btExpandir_Click);
			// 
			// treeList1
			// 
			this.treeList1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.treeList1.Appearance.Empty.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(232)), ((System.Byte)(228)), ((System.Byte)(220)));
			this.treeList1.Appearance.Empty.Options.UseBackColor = true;
			this.treeList1.Appearance.EvenRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.treeList1.Appearance.EvenRow.BackColor2 = System.Drawing.Color.GhostWhite;
			this.treeList1.Appearance.EvenRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.treeList1.Appearance.EvenRow.ForeColor = System.Drawing.Color.Black;
			this.treeList1.Appearance.EvenRow.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal;
			this.treeList1.Appearance.EvenRow.Options.UseBackColor = true;
			this.treeList1.Appearance.EvenRow.Options.UseFont = true;
			this.treeList1.Appearance.EvenRow.Options.UseForeColor = true;
			this.treeList1.Appearance.FocusedCell.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(225)));
			this.treeList1.Appearance.FocusedCell.Font = new System.Drawing.Font("Tahoma", 10F);
			this.treeList1.Appearance.FocusedCell.ForeColor = System.Drawing.Color.Black;
			this.treeList1.Appearance.FocusedCell.Options.UseBackColor = true;
			this.treeList1.Appearance.FocusedCell.Options.UseFont = true;
			this.treeList1.Appearance.FocusedCell.Options.UseForeColor = true;
			this.treeList1.Appearance.FocusedRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(36)), ((System.Byte)(106)));
			this.treeList1.Appearance.FocusedRow.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(60)), ((System.Byte)(86)), ((System.Byte)(156)));
			this.treeList1.Appearance.FocusedRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.treeList1.Appearance.FocusedRow.ForeColor = System.Drawing.Color.White;
			this.treeList1.Appearance.FocusedRow.Options.UseBackColor = true;
			this.treeList1.Appearance.FocusedRow.Options.UseFont = true;
			this.treeList1.Appearance.FocusedRow.Options.UseForeColor = true;
			this.treeList1.Appearance.FooterPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.treeList1.Appearance.FooterPanel.BorderColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.treeList1.Appearance.FooterPanel.Font = new System.Drawing.Font("Tahoma", 10F);
			this.treeList1.Appearance.FooterPanel.ForeColor = System.Drawing.Color.Black;
			this.treeList1.Appearance.FooterPanel.Options.UseBackColor = true;
			this.treeList1.Appearance.FooterPanel.Options.UseBorderColor = true;
			this.treeList1.Appearance.FooterPanel.Options.UseFont = true;
			this.treeList1.Appearance.FooterPanel.Options.UseForeColor = true;
			this.treeList1.Appearance.GroupButton.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.treeList1.Appearance.GroupButton.BorderColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.treeList1.Appearance.GroupButton.ForeColor = System.Drawing.Color.Black;
			this.treeList1.Appearance.GroupButton.Options.UseBackColor = true;
			this.treeList1.Appearance.GroupButton.Options.UseBorderColor = true;
			this.treeList1.Appearance.GroupButton.Options.UseForeColor = true;
			this.treeList1.Appearance.GroupFooter.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(222)), ((System.Byte)(218)), ((System.Byte)(210)));
			this.treeList1.Appearance.GroupFooter.BorderColor = System.Drawing.Color.FromArgb(((System.Byte)(222)), ((System.Byte)(218)), ((System.Byte)(210)));
			this.treeList1.Appearance.GroupFooter.ForeColor = System.Drawing.Color.Black;
			this.treeList1.Appearance.GroupFooter.Options.UseBackColor = true;
			this.treeList1.Appearance.GroupFooter.Options.UseBorderColor = true;
			this.treeList1.Appearance.GroupFooter.Options.UseForeColor = true;
			this.treeList1.Appearance.HeaderPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.treeList1.Appearance.HeaderPanel.BorderColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.treeList1.Appearance.HeaderPanel.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.treeList1.Appearance.HeaderPanel.ForeColor = System.Drawing.Color.Black;
			this.treeList1.Appearance.HeaderPanel.Options.UseBackColor = true;
			this.treeList1.Appearance.HeaderPanel.Options.UseBorderColor = true;
			this.treeList1.Appearance.HeaderPanel.Options.UseFont = true;
			this.treeList1.Appearance.HeaderPanel.Options.UseForeColor = true;
			this.treeList1.Appearance.HideSelectionRow.BackColor = System.Drawing.Color.Gray;
			this.treeList1.Appearance.HideSelectionRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.treeList1.Appearance.HideSelectionRow.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.treeList1.Appearance.HideSelectionRow.Options.UseBackColor = true;
			this.treeList1.Appearance.HideSelectionRow.Options.UseFont = true;
			this.treeList1.Appearance.HideSelectionRow.Options.UseForeColor = true;
			this.treeList1.Appearance.HorzLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.treeList1.Appearance.HorzLine.Options.UseBackColor = true;
			this.treeList1.Appearance.OddRow.BackColor = System.Drawing.Color.White;
			this.treeList1.Appearance.OddRow.BackColor2 = System.Drawing.Color.White;
			this.treeList1.Appearance.OddRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.treeList1.Appearance.OddRow.ForeColor = System.Drawing.Color.Black;
			this.treeList1.Appearance.OddRow.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal;
			this.treeList1.Appearance.OddRow.Options.UseBackColor = true;
			this.treeList1.Appearance.OddRow.Options.UseFont = true;
			this.treeList1.Appearance.OddRow.Options.UseForeColor = true;
			this.treeList1.Appearance.Preview.BackColor = System.Drawing.Color.White;
			this.treeList1.Appearance.Preview.Font = new System.Drawing.Font("Tahoma", 10F);
			this.treeList1.Appearance.Preview.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(36)), ((System.Byte)(106)));
			this.treeList1.Appearance.Preview.Options.UseBackColor = true;
			this.treeList1.Appearance.Preview.Options.UseFont = true;
			this.treeList1.Appearance.Preview.Options.UseForeColor = true;
			this.treeList1.Appearance.Row.BackColor = System.Drawing.Color.White;
			this.treeList1.Appearance.Row.Font = new System.Drawing.Font("Tahoma", 10F);
			this.treeList1.Appearance.Row.ForeColor = System.Drawing.Color.Black;
			this.treeList1.Appearance.Row.Options.UseBackColor = true;
			this.treeList1.Appearance.Row.Options.UseFont = true;
			this.treeList1.Appearance.Row.Options.UseForeColor = true;
			this.treeList1.Appearance.SelectedRow.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(20)), ((System.Byte)(46)), ((System.Byte)(116)));
			this.treeList1.Appearance.SelectedRow.Font = new System.Drawing.Font("Tahoma", 10F);
			this.treeList1.Appearance.SelectedRow.ForeColor = System.Drawing.Color.White;
			this.treeList1.Appearance.SelectedRow.Options.UseBackColor = true;
			this.treeList1.Appearance.SelectedRow.Options.UseFont = true;
			this.treeList1.Appearance.SelectedRow.Options.UseForeColor = true;
			this.treeList1.Appearance.TreeLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(58)), ((System.Byte)(58)), ((System.Byte)(58)));
			this.treeList1.Appearance.TreeLine.Options.UseBackColor = true;
			this.treeList1.Appearance.VertLine.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)));
			this.treeList1.Appearance.VertLine.Options.UseBackColor = true;
			this.treeList1.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
																																														 this.colNombre,
																																														 this.colidCargo});
			this.treeList1.DataMember = "Personal";
			this.treeList1.DataSource = this.cdsPersonal;
			this.treeList1.KeyFieldName = "idPersonal";
			this.treeList1.Location = new System.Drawing.Point(0, 0);
			this.treeList1.Name = "treeList1";
			this.treeList1.OptionsView.EnableAppearanceEvenRow = true;
			this.treeList1.OptionsView.EnableAppearanceOddRow = true;
			this.treeList1.OptionsView.ShowIndicator = false;
			this.treeList1.ParentFieldName = "idJefe";
			this.treeList1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																			 this.repositoryItemLookUpEdit1});
			this.treeList1.Size = new System.Drawing.Size(596, 360);
			this.treeList1.TabIndex = 0;
			// 
			// colNombre
			// 
			this.colNombre.Caption = "Nombre";
			this.colNombre.FieldName = "Nombre";
			this.colNombre.Name = "colNombre";
			this.colNombre.OptionsColumn.ReadOnly = true;
			this.colNombre.VisibleIndex = 1;
			this.colNombre.Width = 288;
			// 
			// colidCargo
			// 
			this.colidCargo.Caption = "Cargo";
			this.colidCargo.ColumnEdit = this.repositoryItemLookUpEdit1;
			this.colidCargo.FieldName = "idCargo";
			this.colidCargo.Name = "colidCargo";
			this.colidCargo.OptionsColumn.ReadOnly = true;
			this.colidCargo.VisibleIndex = 0;
			this.colidCargo.Width = 287;
			// 
			// repositoryItemLookUpEdit1
			// 
			this.repositoryItemLookUpEdit1.AutoHeight = false;
			this.repositoryItemLookUpEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit1.DataSource = this.cdvCargo;
			this.repositoryItemLookUpEdit1.DisplayMember = "Descripcion";
			this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
			this.repositoryItemLookUpEdit1.ValueMember = "idCargo";
			// 
			// cmbDepartamento
			// 
			appearance1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance1.BackColor2 = System.Drawing.Color.White;
			appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbDepartamento.Appearance = appearance1;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idDepartamento"));
			this.cmbDepartamento.DataMember = "RolDepartamento";
			this.cmbDepartamento.DataSource = this.cdsSucursal;
			ultraGridBand1.AddButtonCaption = "RolSucursal";
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Width = 111;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 271;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(64, 296);
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.ReadOnly = true;
			this.cmbDepartamento.Size = new System.Drawing.Size(208, 21);
			this.cmbDepartamento.TabIndex = 30;
			this.cmbDepartamento.ValueMember = "idDepartamento";
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(8, 296);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(38, 16);
			this.label13.TabIndex = 29;
			this.label13.Text = "Depto:";
			// 
			// cmbSucursal
			// 
			appearance2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance2.BackColor2 = System.Drawing.Color.White;
			appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbSucursal.Appearance = appearance2;
			this.cmbSucursal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSucursal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSucursal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idRolSucursal"));
			this.cmbSucursal.DataSource = this.cdvSucursal;
			ultraGridColumn4.Header.VisiblePosition = 8;
			ultraGridColumn4.Width = 101;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 234;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 66;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbSucursal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSucursal.DisplayMember = "Sucursal";
			this.cmbSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSucursal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSucursal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSucursal.Location = new System.Drawing.Point(64, 264);
			this.cmbSucursal.Name = "cmbSucursal";
			this.cmbSucursal.ReadOnly = true;
			this.cmbSucursal.Size = new System.Drawing.Size(208, 21);
			this.cmbSucursal.TabIndex = 28;
			this.cmbSucursal.ValueMember = "idRolSucursal";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(8, 264);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(50, 16);
			this.label12.TabIndex = 27;
			this.label12.Text = "Sucursal:";
			// 
			// txtNombre
			// 
			appearance3.BackColor = System.Drawing.Color.Gainsboro;
			appearance3.BackColor2 = System.Drawing.Color.White;
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNombre.Appearance = appearance3;
			this.txtNombre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.Nombre"));
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(64, 232);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.ReadOnly = true;
			this.txtNombre.Size = new System.Drawing.Size(208, 20);
			this.txtNombre.TabIndex = 26;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 240);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 25;
			this.label1.Text = "Nombre:";
			// 
			// picFoto
			// 
			this.picFoto.BorderShadowColor = System.Drawing.Color.Empty;
			this.picFoto.BorderStyle = Infragistics.Win.UIElementBorderStyle.Inset;
			this.picFoto.DataBindings.Add(new System.Windows.Forms.Binding("Image", this.cdsPersonal, "Personal.Fotos"));
			this.picFoto.Location = new System.Drawing.Point(64, 72);
			this.picFoto.Name = "picFoto";
			this.picFoto.ScaleImage = Infragistics.Win.ScaleImage.Always;
			this.picFoto.Size = new System.Drawing.Size(200, 120);
			this.picFoto.TabIndex = 24;
			// 
			// P_RolOrganigrama
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(888, 390);
			this.Controls.Add(this.splitContainerControl1);
			this.Name = "P_RolOrganigrama";
			this.Text = "Organigrama";
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.splitContainerControl1)).EndInit();
			this.splitContainerControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.treeList1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btExpandir_Click(object sender, System.EventArgs e)
		{
			treeList1.ExpandAll();
		}

		private void btContraer_Click(object sender, System.EventArgs e)
		{
			treeList1.CollapseAll();
		}
	}
}

