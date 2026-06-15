using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolSueldoIncr.
	/// </summary>
	public class RolSueldoIncr : DevExpress.XtraEditors.XtraForm
	{
		public DevExpress.XtraEditors.SimpleButton btCalcular;
		public DevExpress.XtraEditors.SimpleButton btSalir;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSucursal;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDepartamento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCargo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGrupo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCentroCosto;
		private C1.Data.C1DataSet cdsRolConcepto;
		private C1.Data.C1DataSet cdsSucursal;
		private C1.Data.C1DataView cdvSucursal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCargo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSucursal;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvProyecto;
		private C1.Data.C1DataSet c1DataSet1;
		private C1.Data.C1DataView c1DataView1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroCosto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolSueldoIncr()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolCargo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolDepartamento", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolSucursal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OrigenFondos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSeguro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipoEmpleador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.btCalcular = new DevExpress.XtraEditors.SimpleButton();
			this.btSalir = new DevExpress.XtraEditors.SimpleButton();
			this.chkSucursal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDepartamento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCargo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkGrupo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCentroCosto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsRolConcepto = new C1.Data.C1DataSet();
			this.cdsSucursal = new C1.Data.C1DataSet();
			this.cdvSucursal = new C1.Data.C1DataView();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbSucursal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.c1DataSet1 = new C1.Data.C1DataSet();
			this.c1DataView1 = new C1.Data.C1DataView();
			this.cmbCentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataSet1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			this.SuspendLayout();
			// 
			// btCalcular
			// 
			this.btCalcular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCalcular.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btCalcular.Location = new System.Drawing.Point(88, 271);
			this.btCalcular.Name = "btCalcular";
			this.btCalcular.Size = new System.Drawing.Size(72, 25);
			this.btCalcular.TabIndex = 21;
			this.btCalcular.Text = "Calcular";
			// 
			// btSalir
			// 
			this.btSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btSalir.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.btSalir.Location = new System.Drawing.Point(272, 271);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(72, 25);
			this.btSalir.TabIndex = 22;
			this.btSalir.Text = "Salir";
			// 
			// chkSucursal
			// 
			this.chkSucursal.Checked = true;
			this.chkSucursal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkSucursal.Location = new System.Drawing.Point(32, 34);
			this.chkSucursal.Name = "chkSucursal";
			this.chkSucursal.Size = new System.Drawing.Size(96, 18);
			this.chkSucursal.TabIndex = 37;
			this.chkSucursal.Text = "Toda Sucursal";
			this.chkSucursal.CheckedChanged += new System.EventHandler(this.chkSucursal_CheckedChanged);
			// 
			// chkDepartamento
			// 
			this.chkDepartamento.Checked = true;
			this.chkDepartamento.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDepartamento.Location = new System.Drawing.Point(32, 69);
			this.chkDepartamento.Name = "chkDepartamento";
			this.chkDepartamento.Size = new System.Drawing.Size(128, 17);
			this.chkDepartamento.TabIndex = 38;
			this.chkDepartamento.Text = "Todo Departamento";
			this.chkDepartamento.CheckedChanged += new System.EventHandler(this.chkDepartamento_CheckedChanged);
			// 
			// chkCargo
			// 
			this.chkCargo.Checked = true;
			this.chkCargo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCargo.Location = new System.Drawing.Point(32, 103);
			this.chkCargo.Name = "chkCargo";
			this.chkCargo.Size = new System.Drawing.Size(128, 18);
			this.chkCargo.TabIndex = 39;
			this.chkCargo.Text = "Todo Cargo";
			// 
			// chkGrupo
			// 
			this.chkGrupo.Checked = true;
			this.chkGrupo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkGrupo.Location = new System.Drawing.Point(32, 138);
			this.chkGrupo.Name = "chkGrupo";
			this.chkGrupo.Size = new System.Drawing.Size(128, 17);
			this.chkGrupo.TabIndex = 40;
			this.chkGrupo.Text = "Todo Grupo";
			// 
			// chkCentroCosto
			// 
			this.chkCentroCosto.Checked = true;
			this.chkCentroCosto.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCentroCosto.Location = new System.Drawing.Point(32, 172);
			this.chkCentroCosto.Name = "chkCentroCosto";
			this.chkCentroCosto.Size = new System.Drawing.Size(128, 18);
			this.chkCentroCosto.TabIndex = 41;
			this.chkCentroCosto.Text = "Todo Centro Costo";
			// 
			// cdsRolConcepto
			// 
			this.cdsRolConcepto.BindingContextCtrl = this;
			this.cdsRolConcepto.DataLibrary = "LibPersonal";
			this.cdsRolConcepto.DataLibraryUrl = "";
			this.cdsRolConcepto.DataSetDef = "dsRolGrupo";
			this.cdsRolConcepto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolConcepto.Name = "cdsRolConcepto";
			this.cdsRolConcepto.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolConcepto.SchemaDef = null;
			this.cdsRolConcepto.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolConcepto_BeforeFill);
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
			// cdvSucursal
			// 
			this.cdvSucursal.BindingContextCtrl = this;
			this.cdvSucursal.DataSet = this.cdsSucursal;
			this.cdvSucursal.Sort = "Sucursal";
			this.cdvSucursal.TableName = "";
			this.cdvSucursal.TableViewName = "RolSucursal";
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataMember = "RolGrupo";
			this.cmbGrupo.DataSource = this.cdsRolConcepto;
			ultraGridBand1.AddButtonCaption = "RolTipo";
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Width = 105;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 296;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(184, 138);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(208, 21);
			this.cmbGrupo.TabIndex = 45;
			this.cmbGrupo.ValueMember = "idRolGrupo";
			// 
			// cmbCargo
			// 
			appearance1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance1.BackColor2 = System.Drawing.Color.White;
			appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCargo.Appearance = appearance1;
			this.cmbCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargo.DataMember = "RolCargo";
			this.cmbCargo.DataSource = this.cdsSucursal;
			ultraGridBand2.AddButtonCaption = "RolSucursal";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 125;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 215;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCargo.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCargo.DisplayMember = "Descripcion";
			this.cmbCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCargo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargo.Location = new System.Drawing.Point(184, 103);
			this.cmbCargo.Name = "cmbCargo";
			this.cmbCargo.Size = new System.Drawing.Size(208, 21);
			this.cmbCargo.TabIndex = 44;
			this.cmbCargo.ValueMember = "idCargo";
			// 
			// cmbDepartamento
			// 
			appearance2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance2.BackColor2 = System.Drawing.Color.White;
			appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbDepartamento.Appearance = appearance2;
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento.DataMember = "RolDepartamento";
			this.cmbDepartamento.DataSource = this.cdsSucursal;
			ultraGridBand3.AddButtonCaption = "RolSucursal";
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 111;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Width = 271;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(184, 69);
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(208, 21);
			this.cmbDepartamento.TabIndex = 43;
			this.cmbDepartamento.ValueMember = "idDepartamento";
			// 
			// cmbSucursal
			// 
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance3.BackColor2 = System.Drawing.Color.White;
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbSucursal.Appearance = appearance3;
			this.cmbSucursal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSucursal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSucursal.DataSource = this.cdvSucursal;
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Width = 101;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Width = 234;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 66;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbSucursal.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbSucursal.DisplayMember = "Sucursal";
			this.cmbSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSucursal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSucursal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSucursal.Location = new System.Drawing.Point(184, 34);
			this.cmbSucursal.Name = "cmbSucursal";
			this.cmbSucursal.Size = new System.Drawing.Size(208, 21);
			this.cmbSucursal.TabIndex = 42;
			this.cmbSucursal.ValueMember = "idRolSucursal";
			this.cmbSucursal.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbSucursal_InitializeLayout);
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// c1DataSet1
			// 
			this.c1DataSet1.BindingContextCtrl = this;
			this.c1DataSet1.DataLibrary = "LibContabilidad";
			this.c1DataSet1.DataLibraryUrl = "";
			this.c1DataSet1.DataSetDef = "dsProyecto";
			this.c1DataSet1.Locale = new System.Globalization.CultureInfo("es-EC");
			this.c1DataSet1.Name = "c1DataSet1";
			this.c1DataSet1.SchemaClassName = "LibContabilidad.DataClass";
			this.c1DataSet1.SchemaDef = null;
			// 
			// c1DataView1
			// 
			this.c1DataView1.BindingContextCtrl = this;
			this.c1DataView1.DataSet = this.c1DataSet1;
			this.c1DataView1.Sort = "Nombre";
			this.c1DataView1.TableName = "";
			this.c1DataView1.TableViewName = "Proyecto";
			// 
			// cmbCentroCosto
			// 
			this.cmbCentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroCosto.DataSource = this.cdvProyecto;
			ultraGridBand5.AddButtonCaption = "RolTipo";
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Width = 85;
			ultraGridColumn20.Header.VisiblePosition = 6;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 14;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Width = 311;
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 24;
			ultraGridColumn23.Header.VisiblePosition = 7;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 24;
			ultraGridColumn24.Header.VisiblePosition = 4;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 24;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 94;
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 22;
			ultraGridColumn27.Header.VisiblePosition = 8;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 17;
			ultraGridColumn28.Header.VisiblePosition = 9;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbCentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCentroCosto.DisplayMember = "Nombre";
			this.cmbCentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroCosto.Location = new System.Drawing.Point(184, 172);
			this.cmbCentroCosto.Name = "cmbCentroCosto";
			this.cmbCentroCosto.Size = new System.Drawing.Size(208, 21);
			this.cmbCentroCosto.TabIndex = 46;
			this.cmbCentroCosto.ValueMember = "idProyecto";
			// 
			// txtValor
			// 
			this.txtValor.Location = new System.Drawing.Point(312, 224);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(72, 22);
			this.txtValor.TabIndex = 47;
			this.txtValor.ValueChanged += new System.EventHandler(this.txtValor_ValueChanged);
			// 
			// optTipo
			// 
			this.optTipo.ItemAppearance = appearance4;
			this.optTipo.ItemOrigin = new System.Drawing.Point(4, 4);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Porcentaje";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Incremento";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.Location = new System.Drawing.Point(32, 224);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(176, 26);
			this.optTipo.TabIndex = 48;
			this.optTipo.ValueChanged += new System.EventHandler(this.optTipo_ValueChanged);
			// 
			// RolSueldoIncr
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(408, 324);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.cmbCentroCosto);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.cmbCargo);
			this.Controls.Add(this.cmbDepartamento);
			this.Controls.Add(this.cmbSucursal);
			this.Controls.Add(this.chkCentroCosto);
			this.Controls.Add(this.chkGrupo);
			this.Controls.Add(this.chkCargo);
			this.Controls.Add(this.chkDepartamento);
			this.Controls.Add(this.chkSucursal);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btCalcular);
			this.Name = "RolSueldoIncr";
			this.Text = "Incremento Sueldos";
			this.Load += new System.EventHandler(this.RolSueldoIncr_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataSet1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsRolConcepto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRolConcepto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void RolSueldoIncr_Load(object sender, System.EventArgs e)
		{
		
		}

		private void chkSucursal_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void optTipo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtValor_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbSucursal_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkDepartamento_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
