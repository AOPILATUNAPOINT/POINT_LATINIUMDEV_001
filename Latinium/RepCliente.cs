using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepCliente.
	/// </summary>
	public class RepCliente : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private C1.Data.C1DataSet cdsClienteTable;
		private C1.Data.C1DataView cdvGrupo;
		private C1.Data.C1DataView cdvCiudad;
		private C1.Data.C1DataView cdvPrecio;
		private C1.Data.C1DataView cdvSector;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCiudad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodaCiudad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPrecio;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoPrecio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSector;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoSector;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVendedor;
		private System.Windows.Forms.ErrorProvider errorProvider;
		public Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPublicidad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPublicidad;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbDesde;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
		private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProcedimiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RepCliente()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteCiudad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClientePrecio", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteSector", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClientePublicidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Medio");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProcPantalla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Procedimiento");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProcPantalla");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Procedimiento");
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.chkTodoGrupo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvGrupo = new C1.Data.C1DataView();
			this.cdsClienteTable = new C1.Data.C1DataSet();
			this.cdvCiudad = new C1.Data.C1DataView();
			this.cdvPrecio = new C1.Data.C1DataView();
			this.cdvSector = new C1.Data.C1DataView();
			this.cmbCiudad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkTodaCiudad = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbPrecio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkTodoPrecio = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbSector = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkTodoSector = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.chkVendedor = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbPublicidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkPublicidad = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbDesde = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.cmbProcedimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClienteTable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPublicidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProcedimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(224, 296);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(56, 24);
			this.btAceptar.TabIndex = 0;
			this.btAceptar.Text = "&Reporte";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// chkTodoGrupo
			// 
			this.chkTodoGrupo.Checked = true;
			this.chkTodoGrupo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodoGrupo.Location = new System.Drawing.Point(24, 43);
			this.chkTodoGrupo.Name = "chkTodoGrupo";
			this.chkTodoGrupo.Size = new System.Drawing.Size(88, 22);
			this.chkTodoGrupo.TabIndex = 2;
			this.chkTodoGrupo.Text = "Todo Grupo";
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataSource = this.cdvGrupo;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 470;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(128, 43);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(152, 21);
			this.cmbGrupo.TabIndex = 3;
			this.cmbGrupo.ValueMember = "idGrupoCliente";
			// 
			// cdvGrupo
			// 
			this.cdvGrupo.BindingContextCtrl = this;
			this.cdvGrupo.DataSet = this.cdsClienteTable;
			this.cdvGrupo.Sort = "Grupo";
			this.cdvGrupo.TableName = "";
			this.cdvGrupo.TableViewName = "ClienteGrupo";
			// 
			// cdsClienteTable
			// 
			this.cdsClienteTable.BindingContextCtrl = this;
			this.cdsClienteTable.DataLibrary = "LibFacturacion";
			this.cdsClienteTable.DataLibraryUrl = "";
			this.cdsClienteTable.DataSetDef = "dsClienteTabla";
			this.cdsClienteTable.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsClienteTable.Name = "cdsClienteTable";
			this.cdsClienteTable.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsClienteTable.SchemaDef = null;
			// 
			// cdvCiudad
			// 
			this.cdvCiudad.BindingContextCtrl = this;
			this.cdvCiudad.DataSet = this.cdsClienteTable;
			this.cdvCiudad.Sort = "Ciudad";
			this.cdvCiudad.TableName = "";
			this.cdvCiudad.TableViewName = "ClienteCiudad";
			// 
			// cdvPrecio
			// 
			this.cdvPrecio.BindingContextCtrl = this;
			this.cdvPrecio.DataSet = this.cdsClienteTable;
			this.cdvPrecio.Sort = "Precio";
			this.cdvPrecio.TableName = "";
			this.cdvPrecio.TableViewName = "ClientePrecio";
			// 
			// cdvSector
			// 
			this.cdvSector.BindingContextCtrl = this;
			this.cdvSector.DataSet = this.cdsClienteTable;
			this.cdvSector.Sort = "Sector";
			this.cdvSector.TableName = "";
			this.cdvSector.TableViewName = "ClienteSector";
			// 
			// cmbCiudad
			// 
			this.cmbCiudad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCiudad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCiudad.DataSource = this.cdvCiudad;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Width = 387;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 99;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbCiudad.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCiudad.DisplayMember = "Ciudad";
			this.cmbCiudad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCiudad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCiudad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCiudad.Location = new System.Drawing.Point(128, 78);
			this.cmbCiudad.Name = "cmbCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(152, 21);
			this.cmbCiudad.TabIndex = 5;
			this.cmbCiudad.ValueMember = "idCiudad";
			// 
			// chkTodaCiudad
			// 
			this.chkTodaCiudad.Checked = true;
			this.chkTodaCiudad.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodaCiudad.Location = new System.Drawing.Point(24, 78);
			this.chkTodaCiudad.Name = "chkTodaCiudad";
			this.chkTodaCiudad.Size = new System.Drawing.Size(88, 21);
			this.chkTodaCiudad.TabIndex = 4;
			this.chkTodaCiudad.Text = "Toda Ciudad";
			// 
			// cmbPrecio
			// 
			this.cmbPrecio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPrecio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPrecio.DataSource = this.cdvPrecio;
			this.cmbPrecio.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 133;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbPrecio.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPrecio.DisplayMember = "Precio";
			this.cmbPrecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPrecio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPrecio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPrecio.Location = new System.Drawing.Point(128, 112);
			this.cmbPrecio.Name = "cmbPrecio";
			this.cmbPrecio.Size = new System.Drawing.Size(152, 21);
			this.cmbPrecio.TabIndex = 7;
			this.cmbPrecio.ValueMember = "idTipoPrecio";
			// 
			// chkTodoPrecio
			// 
			this.chkTodoPrecio.Checked = true;
			this.chkTodoPrecio.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodoPrecio.Location = new System.Drawing.Point(24, 112);
			this.chkTodoPrecio.Name = "chkTodoPrecio";
			this.chkTodoPrecio.Size = new System.Drawing.Size(88, 22);
			this.chkTodoPrecio.TabIndex = 6;
			this.chkTodoPrecio.Text = "Todo Precio";
			// 
			// cmbSector
			// 
			this.cmbSector.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSector.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSector.DataSource = this.cdvSector;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Width = 434;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbSector.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbSector.DisplayMember = "Sector";
			this.cmbSector.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSector.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSector.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSector.Location = new System.Drawing.Point(128, 146);
			this.cmbSector.Name = "cmbSector";
			this.cmbSector.Size = new System.Drawing.Size(152, 21);
			this.cmbSector.TabIndex = 9;
			this.cmbSector.ValueMember = "idSector";
			// 
			// chkTodoSector
			// 
			this.chkTodoSector.Checked = true;
			this.chkTodoSector.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodoSector.Location = new System.Drawing.Point(24, 146);
			this.chkTodoSector.Name = "chkTodoSector";
			this.chkTodoSector.Size = new System.Drawing.Size(88, 22);
			this.chkTodoSector.TabIndex = 8;
			this.chkTodoSector.Text = "Todo Sector";
			// 
			// cmbVendedor
			// 
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 300;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(128, 181);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(152, 21);
			this.cmbVendedor.TabIndex = 11;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.RowFilter = "Vendedor = 1";
			this.cdvPersonal.Sort = "Nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.FillOnRequest = false;
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// chkVendedor
			// 
			this.chkVendedor.Checked = true;
			this.chkVendedor.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkVendedor.Location = new System.Drawing.Point(24, 181);
			this.chkVendedor.Name = "chkVendedor";
			this.chkVendedor.Size = new System.Drawing.Size(106, 21);
			this.chkVendedor.TabIndex = 10;
			this.chkVendedor.Text = "Todo Vendedor";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// optTipo
			// 
			this.optTipo.CheckedIndex = 0;
			this.optTipo.ItemAppearance = appearance1;
			this.optTipo.ItemOrigin = new System.Drawing.Point(6, 4);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Clientes";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Proveedores";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Todos";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.Items.Add(valueListItem3);
			this.optTipo.Location = new System.Drawing.Point(24, 9);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(256, 25);
			this.optTipo.TabIndex = 12;
			this.optTipo.Text = "Clientes";
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
			// cmbPublicidad
			// 
			this.cmbPublicidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPublicidad.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbPublicidad.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbPublicidad.DisplayMember = "Medio";
			this.cmbPublicidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPublicidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPublicidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPublicidad.Location = new System.Drawing.Point(128, 216);
			this.cmbPublicidad.Name = "cmbPublicidad";
			this.cmbPublicidad.Size = new System.Drawing.Size(152, 21);
			this.cmbPublicidad.TabIndex = 14;
			this.cmbPublicidad.ValueMember = "idClientePublicidad";
			// 
			// chkPublicidad
			// 
			this.chkPublicidad.Checked = true;
			this.chkPublicidad.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPublicidad.Location = new System.Drawing.Point(24, 216);
			this.chkPublicidad.Name = "chkPublicidad";
			this.chkPublicidad.Size = new System.Drawing.Size(106, 21);
			this.chkPublicidad.TabIndex = 13;
			this.chkPublicidad.Text = "Toda Publicidad";
			// 
			// cmbDesde
			// 
			this.cmbDesde.FormatString = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(48, 256);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.PromptChar = ' ';
			this.cmbDesde.Size = new System.Drawing.Size(112, 22);
			this.cmbDesde.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbDesde.TabIndex = 15;
			// 
			// cmbHasta
			// 
			this.cmbHasta.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(168, 256);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.PromptChar = ' ';
			this.cmbHasta.Size = new System.Drawing.Size(112, 22);
			this.cmbHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbHasta.TabIndex = 16;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 256);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(22, 16);
			this.label1.TabIndex = 17;
			this.label1.Text = "De:";
			// 
			// gridControl1
			// 
			this.gridControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// gridControl1.EmbeddedNavigator
			// 
			this.gridControl1.EmbeddedNavigator.Name = "";
			this.gridControl1.Location = new System.Drawing.Point(296, 40);
			this.gridControl1.MainView = this.gridView1;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
																																																					this.repositoryItemLookUpEdit1});
			this.gridControl1.Size = new System.Drawing.Size(424, 288);
			this.gridControl1.TabIndex = 18;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																																this.gridView1});
			// 
			// gridView1
			// 
			this.gridView1.GridControl = this.gridControl1;
			this.gridView1.Name = "gridView1";
			this.gridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Bottom;
			// 
			// repositoryItemLookUpEdit1
			// 
			this.repositoryItemLookUpEdit1.AutoHeight = false;
			this.repositoryItemLookUpEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																					 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.repositoryItemLookUpEdit1.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																							 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Nombre", "Nombre", 43, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.repositoryItemLookUpEdit1.DisplayMember = "Nombre";
			this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
			this.repositoryItemLookUpEdit1.NullText = "";
			this.repositoryItemLookUpEdit1.ValueMember = "idTipoFactura";
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(456, 8);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(56, 24);
			this.btGenerar.TabIndex = 19;
			this.btGenerar.Text = "&Listado";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// cmbProcedimiento
			// 
			this.cmbProcedimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProcedimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProcedimiento.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 575;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbProcedimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbProcedimiento.DisplayMember = "Procedimiento";
			this.cmbProcedimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProcedimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProcedimiento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProcedimiento.Location = new System.Drawing.Point(296, 8);
			this.cmbProcedimiento.Name = "cmbProcedimiento";
			this.cmbProcedimiento.Size = new System.Drawing.Size(152, 21);
			this.cmbProcedimiento.TabIndex = 22;
			this.cmbProcedimiento.ValueMember = "idProcPantalla";
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 133;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbReporte.DisplayMember = "Procedimiento";
			this.cmbReporte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbReporte.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(24, 296);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(152, 21);
			this.cmbReporte.TabIndex = 23;
			this.cmbReporte.ValueMember = "idProcPantalla";
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(600, 8);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(56, 24);
			this.ultraButton1.TabIndex = 24;
			this.ultraButton1.Text = "&Factura";
			// 
			// ultraButton2
			// 
			this.ultraButton2.Location = new System.Drawing.Point(664, 8);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.Size = new System.Drawing.Size(56, 24);
			this.ultraButton2.TabIndex = 25;
			this.ultraButton2.Text = "&Cliente";
			// 
			// RepCliente
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(736, 342);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.cmbReporte);
			this.Controls.Add(this.cmbProcedimiento);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.gridControl1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.cmbPublicidad);
			this.Controls.Add(this.chkPublicidad);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.chkVendedor);
			this.Controls.Add(this.cmbSector);
			this.Controls.Add(this.chkTodoSector);
			this.Controls.Add(this.cmbPrecio);
			this.Controls.Add(this.chkTodoPrecio);
			this.Controls.Add(this.cmbCiudad);
			this.Controls.Add(this.chkTodaCiudad);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.chkTodoGrupo);
			this.Controls.Add(this.btAceptar);
			this.Name = "RepCliente";
			this.Text = "Reporte de Clientes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RepCliente_Closing);
			this.Load += new System.EventHandler(this.RepCliente_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsClienteTable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPublicidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProcedimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private bool Verifica()
		{
			errorProvider.SetError(cmbProcedimiento, "");
			errorProvider.SetError(cmbCiudad, "");
			errorProvider.SetError(cmbGrupo, "");
			errorProvider.SetError(cmbPrecio, "");
			errorProvider.SetError(cmbSector, "");
			errorProvider.SetError(cmbVendedor, "");
			errorProvider.SetError(cmbPublicidad, "");
			errorProvider.SetError(cmbReporte, "");

			bool bOk = true;
			if (!this.chkTodaCiudad.Checked)
				if (this.cmbCiudad.Value == null)
				{
					bOk = false;
					this.errorProvider.SetError(this.cmbCiudad, "Ingrese Ciudad");
				}
			if (!this.chkTodoGrupo.Checked)
				if (this.cmbGrupo.Value == null)
				{
					bOk = false;
					this.errorProvider.SetError(this.cmbGrupo, "Ingrese Grupo");
				}
			if (!this.chkTodoPrecio.Checked)
				if (this.cmbPrecio.Value == null)
				{
					bOk = false;
					this.errorProvider.SetError(this.cmbPrecio, "Ingrese Precio");
				}
			if (!this.chkTodoSector.Checked)
				if (this.cmbSector.Value == null)
				{
					bOk = false;
					this.errorProvider.SetError(this.cmbSector, "Ingrese Sector");
				}
			if (!this.chkVendedor.Checked)
				if (this.cmbVendedor.Value == null)
				{
					bOk = false;
					this.errorProvider.SetError(this.cmbVendedor, "Ingrese Vendedor");
				}
			if (!this.chkPublicidad.Checked)
				if (this.cmbPublicidad.Value == null)
				{
					bOk = false;
					this.errorProvider.SetError(this.cmbPublicidad, "Ingrese Publicidad");
				}
			return bOk;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			if (cmbReporte.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbReporte, "Seleccione reporte a mostrar");
				return;
			}

			Cursor = Cursors.WaitCursor;
			string stTipo = "false";
			string stTitulo = "Clientes";
			if ((int) optTipo.Value == 1) 
			{
				stTipo = "true";
				stTitulo = "Proveedores";
			}
			string stFiltro = "({Cliente.Proveedor} = " + stTipo + " Or {Cliente.Ambos} = true)";
			if ((int) optTipo.Value == 2) stFiltro = "";

			if (!this.chkTodaCiudad.Checked)
				stFiltro += " AND {Cliente.idCiudad}=" + this.cmbCiudad.Value.ToString().Trim();
			if (!this.chkTodoGrupo.Checked)
				stFiltro += " AND {Cliente.idGrupoCliente}=" + this.cmbGrupo.Value.ToString().Trim();
			if (!this.chkTodoPrecio.Checked)
				stFiltro += " AND {Cliente.idTipoPrecio}=" + this.cmbPrecio.Value.ToString().Trim();
			if (!this.chkTodoSector.Checked)
				stFiltro += " AND {Cliente.idSector}=" + this.cmbSector.Value.ToString().Trim();
			if (!this.chkVendedor.Checked)
				stFiltro += " AND {Cliente.idVendedor}=" + this.cmbVendedor.Value.ToString().Trim();
			if (!this.chkPublicidad.Checked)
				stFiltro += " AND {Cliente.idClientePublicidad}=" + this.cmbPublicidad.Value.ToString().Trim();

			Reporte miReporte = new Reporte(cmbReporte.Text, stFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Titulo(stTitulo);
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void RepCliente_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'RepCliente'";
			Funcion.EjecutaSQL(cdsClienteTable, stAudita, true);
			cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Where Vendedor = 1");
			cmbPublicidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idClientePublicidad, Medio From ClientePublicidad Order By Medio");
			cmbProcedimiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProcPantalla, Procedimiento From ProcPantalla Where Pantalla = 'RepCliente' And Tipo = 1");
			cmbReporte.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProcPantalla, Procedimiento From ProcPantalla Where Pantalla = 'RepCliente' And Tipo = 2");

			Cursor = Cursors.Default;
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		
		}

		private void RepCliente_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'RepCliente'";
			Funcion.EjecutaSQL(cdsClienteTable, stAudita, true);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			if (cmbProcedimiento.Text.Trim().Length == 0)
			{
				errorProvider.SetError(cmbProcedimiento, "Seleccione procedimiento a mostrar");
				return;
			}

			int idGrupo = 0;
			int idCiudad = 0;
			int idPrecio = 0;
			int idSector = 0;
			int idVendedor = 0;
			int idPublicidad = 0;
			if (!chkTodoGrupo.Checked) idGrupo = (int) cmbGrupo.Value;
			if (!chkPublicidad.Checked) idPublicidad = (int) cmbPublicidad.Value;
			if (!chkTodaCiudad.Checked) idCiudad = (int) cmbCiudad.Value;
			if (!chkTodoPrecio.Checked) idPrecio = (int) cmbPrecio.Value;
			if (!chkTodoSector.Checked) idSector = (int) cmbSector.Value;
			if (!chkVendedor.Checked) idVendedor = (int) cmbVendedor.Value;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			int idTipo = 1;
			string stProcedimiento = cmbProcedimiento.Text;
			string stExec = string.Format("Exec {9} {0}, {1}, {2}, {3}, {4}, {5}, '{6}', '{7}', {8}",
				idGrupo, idCiudad, idPrecio, idSector, idVendedor, idPublicidad, dtDesde.ToString("yyyyMMdd"), 
				dtHasta.ToString("yyyyMMdd"), idTipo, stProcedimiento);
			gridControl1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);
		}
	}
}
