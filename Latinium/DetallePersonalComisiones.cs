using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de DetallePersonalComisiones.
	/// </summary>
	public class DetallePersonalComisiones : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button cmbVendedores;
		private System.Windows.Forms.Button cmbCajeros;
		private System.Windows.Forms.Button cmbGuardias;
		private System.Windows.Forms.Button cmbTodos;
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataSet cdsSeteoRepF;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataSet cdsSucursal;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.RadioButton radioButton1;
		private System.Windows.Forms.RadioButton radioButton2;
		private Infragistics.Win.UltraWinGrid.UltraGrid grPerVCG;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCentrodCostos;
		private C1.Data.C1DataSet cdsProyecto;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DetallePersonalComisiones()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbCentrodCostos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolCargo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CajaChica");
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbTodos = new System.Windows.Forms.Button();
			this.cmbGuardias = new System.Windows.Forms.Button();
			this.cmbCajeros = new System.Windows.Forms.Button();
			this.cmbVendedores = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cdsSeteoRepF = new C1.Data.C1DataSet();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.grPerVCG = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSucursal = new C1.Data.C1DataSet();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.radioButton2 = new System.Windows.Forms.RadioButton();
			this.radioButton1 = new System.Windows.Forms.RadioButton();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCentrodCostos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRepF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grPerVCG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentrodCostos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cmbTodos);
			this.groupBox1.Controls.Add(this.cmbGuardias);
			this.groupBox1.Controls.Add(this.cmbCajeros);
			this.groupBox1.Controls.Add(this.cmbVendedores);
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(392, 64);
			this.groupBox1.TabIndex = 164;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Buscar por categoría";
			// 
			// cmbTodos
			// 
			this.cmbTodos.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.cmbTodos.Location = new System.Drawing.Point(288, 24);
			this.cmbTodos.Name = "cmbTodos";
			this.cmbTodos.TabIndex = 4;
			this.cmbTodos.Text = "TODOS";
			this.cmbTodos.Click += new System.EventHandler(this.cmbTodos_Click);
			// 
			// cmbGuardias
			// 
			this.cmbGuardias.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.cmbGuardias.Location = new System.Drawing.Point(200, 24);
			this.cmbGuardias.Name = "cmbGuardias";
			this.cmbGuardias.TabIndex = 3;
			this.cmbGuardias.Text = "Guardias";
			this.cmbGuardias.Click += new System.EventHandler(this.cmbGuardias_Click);
			// 
			// cmbCajeros
			// 
			this.cmbCajeros.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.cmbCajeros.Location = new System.Drawing.Point(112, 24);
			this.cmbCajeros.Name = "cmbCajeros";
			this.cmbCajeros.TabIndex = 2;
			this.cmbCajeros.Text = "Cajeros";
			this.cmbCajeros.Click += new System.EventHandler(this.cmbCajeros_Click);
			// 
			// cmbVendedores
			// 
			this.cmbVendedores.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.cmbVendedores.Location = new System.Drawing.Point(24, 24);
			this.cmbVendedores.Name = "cmbVendedores";
			this.cmbVendedores.TabIndex = 1;
			this.cmbVendedores.Text = "Vendedores";
			this.cmbVendedores.Click += new System.EventHandler(this.cmbVendedores_Click);
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
			// cdsSeteoRepF
			// 
			this.cdsSeteoRepF.BindingContextCtrl = this;
			this.cdsSeteoRepF.DataLibrary = "LibFacturacion";
			this.cdsSeteoRepF.DataLibraryUrl = "";
			this.cdsSeteoRepF.DataSetDef = "dsSeteoF";
			this.cdsSeteoRepF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoRepF.Name = "cdsSeteoRepF";
			this.cdsSeteoRepF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoRepF.SchemaDef = null;
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsClientes";
			this.cdsCliente.FillOnRequest = false;
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// grPerVCG
			// 
			this.grPerVCG.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grPerVCG.Cursor = System.Windows.Forms.Cursors.Default;
			this.grPerVCG.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grPerVCG.DisplayLayout.Appearance = appearance1;
			this.grPerVCG.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 37;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 256;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 125;
			ultraGridColumn4.Header.Caption = "Centro de Costo";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 149;
			ultraGridColumn5.Header.Caption = "Cargo";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 129;
			ultraGridColumn6.Header.Caption = "F. Ingreso";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownCalendar;
			ultraGridColumn6.Width = 73;
			ultraGridColumn7.Header.Caption = "F. Salida";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownCalendar;
			ultraGridColumn7.Width = 78;
			ultraGridColumn8.Header.Caption = "Estd";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn8.Width = 49;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand1.SummaryFooterCaption = "Total GENERAL";
			this.grPerVCG.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grPerVCG.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grPerVCG.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.grPerVCG.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grPerVCG.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance4.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grPerVCG.DisplayLayout.Override.RowAlternateAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grPerVCG.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grPerVCG.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.grPerVCG.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grPerVCG.Location = new System.Drawing.Point(8, 104);
			this.grPerVCG.Name = "grPerVCG";
			this.grPerVCG.Size = new System.Drawing.Size(880, 384);
			this.grPerVCG.TabIndex = 165;
			this.grPerVCG.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grPerVCG_CellChange);
			this.grPerVCG.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grPerVCG_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargos.DataMember = "RolCargo";
			this.cmbCargos.DataSource = this.cdsSucursal;
			ultraGridBand2.AddButtonCaption = "RolVendedores";
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Width = 200;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCargos.DisplayMember = "Descripcion";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargos.Location = new System.Drawing.Point(24, 352);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(176, 86);
			this.cmbCargos.TabIndex = 166;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
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
			this.cdsSucursal.BeforeFill += new C1.Data.FillEventHandler(this.cdsSucursal_BeforeFill);
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
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.radioButton2);
			this.groupBox2.Controls.Add(this.radioButton1);
			this.groupBox2.Location = new System.Drawing.Point(416, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(152, 48);
			this.groupBox2.TabIndex = 167;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Estado";
			// 
			// radioButton2
			// 
			this.radioButton2.Location = new System.Drawing.Point(80, 16);
			this.radioButton2.Name = "radioButton2";
			this.radioButton2.Size = new System.Drawing.Size(64, 24);
			this.radioButton2.TabIndex = 1;
			this.radioButton2.Text = "Inactivo";
			// 
			// radioButton1
			// 
			this.radioButton1.Checked = true;
			this.radioButton1.Location = new System.Drawing.Point(16, 16);
			this.radioButton1.Name = "radioButton1";
			this.radioButton1.Size = new System.Drawing.Size(56, 24);
			this.radioButton1.TabIndex = 0;
			this.radioButton1.TabStop = true;
			this.radioButton1.Text = "Activo";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 80);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 16);
			this.label1.TabIndex = 169;
			this.label1.Text = "Nombre:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance7;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(64, 78);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(264, 21);
			this.txtNombre.TabIndex = 168;
			this.txtNombre.TextChanged += new System.EventHandler(this.txtNombre_TextChanged);
			// 
			// cmbCentrodCostos
			// 
			this.cmbCentrodCostos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentrodCostos.DataMember = "Proyecto";
			this.cmbCentrodCostos.DataSource = this.cdsProyecto;
			ultraGridBand3.AddButtonCaption = "RolVendedores";
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Width = 190;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 10;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 9;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 7;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 8;
			ultraGridColumn24.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbCentrodCostos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCentrodCostos.DisplayMember = "Nombre";
			this.cmbCentrodCostos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentrodCostos.Location = new System.Drawing.Point(208, 352);
			this.cmbCentrodCostos.Name = "cmbCentrodCostos";
			this.cmbCentrodCostos.Size = new System.Drawing.Size(176, 86);
			this.cmbCentrodCostos.TabIndex = 170;
			this.cmbCentrodCostos.ValueMember = "idProyecto";
			this.cmbCentrodCostos.Visible = false;
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
			this.cdsProyecto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// DetallePersonalComisiones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(896, 493);
			this.Controls.Add(this.cmbCentrodCostos);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.cmbCargos);
			this.Controls.Add(this.grPerVCG);
			this.Controls.Add(this.groupBox1);
			this.Name = "DetallePersonalComisiones";
			this.Text = "Detalle personal de Locales Comerciales";
			this.Load += new System.EventHandler(this.DetallePersonalComisiones_Load);
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRepF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grPerVCG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentrodCostos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		C1.Data.C1DataRow dr;
		C1.Data.C1DataRow drSeteoF;
		//CultureInfo us = new CultureInfo("en-US");

		int SwCambio = 0;
		int Proyecto = 0;
		int Cargo = 0;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		

		}

		private void DetallePersonalComisiones_Load(object sender, System.EventArgs e)
		{
			# region Datos personal
			//string stExec = string.Format("Exec POINT_Porcentaje_Comision {0}",idComsion);
			//grdComisiones.DataSource = Funcion.dvProcedimiento(cdsComision, stExec);
			# endregion Datos personal
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCliente.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirRoles);

		}

		private void cdsSucursal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSucursal.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirRoles);		

		}

		private void cmbVendedores_Click(object sender, System.EventArgs e)
		{
			int tpbuskda = 0;
			string ClteBusk = (string)this.txtNombre.Value;
			if (this.radioButton1.Checked)
				tpbuskda = 1;
			else
				tpbuskda = 0;

			string stExec = string.Format("Exec Buscar_V_C_G {0},{1},'{2}'",1,tpbuskda,ClteBusk);
			this.grPerVCG.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);

			SwCambio = 1;

		}

		private void cmbCajeros_Click(object sender, System.EventArgs e)
		{
			int tpbuskda = 0;
			string ClteBusk = (string)this.txtNombre.Value;
			if (this.radioButton1.Checked)
				tpbuskda = 1;
			else
				tpbuskda = 0;

			string stExec = string.Format("Exec Buscar_V_C_G {0},{1},'{2}'",2,tpbuskda,ClteBusk);
			this.grPerVCG.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);

			SwCambio = 2;

		}

		private void cmbGuardias_Click(object sender, System.EventArgs e)
		{
			int tpbuskda = 0;
			string ClteBusk = (string)this.txtNombre.Value;
			if (this.radioButton1.Checked)
				tpbuskda = 1;
			else
				tpbuskda = 0;

			string stExec = string.Format("Exec Buscar_V_C_G {0},{1},'{2}'",3,tpbuskda,ClteBusk);
			this.grPerVCG.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);

			SwCambio = 3;

		}

		private void cmbTodos_Click(object sender, System.EventArgs e)
		{
			int tpbuskda = 0;
			string ClteBusk = (string)this.txtNombre.Value;
			if (this.radioButton1.Checked)
				tpbuskda = 1;
			else
				tpbuskda = 0;

			string stExec = string.Format("Exec Buscar_V_C_G {0},{1},'{2}'",4,tpbuskda,ClteBusk);
			this.grPerVCG.DataSource = Funcion.dvProcedimiento(cdsSeteoF, stExec);

			SwCambio = 4;
		
		}

		private void grPerVCG_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int idPers = (int)this.grPerVCG.ActiveRow.Cells["idPersonal"].Value;

			#region Cambio Estado
			//int idPers = 0;
			int ValorEstado = 0;
			if (e.Cell.Column.ToString()=="Estado") 
			{
				//idPers = (int)this.grPerVCG.ActiveRow.Cells["idPersonal"].Value;
				if ((int)this.grPerVCG.ActiveRow.Cells["Estado"].Value==0)
          ValorEstado = 1;
				else
					ValorEstado = 0;
				string stExec = string.Format("Exec Actualiza_EstadoPersonal {0},{1},{2}",idPers,ValorEstado,1);
				Funcion.EjecutaSQL(cdsSeteoF, stExec,true);
			}
			#endregion Cambio Estado

			#region Cambio Personal
			int Proyct = 0;
			if (e.Cell.Column.ToString()=="idProyecto")
			{
				idPers = (int)e.Cell.Row.Cells["idPersonal"].Value;
				if (e.Cell.Row.Cells["idProyecto"].Value!=System.DBNull.Value) //this.grPerVCG.ActiveRow.Cells["idProyecto"].Value == DBNull.Value)
					Proyct = (int)e.Cell.Row.Cells["idProyecto"].Value;
				//else Proyct = (int)e.Cell.Row.Cells["idProyecto"].Value;
				string stExec2p = string.Format("Exec Actualiza_EstadoPersonal {0},{1},{2}",idPers,Proyct,2);
				Funcion.EjecutaSQL(cdsSeteoF, stExec2p,true);
			}
      #endregion Cambio Personal

			//if (false) Refresca(sender,e);			
		}

		private void txtNombre_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void grPerVCG_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
//      int Proyct = 0;
//			if (e.Cell.Column.ToString()=="idProyecto")
//			{
//        int idPers = (int)e.Cell.Row.Cells["idPersonal"].Value;
//				if (e.Cell.Row.Cells["idProyecto"].Value!=System.DBNull.Value) //this.grPerVCG.ActiveRow.Cells["idProyecto"].Value == DBNull.Value)
//					Proyct = (int)e.Cell.Row.Cells["idProyecto"].Value;
//				//else Proyct = (int)e.Cell.Row.Cells["idProyecto"].Value;
//				string stExec2p = string.Format("Exec Actualiza_EstadoPersonal {0},{1},{2}",idPers,Proyct,2);
//				Funcion.EjecutaSQL(cdsSeteoF, stExec2p,true);
				
			//}
		}

		private void Refresca(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (SwCambio == 1) this.cmbVendedores_Click(sender,e);
			if (SwCambio == 2) this.cmbCajeros_Click(sender,e);
			if (SwCambio == 3) this.cmbGuardias_Click(sender,e);
			if (SwCambio == 4) this.cmbTodos_Click(sender,e);
		}
	}
}
