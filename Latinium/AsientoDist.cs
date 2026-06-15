using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de AsientoDist.
	/// </summary>
	public class AsientoDist : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsAsientoDistrib;
		private DevExpress.XtraEditors.SimpleButton btAceptar;
		private DevExpress.XtraEditors.SimpleButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private C1.Data.C1DataSet cdsCentroC;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCentro;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDistribucion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiGrupoDist;
		private C1.Data.C1DataView cdvCentro;
		private C1.Data.C1DataView cdvProyecto;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public AsientoDist()
		{
			InitializeComponent();
		}

		private int IdDetAsiento = 0;
		public AsientoDist(int idDetAsiento)
		{
			InitializeComponent();
			IdDetAsiento = idDetAsiento;
		}

		private bool bCompra = false;
		public AsientoDist(int idDetAsiento, bool BCompra)
		{
			InitializeComponent();
			IdDetAsiento = idDetAsiento;
			bCompra = BCompra;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_AsientoDist", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistrib");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbCentro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto", -1, "cmbProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("SubProyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RespNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsAsientoDistrib = new C1.Data.C1DataSet();
			this.btAceptar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cdsCentroC = new C1.Data.C1DataSet();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.cmbCentro = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdvCentro = new C1.Data.C1DataView();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtDistribucion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.txtIdAsiGrupoDist = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsientoDistrib)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCentroC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCentro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDistribucion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiGrupoDist)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_AsientoDistGrupo.Asiento Dist";
			this.ultraGrid1.DataSource = this.cdsAsientoDistrib;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "AsientoDist";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 132;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 219;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 59;
			ultraGridColumn4.Header.Caption = "Centro C.";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn4.Width = 81;
			ultraGridColumn5.Header.Caption = "Proyecto";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn5.Width = 108;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 79;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(488, 216);
			this.ultraGrid1.TabIndex = 0;
			// 
			// cdsAsientoDistrib
			// 
			this.cdsAsientoDistrib.BindingContextCtrl = this;
			this.cdsAsientoDistrib.DataLibrary = "LibContabilidad";
			this.cdsAsientoDistrib.DataLibraryUrl = "";
			this.cdsAsientoDistrib.DataSetDef = "dsAsientoDist";
			this.cdsAsientoDistrib.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsientoDistrib.Name = "cdsAsientoDistrib";
			this.cdsAsientoDistrib.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsientoDistrib.SchemaDef = null;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.Location = new System.Drawing.Point(360, 288);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(64, 23);
			this.btAceptar.TabIndex = 1;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(432, 288);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(64, 23);
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			// 
			// optTipo
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipo.Appearance = appearance7;
			this.optTipo.CheckedIndex = 0;
			this.optTipo.ImeMode = System.Windows.Forms.ImeMode.On;
			this.optTipo.ItemAppearance = appearance8;
			this.optTipo.ItemOrigin = new System.Drawing.Point(8, 3);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Aplicar a Linea  ";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Aplicar a Asiento";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.Location = new System.Drawing.Point(248, 0);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(240, 24);
			this.optTipo.TabIndex = 4;
			this.optTipo.Text = "Aplicar a Linea  ";
			// 
			// cdsCentroC
			// 
			this.cdsCentroC.BindingContextCtrl = this;
			this.cdsCentroC.DataLibrary = "LibContabilidad";
			this.cdsCentroC.DataLibraryUrl = "";
			this.cdsCentroC.DataSetDef = "dsProyecto";
			this.cdsCentroC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCentroC.Name = "cdsCentroC";
			this.cdsCentroC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCentroC.SchemaDef = null;
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataSource = this.cdvProyecto;
			ultraGridBand2.AddButtonCaption = "Proyecto";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 173;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(216, 96);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(272, 80);
			this.cmbProyecto.TabIndex = 6;
			this.cmbProyecto.Text = "ultraDropDown1";
			this.cmbProyecto.ValueMember = "idSubProyecto";
			this.cmbProyecto.Visible = false;
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsCentroC;
			this.cdvProyecto.RowFilter = "Visible = 1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "SubProyecto";
			// 
			// cmbCentro
			// 
			this.cmbCentro.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentro.DataSource = this.cdvCentro;
			ultraGridBand3.AddButtonCaption = "SubProyecto";
			ultraGridColumn15.Header.VisiblePosition = 2;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Width = 188;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 7;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn22.Header.VisiblePosition = 5;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 8;
			ultraGridColumn23.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbCentro.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCentro.DisplayMember = "Nombre";
			this.cmbCentro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentro.Location = new System.Drawing.Point(32, 152);
			this.cmbCentro.Name = "cmbCentro";
			this.cmbCentro.Size = new System.Drawing.Size(288, 72);
			this.cmbCentro.TabIndex = 7;
			this.cmbCentro.Text = "ultraDropDown2";
			this.cmbCentro.ValueMember = "idProyecto";
			this.cmbCentro.Visible = false;
			// 
			// cdvCentro
			// 
			this.cdvCentro.BindingContextCtrl = this;
			this.cdvCentro.DataSet = this.cdsCentroC;
			this.cdvCentro.RowFilter = "Visible = 1";
			this.cdvCentro.Sort = "Nombre";
			this.cdvCentro.TableName = "";
			this.cdvCentro.TableViewName = "Proyecto";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDatoSQL1.BarraMovimiento = true;
			this.barraDatoSQL1.DataMember = "_AsientoDistGrupo";
			this.barraDatoSQL1.DataNombreId = "idAsiDistGrupo";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsAsientoDistrib;
			this.barraDatoSQL1.DataTabla = "AsientoDistGrupo";
			this.barraDatoSQL1.DataTablaHija = "AsientoDist";
			this.barraDatoSQL1.Location = new System.Drawing.Point(16, 288);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(264, 24);
			this.barraDatoSQL1.TabIndex = 8;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = false;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = false;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// txtDistribucion
			// 
			this.txtDistribucion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsientoDistrib, "_AsientoDistGrupo.Distribucion"));
			this.txtDistribucion.Location = new System.Drawing.Point(104, 24);
			this.txtDistribucion.Name = "txtDistribucion";
			this.txtDistribucion.Size = new System.Drawing.Size(384, 21);
			this.txtDistribucion.TabIndex = 9;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(16, 24);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(67, 14);
			this.ultraLabel1.TabIndex = 10;
			this.ultraLabel1.Text = "Distribucion:";
			// 
			// txtIdAsiGrupoDist
			// 
			this.txtIdAsiGrupoDist.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsAsientoDistrib, "_AsientoDistGrupo.idAsiDistGrupo"));
			this.txtIdAsiGrupoDist.Location = new System.Drawing.Point(136, 0);
			this.txtIdAsiGrupoDist.Name = "txtIdAsiGrupoDist";
			this.txtIdAsiGrupoDist.PromptChar = ' ';
			this.txtIdAsiGrupoDist.Size = new System.Drawing.Size(100, 21);
			this.txtIdAsiGrupoDist.TabIndex = 11;
			// 
			// AsientoDist
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(512, 326);
			this.Controls.Add(this.txtIdAsiGrupoDist);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtDistribucion);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.cmbCentro);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "AsientoDist";
			this.Text = "Distribucion de Asiento";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.AsientoDist_Closing);
			this.Load += new System.EventHandler(this.AsientoDist_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsientoDistrib)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCentroC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCentro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDistribucion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiGrupoDist)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void AsientoDist_Load(object sender, System.EventArgs e)
		{
			optTipo.Width = 0;
			txtIdAsiGrupoDist.Width = 0;
			barraDatoSQL1.ProximoId(1);
			barraDatoSQL1.HabilitaControles(false);
			string stAudita = "Exec AuditaCrear 53, 6, 'Asi. Dist.'";
			Funcion.EjecutaSQL(cdsAsientoDistrib, stAudita, true);
		}

		private void BorraErrores()
		{
			errorProvider.SetError(ultraGrid1, "");
			errorProvider.SetError(txtDistribucion, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			double dPorcentaje = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["Porcentaje"].Value != System.DBNull.Value)
					dPorcentaje += (double) dr.Cells["Porcentaje"].Value;
			}
			BorraErrores();
			if (txtDistribucion.Text.Trim().Length == 0)
			{
				errorProvider.SetError(txtDistribucion, "Ingrese Descripcion de Distribucion");
				bOk = false;
			}
			if (Math.Round(dPorcentaje, 2) != 100.0) 
			{
				errorProvider.SetError(ultraGrid1, "Sum de Porcentaje es: " 
					+ dPorcentaje.ToString("0.00") +" debe ser 100");
				bOk = false;
			}
			return bOk;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (barraDatoSQL1.bEditar == true)
			{
				MessageBox.Show("Primero grabe el registro que esta editando", "Información");
				return;
			}
			string stSelect = "";
			if (!bCompra)
			{
				stSelect = "Exec AsientoDistrib " + IdDetAsiento.ToString() 
					+ ", " + optTipo.Value.ToString() + ", " + txtIdAsiGrupoDist.Value.ToString();
			}
			else
			{
				stSelect = "Exec CompraDistrib " + IdDetAsiento.ToString() 
					+ ", " + optTipo.Value.ToString() + ", " + txtIdAsiGrupoDist.Value.ToString();

			}
			Funcion.EjecutaSQL(cdsAsientoDistrib, stSelect, true);
			Close();
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
			string stSelect = "Select Count(*) From Personal Where IdAsiDistGrupo = "
				+ txtIdAsiGrupoDist.Value.ToString();
			int iCuenta = Funcion.iEscalarSQL(cdsAsientoDistrib, stSelect);
			if (iCuenta > 0)
			{
				MessageBox.Show("No puede borrar Distribucion porque tiene relacionados "
					+ iCuenta.ToString() + " Registros en Personal");
				return;
			}

			if (DialogResult.No == MessageBox.Show("¿Está seguro que desea borrar el registro?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				return;
			string stBorrar = "Delete AsientoDist Where IdAsiDistGrupo = " + txtIdAsiGrupoDist.Value.ToString();
			Funcion.EjecutaSQL(cdsAsientoDistrib, stBorrar);
			barraDatoSQL1.BorraRegistro();
		}

		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			barraDatoSQL1.CrearRegistro();
			barraDatoSQL1.GrabaRegistro();
			barraDatoSQL1.EditarRegistro();
			barraDatoSQL1_Refresca(this, e);
			barraDatoSQL1.bNuevo = true;
		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			cdsAsientoDistrib.RejectChanges();
			barraDatoSQL1.DeshacerRegistro();		
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			barraDatoSQL1.EditarRegistro();
		}

		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			barraDatoSQL1.GrabaRegistro();
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void AsientoDist_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Asi. Dist.'";
			Funcion.EjecutaSQL(cdsAsientoDistrib, stAudita, true);
		}
	}
}
