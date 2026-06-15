using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RetRenta.
	/// </summary>
	public class RetIVA : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsRetencion;
		private C1.Data.C1DataSet cdsCuenta;
		private C1.Data.C1DataView cdvCuenta;
		private C1.Data.C1DataView cdvRetIVA;
		private C1.Data.C1DataView cdvCodCuenta;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCodCuenta;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaD;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btNoAsignar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RetIVA()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cuenta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodRapido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodPresup");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescIngles");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaEjecucionP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaBalance");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RetencionIva", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetIva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaI", -1, "cmbCuentaD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaE", -1, "cmbCuentaD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodEgreso", 0, "cmbCodCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso", 1, "cmbCodCuenta");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.cdvRetIVA = new C1.Data.C1DataView();
			this.cdsRetencion = new C1.Data.C1DataSet();
			this.cdvCodCuenta = new C1.Data.C1DataView();
			this.cdsCuenta = new C1.Data.C1DataSet();
			this.cdvCuenta = new C1.Data.C1DataView();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.cmbCodCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCuentaD = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btNoAsignar = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cdvRetIVA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// cdvRetIVA
			// 
			this.cdvRetIVA.BindingContextCtrl = this;
			this.cdvRetIVA.DataSet = this.cdsRetencion;
			this.cdvRetIVA.TableName = "";
			this.cdvRetIVA.TableViewName = "RetencionIva";
			// 
			// cdsRetencion
			// 
			this.cdsRetencion.BindingContextCtrl = this;
			this.cdsRetencion.DataLibrary = "LibContabilidad";
			this.cdsRetencion.DataLibraryUrl = "";
			this.cdsRetencion.DataSetDef = "dsRetencion";
			this.cdsRetencion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRetencion.Name = "cdsRetencion";
			this.cdsRetencion.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsRetencion.SchemaDef = null;
			this.cdsRetencion.BeforeFill += new C1.Data.FillEventHandler(this.cdsRetencion_BeforeFill);
			// 
			// cdvCodCuenta
			// 
			this.cdvCodCuenta.BindingContextCtrl = this;
			this.cdvCodCuenta.DataSet = this.cdsCuenta;
			this.cdvCodCuenta.Sort = "Codigo";
			this.cdvCodCuenta.TableName = "";
			this.cdvCodCuenta.TableViewName = "Cuenta";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuenta2";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			this.cdsCuenta.BeforeFill += new C1.Data.FillEventHandler(this.cdsCuenta_BeforeFill);
			// 
			// cdvCuenta
			// 
			this.cdvCuenta.BindingContextCtrl = this;
			this.cdvCuenta.DataSet = this.cdsCuenta;
			this.cdvCuenta.Sort = "Descripcion";
			this.cdvCuenta.TableName = "";
			this.cdvCuenta.TableViewName = "Cuenta";
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(472, 225);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 12;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btAceptar.Location = new System.Drawing.Point(96, 225);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 11;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// cmbCodCuenta
			// 
			this.cmbCodCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCodCuenta.DataSource = this.cdvCodCuenta;
			ultraGridColumn1.Header.VisiblePosition = 8;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 12;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn4.Header.VisiblePosition = 9;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 170;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.cmbCodCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCodCuenta.DisplayMember = "Codigo";
			this.cmbCodCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodCuenta.Location = new System.Drawing.Point(32, 99);
			this.cmbCodCuenta.Name = "cmbCodCuenta";
			this.cmbCodCuenta.Size = new System.Drawing.Size(280, 80);
			this.cmbCodCuenta.TabIndex = 10;
			this.cmbCodCuenta.ValueMember = "idCuenta";
			this.cmbCodCuenta.Visible = false;
			// 
			// cmbCuentaD
			// 
			this.cmbCuentaD.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaD.DataSource = this.cdvCuenta;
			ultraGridColumn20.Header.VisiblePosition = 8;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 12;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn23.Header.VisiblePosition = 9;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Width = 196;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 3;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 5;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 7;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 6;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 10;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 11;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 13;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 14;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 15;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 16;
			ultraGridColumn37.Header.VisiblePosition = 17;
			ultraGridColumn38.Header.VisiblePosition = 18;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38});
			this.cmbCuentaD.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuentaD.DisplayMember = "Descripcion";
			this.cmbCuentaD.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaD.Location = new System.Drawing.Point(336, 115);
			this.cmbCuentaD.Name = "cmbCuentaD";
			this.cmbCuentaD.Size = new System.Drawing.Size(296, 80);
			this.cmbCuentaD.TabIndex = 9;
			this.cmbCuentaD.ValueMember = "idCuenta";
			this.cmbCuentaD.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvRetIVA;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn39.Header.VisiblePosition = 7;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 53;
			ultraGridColumn40.Header.Caption = "Venta Cuenta";
			ultraGridColumn40.Header.VisiblePosition = 6;
			ultraGridColumn40.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn40.Width = 120;
			ultraGridColumn41.Header.Caption = "Compra Cuenta";
			ultraGridColumn41.Header.VisiblePosition = 4;
			ultraGridColumn41.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn41.Width = 134;
			ultraGridColumn42.Header.VisiblePosition = 0;
			ultraGridColumn42.Width = 36;
			ultraGridColumn43.Header.VisiblePosition = 2;
			ultraGridColumn43.Width = 53;
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Width = 106;
			ultraGridColumn45.Header.Caption = "Compra Codigo";
			ultraGridColumn45.Header.VisiblePosition = 3;
			ultraGridColumn45.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn45.Width = 91;
			ultraGridColumn46.Header.Caption = "Venta Codigo";
			ultraGridColumn46.Header.VisiblePosition = 5;
			ultraGridColumn46.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn46.Width = 79;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.Gainsboro;
			appearance4.BackColor2 = System.Drawing.Color.Silver;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 19);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(640, 193);
			this.ultraGrid1.TabIndex = 8;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			// 
			// btNoAsignar
			// 
			this.btNoAsignar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btNoAsignar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btNoAsignar.Location = new System.Drawing.Point(192, 225);
			this.btNoAsignar.Name = "btNoAsignar";
			this.btNoAsignar.TabIndex = 13;
			this.btNoAsignar.Text = "&No Asignar";
			this.btNoAsignar.Click += new System.EventHandler(this.btNoAsignar_Click);
			// 
			// RetIVA
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(656, 270);
			this.Controls.Add(this.btNoAsignar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbCodCuenta);
			this.Controls.Add(this.cmbCuentaD);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "RetIVA";
			this.Text = "Porcentaje de Retencion del Iva";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RetIVA_Closing);
			this.Load += new System.EventHandler(this.RetIVA_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvRetIVA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCodCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.cdsRetencion.Update();
		}

		private void simpleButton2_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsRetencion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRetencion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRetenc);
		}

		private void cdsCuenta_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCuenta.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.IsInEditMode)
			{
				if (e.Cell.Column.ToString() == "CodEgreso")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["idCuentaE"].Value = e.Cell.Row.Cells["CodEgreso"].Value;
					}
				if (e.Cell.Column.ToString() == "CodIngreso")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["idCuentaI"].Value = e.Cell.Row.Cells["CodIngreso"].Value;
					}
			}
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["CodEgreso"].Value = e.Row.Cells["idCuentaE"].Value;
			e.Row.Cells["CodIngreso"].Value = e.Row.Cells["idCuentaI"].Value;
		}

		private void btNoAsignar_Click(object sender, System.EventArgs e)
		{
			this.ultraGrid1.ActiveRow.Cells["idCuentaE"].Value = 0;
			this.ultraGrid1.ActiveRow.Cells["idCuentaI"].Value = 0;
		}

		private void RetIVA_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Ret. Iva'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
			Cursor = Cursors.Default;
		}

		private void RetIVA_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Ret. Iva'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
		}

	}
}
