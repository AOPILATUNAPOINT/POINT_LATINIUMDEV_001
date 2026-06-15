using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Text;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmEdoProductosDT.
	/// </summary>
	public class frmEdoProductosDT : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnBuscar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Button btnTransferir;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCliente;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTransportista;
		private System.Windows.Forms.Label lblNumero;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtPlaca;
		private System.ComponentModel.IContainer components;
		private Acceso miAcceso;

		public frmEdoProductosDT()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Recibe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Daño");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmEdoProductosDT));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Recibe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Daño");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnTransferir = new System.Windows.Forms.Button();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.label36 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label31 = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbTransportista = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtPlaca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransportista)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(bool);
			ultraDataColumn2.DefaultValue = false;
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// btnBuscar
			// 
			this.btnBuscar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnBuscar.Location = new System.Drawing.Point(304, 56);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(77, 23);
			this.btnBuscar.TabIndex = 1;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnBuscar, "Buscar información...");
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15});
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
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 77;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "...";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 20;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 70;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Cédula";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 90;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 200;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 69;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 180;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 133;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 120;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 180;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 300;
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
																										 ultraGridColumn13});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 8.5F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 120);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1040, 240);
			this.ultraGrid1.TabIndex = 664;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// btnTransferir
			// 
			this.btnTransferir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnTransferir.Enabled = false;
			this.btnTransferir.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnTransferir.ForeColor = System.Drawing.Color.Black;
			this.btnTransferir.Image = ((System.Drawing.Image)(resources.GetObject("btnTransferir.Image")));
			this.btnTransferir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTransferir.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnTransferir.Location = new System.Drawing.Point(904, 80);
			this.btnTransferir.Name = "btnTransferir";
			this.btnTransferir.Size = new System.Drawing.Size(77, 23);
			this.btnTransferir.TabIndex = 5;
			this.btnTransferir.Text = "&Transferir";
			this.btnTransferir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnTransferir, "Transferir productos a DT");
			this.btnTransferir.Click += new System.EventHandler(this.btnTransferir_Click);
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label36.Location = new System.Drawing.Point(8, 35);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(30, 17);
			this.label36.TabIndex = 666;
			this.label36.Text = "Local";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 180;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 32);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(224, 22);
			this.cmbBodega.TabIndex = 3;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(8, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 17);
			this.label4.TabIndex = 669;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 2;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 59);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(38, 17);
			this.label1.TabIndex = 670;
			this.label1.Text = "Cliente";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label31
			// 
			this.label31.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label31.AutoSize = true;
			this.label31.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label31.Location = new System.Drawing.Point(592, 10);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(74, 17);
			this.label31.TabIndex = 672;
			this.label31.Text = "Persona Envia";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbVendedor
			// 
			this.cmbVendedor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataSource = this.ultraDataSource4;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Width = 250;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(760, 7);
			this.cmbVendedor.MaxDropDownItems = 30;
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(224, 22);
			this.cmbVendedor.TabIndex = 4;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVendedor_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// txtCliente
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCliente.Appearance = appearance6;
			this.txtCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCliente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCliente.Location = new System.Drawing.Point(64, 56);
			this.txtCliente.Name = "txtCliente";
			this.txtCliente.Size = new System.Drawing.Size(224, 22);
			this.txtCliente.TabIndex = 0;
			this.txtCliente.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCliente_KeyDown);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(304, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 675;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1056, 8);
			this.groupBox1.TabIndex = 676;
			this.groupBox1.TabStop = false;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(592, 35);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(70, 17);
			this.label2.TabIndex = 678;
			this.label2.Text = "Transportista";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTransportista
			// 
			this.cmbTransportista.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.cmbTransportista.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTransportista.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTransportista.DataSource = this.ultraDataSource2;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 250;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbTransportista.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTransportista.DisplayMember = "Nombre";
			this.cmbTransportista.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTransportista.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTransportista.Enabled = false;
			this.cmbTransportista.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTransportista.Location = new System.Drawing.Point(760, 32);
			this.cmbTransportista.MaxDropDownItems = 30;
			this.cmbTransportista.Name = "cmbTransportista";
			this.cmbTransportista.Size = new System.Drawing.Size(224, 22);
			this.cmbTransportista.TabIndex = 677;
			this.cmbTransportista.ValueMember = "idPersonal";
			this.cmbTransportista.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTransportista_KeyDown);
			this.cmbTransportista.ValueChanged += new System.EventHandler(this.cmbTransportista_ValueChanged);
			this.cmbTransportista.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTransportista_InitializeLayout);
			// 
			// lblNumero
			// 
			this.lblNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(592, 59);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(162, 17);
			this.lblNumero.TabIndex = 680;
			this.lblNumero.Text = "Numero De Guia (Servientrega)";
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance7;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(760, 56);
			this.txtNumero.MaxLength = 7;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 679;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(592, 83);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(90, 17);
			this.label3.TabIndex = 682;
			this.label3.Text = "Numero De Placa";
			// 
			// txtPlaca
			// 
			this.txtPlaca.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPlaca.Appearance = appearance8;
			this.txtPlaca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPlaca.Enabled = false;
			this.txtPlaca.Location = new System.Drawing.Point(760, 80);
			this.txtPlaca.MaxLength = 7;
			this.txtPlaca.Name = "txtPlaca";
			this.txtPlaca.Size = new System.Drawing.Size(104, 22);
			this.txtPlaca.TabIndex = 681;
			// 
			// frmEdoProductosDT
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1056, 366);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtPlaca);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbTransportista);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.txtCliente);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnTransferir);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.ultraGrid1);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmEdoProductosDT";
			this.Text = "Transferencia Servicio Tecnico";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmEdoProductosDT_KeyDown);
			this.Load += new System.EventHandler(this.frmEdoProductosDT_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTransportista)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		string NumeroTRIngreso = "";
		string NumeroTRSalida = "";
		bool bNuevo = false;
		int iBodegaPredefinida = 0;
		bool bAutomatico = true;
		
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmEdoProductosDT_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "1009");
			
			if (!Funcion.Permiso("738", cdsSeteoF))
			{				
				MessageBox.Show("No Puede Ingresar a Transferencias Servicio Tecnico", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			
			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			
			//this.cmbTransportista.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DT_Personal 28"); 
			this.cmbTransportista.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idChoferes as idPersonal,	Nombre FROM dbo.choferes order by nombre");   
			
						
			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			iBodegaPredefinida = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodegaPredefinida > 0) this.cmbBodega.Value = iBodegaPredefinida; else this.cmbBodega.Enabled = true;

			CargaVendedorLocal();
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{			
			int iBodega = 0;
						
			if(this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;
			
			string sSQL = string.Format("Exec EstadoProductosDT {0}, '{1}'", iBodega, this.txtCliente.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);

			this.lblContador.Text = this.ultraGrid1.Rows.Count + " REGISTROS ENCONTRADOS";
			
			CargaVendedorLocal();
		}

		public void CargaVendedorLocal()
		{
			int iBodega = 0;

			this.cmbVendedor.Value = System.DBNull.Value;
			this.cmbTransportista.Value = System.DBNull.Value;
			this.txtNumero.Text = "";

			this.cmbTransportista.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.txtNumero.Enabled = false;

			this.btnTransferir.Enabled = false; 
			
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;

			if (this.cmbBodega.ActiveRow != null) 
			{
				iBodega = (int)this.cmbBodega.Value;
				this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaUsuarios {0}", iBodega));

				this.cmbTransportista.Enabled = true;
				this.cmbVendedor.Enabled = true;
				this.btnTransferir.Enabled = true;				

				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}
		}

		private void ImprimeGuiaDeRemision(int idGuiaRemision)
		{
			ParameterFields paramFields = new ParameterFields ();

			ParameterField Compra = new ParameterField ();
			ParameterDiscreteValue IdCompra = new ParameterDiscreteValue ();
			//Compra.ParameterFieldName = "@idMaestroGuiaRemision";
			Compra.ParameterFieldName = "@idCPA";
			IdCompra.Value = idGuiaRemision;
			Compra.CurrentValues.Add (IdCompra);
			paramFields.Add (Compra);
			
			//Reporte miRep = new Reporte("GuiaDeRemisionCarpa.rpt", "");
			Reporte miRep = new Reporte("Doc_GR.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Guia de Remisión");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
			
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update MaestroGuiaRemision Set Impreso = 1 Where idMaestroGuiaRemision = {0}", idGuiaRemision));
		}

		private void btnTransferir_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay ingresos para transferir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCliente.Focus();
				return;
			}
			
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione la bodega origen de transferencia")) return;
			
			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione la persona que realiza la transferencia")) return;

			if (!Validacion.vbComboVacio(this.cmbTransportista, "Seleccione la persona que transporta la mercaderia")) return;
			
			if ((int)this.cmbTransportista.Value == 1013)
				if (!Validacion.vbTexto(this.txtNumero, 10, 10, "Numero de Guia de remision de Servientrega")) return;

			if ((int)this.cmbTransportista.Value != 1013)
				if (!Validacion.vbValidaPlaca(this.txtPlaca)) return;
		
//			if ((int)this.cmbTransportista.Value != 1013 && !bAutomatico)
//			{
//				if (!Validacion.vbTexto(this.txtNumero, 9, 9, "Guia de remision")) return;				
//
//				string sSQLVal = string.Format("Exec ValidaNumeracionGuiaDeRemision '{0}', {1}", this.txtNumero.Text.ToString(), (int)this.cmbBodega.Value);
//				bool bRes = Funcion.bEscalarSQL(cdsSeteoF, sSQLVal, true);
//				if (bRes)
//				{
//					MessageBox.Show("El Numero " + this.txtNumero.Text.ToString() + " ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.txtNumero.Focus();
//					return;
//				}
//			}

			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)			
				if ((bool)dr.Cells["Sel"].Value) iCont++;

			if (iCont == 0)
			{
				MessageBox.Show("Seleccione al menos un ingreso para crear la transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCliente.Focus();
				return;
			}
			#endregion Validacion

			if (DialogResult.Yes == MessageBox.Show("Esta seguro de generar la transferencia para servicio tecnico", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{								
				this.Cursor = Cursors.WaitCursor;
				try
				{
					#region Guia de remision
					int idGuiaRemision = 0;
					DateTime dtInicio = DateTime.Today;
					DateTime dtFin = DateTime.Today.AddDays(3);
					string sNumeroGR = "";
					
					if ((int)this.cmbTransportista.Value != 1013)
					{	
						if ((int)this.cmbBodega.Value == 95) this.cmbBodega.Value = 156;

						if (bAutomatico) sNumeroGR = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 30, {0}, 0", (int)this.cmbBodega.Value));					
				
						string sSQL = string.Format("Exec GGuiaRemision {0}, {1}, '{2}', '{3}', {4}, '{5}', {6}, {7}, {8}, '{9}', '{10}', {11}, '{12}', '{13}'", 
							0, 0, dtInicio.ToString("yyyyMMdd"), dtFin.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, sNumeroGR, 42, 5, 
							(int)this.cmbTransportista.Value, this.txtPlaca.Text.ToString(), "DOCUMENTO GENERADO POR TRANSFERENCIA DE PRODUCTOS A SERVICIO TECNICO", 20, "SERVICIO TECNICO", 0);				
						idGuiaRemision = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
					}
					else sNumeroGR = this.txtNumero.Text.ToString();
					#endregion Guia de remision

					#region Detalle Guia de Remision - Movimiento
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
					{
						if ((bool)dr.Cells["Sel"].Value) 
						{
							#region Detalle	Guia de Remision
							if ((int)this.cmbTransportista.Value != 1013)
							{
								string sSQLDetalle = string.Format("Exec GuardaGuiaDeRemision {0}, {1}, {2}, {3}",
									0, idGuiaRemision, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idSerial"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle, true);
							}
							#endregion Detalle Guia de Remision
							
							#region Guarda Movimiento
							string sSQLProcesos = string.Format("Exec GuardaDT_Procesos {0}, {1}, 2, {2}, 'ENVIO SERVICIO TECNICO POINT', {3}, '{4}', 0, 0, 0", 
								(int)dr.Cells["idIngreso"].Value, (int)this.cmbBodega.Value, (int)this.cmbVendedor.Value, idGuiaRemision, sNumeroGR);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLProcesos, true);	
							#endregion Guarda Movimiento
						}					
					}
					#endregion Detalle Guia de Remision - Movimiento			

					string sClaveAcceso = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec FE_ClaveDeAcceso {0}, 6", idGuiaRemision));

					string cadInvertida = Facturacion.invertirCadena(sClaveAcceso.Substring(0, 48));
					string ValM11 = Facturacion.Modulo11(cadInvertida).ToString().Trim();
					string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update MaestroGuiaRemision Set Autorizacion = '{0}' Where idMaestroGuiaRemision = {1}", sClaveSRI, idGuiaRemision));

					ImprimeDocumento(idGuiaRemision, sNumeroGR);

					if ((int)this.cmbTransportista.Value != 1013) 
						ImprimeGuiaDeRemision(idGuiaRemision);
								
					MessageBox.Show("Transferencia procesada correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				
					CargaVendedorLocal();
					
					this.btnBuscar_Click(sender, e);					
				}
				catch (Exception ex)
				{
					MessageBox.Show(string.Format("Exception Type: {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
				}
				this.Cursor = Cursors.Default;
			}			
		}

		public void ImprimeDocumento(int idGuiaRemision, string sNumero)
		{			
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pIdGuiaRemision = new ParameterField ();
			ParameterDiscreteValue IdGuiaRemision = new ParameterDiscreteValue ();
			pIdGuiaRemision.ParameterFieldName = "@idGuiaRemision";
			IdGuiaRemision.Value = idGuiaRemision;
			pIdGuiaRemision.CurrentValues.Add (IdGuiaRemision);
			paramFields.Add (pIdGuiaRemision);
				
			ParameterField pNumero = new ParameterField ();
			ParameterDiscreteValue Numero = new ParameterDiscreteValue ();
			pNumero.ParameterFieldName = "@NumeroGuia";
			Numero.Value = sNumero;
			pNumero.CurrentValues.Add (Numero);
			paramFields.Add (pNumero);

			Reporte Reporte = new Reporte("TranferenciasDT.rpt", "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{			
			this.btnBuscar_Click(sender, e);

			CargaVendedorLocal();		

			bAutomatico = false;
			if (this.cmbBodega.ActiveRow != null)
				bAutomatico = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select AutGuiaRemision From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);		
		}

		private void frmEdoProductosDT_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 50, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCliente.Focus();
		}

		private void txtCliente_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnBuscar_Click(sender, e);
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTransportista.Focus();
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void cmbTransportista_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtNumero.Text = "";
			this.txtNumero.Enabled = false;
			this.txtPlaca.Enabled = false; 

			if (this.cmbTransportista.ActiveRow != null)
				if ((int)this.cmbTransportista.Value == 1013) this.txtNumero.Enabled = true; 
				else if ((int)this.cmbTransportista.Value != 1013 || !bAutomatico) this.txtNumero.Enabled = true;

			if (this.cmbTransportista.ActiveRow != null)
				if ((int)this.cmbTransportista.Value == 1013) this.txtPlaca.Enabled = false; 
				else if ((int)this.cmbTransportista.Value != 1013) 
				{
					this.txtPlaca.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Select Placa from choferes where idChoferes = {0} ",
						(int) this.cmbTransportista.Value));	
					this.txtPlaca.Enabled = false;
				}
	
		
		}

		private void cmbTransportista_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.txtNumero.Enabled) this.txtNumero.Focus(); else this.btnTransferir_Click(sender, e);
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnTransferir.Focus();
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNumero.Text.ToString().Trim().Length > 0)
			{
				char Pad = '0';
				int iCar = 9;

				if ((int)this.cmbTransportista.Value == 1013) iCar = 10;

				this.txtNumero.Text = this.txtNumero.Value.ToString().PadLeft(iCar, Pad);
			}
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbTransportista_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
