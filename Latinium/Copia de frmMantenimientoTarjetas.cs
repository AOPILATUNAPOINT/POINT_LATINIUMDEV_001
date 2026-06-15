using System;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Shared;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using System.Diagnostics;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmMantenimientoTarjetas.
	/// </summary>
	public class frmMantenimientoTarjetas : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridBancos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTarjetas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPlanes;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPlazos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Button btnBancos;
		private System.Windows.Forms.Button btnTarjetas;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridRegion;
		private System.Windows.Forms.Button btnNuevo;

		private C1DataRow drSeteoF;

		public frmMantenimientoTarjetas()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlan");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcComision");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcInteres");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcInteres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRegion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRegion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.uGridBancos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridTarjetas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridPlanes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridPlazos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnTarjetas = new System.Windows.Forms.Button();
			this.btnBancos = new System.Windows.Forms.Button();
			this.uGridRegion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnNuevo = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridBancos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTarjetas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlanes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlazos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRegion)).BeginInit();
			this.SuspendLayout();
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
			// uGridBancos
			// 
			this.uGridBancos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridBancos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridBancos.DisplayLayout.Appearance = appearance1;
			this.uGridBancos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 70;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.Caption = "Banco";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 203;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 24;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			ultraGridBand1.Header.Caption = "Bancos";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.uGridBancos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridBancos.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridBancos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridBancos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridBancos.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridBancos.DisplayLayout.Override.CellAppearance = appearance4;
			this.uGridBancos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridBancos.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBancos.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBancos.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridBancos.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridBancos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridBancos.Location = new System.Drawing.Point(176, 32);
			this.uGridBancos.Name = "uGridBancos";
			this.uGridBancos.Size = new System.Drawing.Size(248, 376);
			this.uGridBancos.TabIndex = 166;
			this.uGridBancos.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridBancos_BeforeSelectChange);
			this.uGridBancos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridBancos_CellChange);
			this.uGridBancos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridBancos_InitializeLayout);
			this.uGridBancos.Click += new System.EventHandler(this.uGridBancos_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// uGridTarjetas
			// 
			this.uGridTarjetas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTarjetas.DataSource = this.ultraDataSource2;
			appearance9.BackColor = System.Drawing.Color.White;
			this.uGridTarjetas.DisplayLayout.Appearance = appearance9;
			this.uGridTarjetas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 47;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.Caption = "Tarjeta";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 140;
			ultraGridColumn6.Header.Caption = "...";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 23;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			ultraGridBand2.Header.Caption = "Tarjetas";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand2.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.uGridTarjetas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTarjetas.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.uGridTarjetas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridTarjetas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridTarjetas.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTarjetas.DisplayLayout.Override.CellAppearance = appearance12;
			this.uGridTarjetas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTarjetas.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTarjetas.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTarjetas.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTarjetas.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.uGridTarjetas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTarjetas.Location = new System.Drawing.Point(432, 32);
			this.uGridTarjetas.Name = "uGridTarjetas";
			this.uGridTarjetas.Size = new System.Drawing.Size(184, 376);
			this.uGridTarjetas.TabIndex = 167;
			this.uGridTarjetas.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridTarjetas_BeforeSelectChange);
			this.uGridTarjetas.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridTarjetas_CellChange);
			this.uGridTarjetas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridTarjetas_InitializeLayout);
			this.uGridTarjetas.Click += new System.EventHandler(this.uGridTarjetas_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(bool);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// uGridPlanes
			// 
			this.uGridPlanes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPlanes.DataSource = this.ultraDataSource3;
			appearance17.BackColor = System.Drawing.Color.White;
			this.uGridPlanes.DisplayLayout.Appearance = appearance17;
			this.uGridPlanes.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 32;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.Caption = "Plan";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 115;
			ultraGridColumn9.Header.Caption = "...";
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 25;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			ultraGridBand3.Header.Caption = "Planes";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand3.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.uGridPlanes.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPlanes.DisplayLayout.Override.ActiveRowAppearance = appearance18;
			this.uGridPlanes.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPlanes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance19.BackColor = System.Drawing.Color.Transparent;
			this.uGridPlanes.DisplayLayout.Override.CardAreaAppearance = appearance19;
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPlanes.DisplayLayout.Override.CellAppearance = appearance20;
			this.uGridPlanes.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance21.FontData.BoldAsString = "True";
			appearance21.FontData.Name = "Arial";
			appearance21.FontData.SizeInPoints = 8F;
			appearance21.ForeColor = System.Drawing.Color.White;
			appearance21.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPlanes.DisplayLayout.Override.HeaderAppearance = appearance21;
			appearance22.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance22.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanes.DisplayLayout.Override.RowAlternateAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanes.DisplayLayout.Override.RowSelectorAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanes.DisplayLayout.Override.SelectedRowAppearance = appearance24;
			this.uGridPlanes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPlanes.Location = new System.Drawing.Point(624, 32);
			this.uGridPlanes.Name = "uGridPlanes";
			this.uGridPlanes.Size = new System.Drawing.Size(161, 376);
			this.uGridPlanes.TabIndex = 168;
			this.uGridPlanes.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridPlanes_BeforeSelectChange);
			this.uGridPlanes.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPlanes_CellChange);
			this.uGridPlanes.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridPlanes_InitializeLayout);
			this.uGridPlanes.Click += new System.EventHandler(this.uGridPlanes_Click);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(bool);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// uGridPlazos
			// 
			this.uGridPlazos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPlazos.DataSource = this.ultraDataSource4;
			appearance25.BackColor = System.Drawing.Color.White;
			this.uGridPlazos.DisplayLayout.Appearance = appearance25;
			this.uGridPlazos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 36;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.Caption = "Plazo";
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 151;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance26;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "% Comisión";
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.PromptChar = ' ';
			ultraGridColumn12.Width = 58;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance27;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "% Interes";
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 55;
			ultraGridColumn14.Header.Caption = "...";
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Width = 29;
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Width = 79;
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn16.Width = 79;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			ultraGridBand4.Header.Caption = "Plazos";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand4.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			this.uGridPlazos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPlazos.DisplayLayout.Override.ActiveRowAppearance = appearance28;
			this.uGridPlazos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPlazos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance29.BackColor = System.Drawing.Color.Transparent;
			this.uGridPlazos.DisplayLayout.Override.CardAreaAppearance = appearance29;
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPlazos.DisplayLayout.Override.CellAppearance = appearance30;
			this.uGridPlazos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance31.FontData.BoldAsString = "True";
			appearance31.FontData.Name = "Arial";
			appearance31.FontData.SizeInPoints = 8F;
			appearance31.ForeColor = System.Drawing.Color.White;
			appearance31.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPlazos.DisplayLayout.Override.HeaderAppearance = appearance31;
			appearance32.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance32.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlazos.DisplayLayout.Override.RowAlternateAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlazos.DisplayLayout.Override.RowSelectorAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlazos.DisplayLayout.Override.SelectedRowAppearance = appearance34;
			this.uGridPlazos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPlazos.Location = new System.Drawing.Point(792, 32);
			this.uGridPlazos.Name = "uGridPlazos";
			this.uGridPlazos.Size = new System.Drawing.Size(472, 376);
			this.uGridPlazos.TabIndex = 169;
			this.uGridPlazos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPlazos_CellChange);
			this.uGridPlazos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPlazos_AfterCellUpdate);
			this.uGridPlazos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridPlazos_InitializeLayout);
			this.uGridPlazos.Click += new System.EventHandler(this.uGridPlazos_Click);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn13.DataType = typeof(System.Double);
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn15.DataType = typeof(System.DateTime);
			ultraDataColumn16.DataType = typeof(System.DateTime);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// ultraDataSource5
			// 
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(bool);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// btnTarjetas
			// 
			this.btnTarjetas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnTarjetas.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnTarjetas.Location = new System.Drawing.Point(512, 5);
			this.btnTarjetas.Name = "btnTarjetas";
			this.btnTarjetas.Size = new System.Drawing.Size(75, 24);
			this.btnTarjetas.TabIndex = 649;
			this.btnTarjetas.Text = "Tarjetas";
			this.btnTarjetas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnTarjetas.Click += new System.EventHandler(this.btnTarjetas_Click);
			// 
			// btnBancos
			// 
			this.btnBancos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnBancos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBancos.Location = new System.Drawing.Point(240, 5);
			this.btnBancos.Name = "btnBancos";
			this.btnBancos.Size = new System.Drawing.Size(75, 24);
			this.btnBancos.TabIndex = 648;
			this.btnBancos.Text = "Bancos";
			this.btnBancos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBancos.Click += new System.EventHandler(this.btnBancos_Click);
			// 
			// uGridRegion
			// 
			this.uGridRegion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridRegion.DataSource = this.ultraDataSource5;
			appearance35.BackColor = System.Drawing.Color.White;
			this.uGridRegion.DisplayLayout.Appearance = appearance35;
			this.uGridRegion.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 44;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 114;
			ultraGridColumn19.Header.Caption = "...";
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn19.Width = 25;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			ultraGridBand5.Header.Caption = "Region";
			ultraGridBand5.HeaderVisible = true;
			ultraGridBand5.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridRegion.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance36.ForeColor = System.Drawing.Color.Black;
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridRegion.DisplayLayout.Override.ActiveRowAppearance = appearance36;
			this.uGridRegion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridRegion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridRegion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance37.BackColor = System.Drawing.Color.Transparent;
			this.uGridRegion.DisplayLayout.Override.CardAreaAppearance = appearance37;
			this.uGridRegion.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance38.FontData.BoldAsString = "True";
			appearance38.FontData.Name = "Arial";
			appearance38.FontData.SizeInPoints = 9F;
			appearance38.ForeColor = System.Drawing.Color.White;
			appearance38.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRegion.DisplayLayout.Override.HeaderAppearance = appearance38;
			appearance39.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance39.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRegion.DisplayLayout.Override.RowAlternateAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRegion.DisplayLayout.Override.RowSelectorAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRegion.DisplayLayout.Override.SelectedRowAppearance = appearance41;
			this.uGridRegion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRegion.Location = new System.Drawing.Point(8, 32);
			this.uGridRegion.Name = "uGridRegion";
			this.uGridRegion.Size = new System.Drawing.Size(160, 376);
			this.uGridRegion.TabIndex = 650;
			this.uGridRegion.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridRegion_BeforeSelectChange);
			this.uGridRegion.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridRegion_CellChange);
			this.uGridRegion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridRegion_InitializeLayout);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(1040, 5);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(56, 24);
			this.btnNuevo.TabIndex = 651;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// frmMantenimientoTarjetas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(1266, 416);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.uGridRegion);
			this.Controls.Add(this.btnTarjetas);
			this.Controls.Add(this.btnBancos);
			this.Controls.Add(this.uGridPlazos);
			this.Controls.Add(this.uGridPlanes);
			this.Controls.Add(this.uGridTarjetas);
			this.Controls.Add(this.uGridBancos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmMantenimientoTarjetas";
			this.Text = "Mantenimiento de Tarjetas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmMantenimientoTarjetas_KeyDown);
			this.Load += new System.EventHandler(this.frmMantenimientoTarjetas_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridBancos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridTarjetas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlanes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlazos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRegion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmMantenimientoTarjetas_Load(object sender, System.EventArgs e)
		{
			this.uGridRegion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idRegion, Region , Convert (Bit ,0) as  Sel  from region ");
			//this.uGridBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Order By Nombre Asc");
			FuncionesProcedimientos.EstadoGrids(true,this.uGridRegion);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void uGridBancos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridBancos_Click(object sender, System.EventArgs e)
		{
			
		}

		private void uGridTarjetas_Click(object sender, System.EventArgs e)
		{
			
		}

		private void uGridPlanes_Click(object sender, System.EventArgs e)
		{
			
		}

		private void uGridTarjetas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridTarjetas_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			//if (this.uGridTarjetas.ActiveRow.Cells["idTarjeta"].Value == DBNull.Value) return;		
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridTarjetas.UpdateData();

				string sSQLAct = string.Format("Exec ActualizaMantenimientoTarjetas {0}, {1}, 0, 0, 0, 0, 0, {2},'20150101','20181231'", 
					(int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value, (int)e.Cell.Row.Cells["idTarjeta"].Value, (int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);

//				string sSQL = string.Format("Exec ListaTarjetasPlanes {0}", (int)e.Cell.Row.Cells["idTarjeta"].Value);
//				this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			}
		}

		private void uGridPlanes_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{				
			if (e.Cell.Column.ToString() == "Sel")
			{				
				this.uGridPlanes.UpdateData();

				string sSQLAct = string.Format("Exec ActualizaMantenimientoTarjetas {0}, {1}, {2}, 0, 0, 0, 1, {3},'20150101','20181231'", 
					(int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value, (int)this.uGridTarjetas.ActiveRow.Cells["idTarjeta"].Value, 
					(int)e.Cell.Row.Cells["idPlan"].Value, (int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
	
//				string sSQL = string.Format("Exec ListaPlanesPlazos {0}", (int)e.Cell.Row.Cells["idPlan"].Value);
//				this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);		
			}
		}
		
		private void uGridPlazos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			//if (this.uGridPlazos.ActiveRow.Cells["idPlazo"].Value == DBNull.Value) return;
			
			if (e.Cell.Column.ToString() == "Sel")
			{		
				if (e.Cell.Row.Cells["Desde"].Value == DBNull.Value)
				{
					MessageBox.Show("Ingrese la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridPlazos.ActiveRow.Selected = true;
					e.Cell.Row.Cells["Sel"].Value = false;
					return;
				}
				if (e.Cell.Row.Cells["Hasta"].Value == DBNull.Value)
				{
					MessageBox.Show("Ingrese la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cell.Row.Cells["Sel"].Value = false;
					this.uGridPlazos.ActiveRow.Selected = true;

					return;
				}

				if((bool)e.Cell.Row.Cells["Sel"].Value == false)
				{
					if ((DateTime)e.Cell.Row.Cells["Desde"].Value > (DateTime)e.Cell.Row.Cells["Hasta"].Value )
					{
						MessageBox.Show("La Fecha Inicial es mayor a la Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						//					e.Cell.Row.Cells["Hasta"].Value = (int)e.Cell.OriginalValue;	
						this.uGridPlazos.ActiveRow.Selected = true;
						e.Cell.Row.Cells["Sel"].Value = false;
						return;
					}			
				}

				this.uGridPlazos.UpdateData();
				
				string sSQL = string.Format("Exec ActualizaMantenimientoTarjetas {0}, {1}, {2}, {3}, {4}, {5}, 2, {6}, '{7}', '{8}'", 
					(int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value, (int)this.uGridTarjetas.ActiveRow.Cells["idTarjeta"].Value, 
					(int)this.uGridPlanes.ActiveRow.Cells["idPlan"].Value, (int)e.Cell.Row.Cells["idPlazo"].Value, 
					Convert.ToDecimal(e.Cell.Row.Cells["PorcComision"].Value), Convert.ToDecimal(e.Cell.Row.Cells["PorcInteres"].Value),
					(int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).ToString("yyyyMMdd"),Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd"));
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	
		  
//				if ((int)e.Cell.Row.Cells["idPlazo"].Value == 7 )
//				{	
//          e.Cell.Row.Cells["Sel"].Value = true;
//					string SQL = string.Format("Exec ActualizaMantenimientoTarjetas {0}, {1}, {2}, {3}, {4}, {5}, 2, {6}, '{7}', '{8}'", 
//						(int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value, (int)this.uGridTarjetas.ActiveRow.Cells["idTarjeta"].Value, 
//						(int)this.uGridPlanes.ActiveRow.Cells["idPlan"].Value, (int)e.Cell.Row.Cells["idPlazo"].Value, 
//						Convert.ToDecimal(e.Cell.Row.Cells["PorcComision"].Value), Convert.ToDecimal(e.Cell.Row.Cells["PorcInteres"].Value),
//						(int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).ToString("yyyyMMdd"),Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd"));
//					Funcion.EjecutaSQL(cdsSeteoF, SQL, true);	
//				}
			}			
		}

		private void frmMantenimientoTarjetas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();
		}

		private void uGridPlazos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "PorcComision", 2);
			FuncionesProcedimientos.FormatoGrid(e, "PorcInteres", 2);
		}

		private void uGridPlanes_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridPlazos_Click(object sender, System.EventArgs e)
		{
//			string sSQL = string.Format("Select dbo.PorcComisionTarjeta({0}, {1}, {2}, {3})", (int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value, (int)this.uGridTarjetas.ActiveRow.Cells["idTarjeta"].Value, 	(int)this.uGridPlanes.ActiveRow.Cells["idPlan"].Value,	(int)this.uGridPlazos.ActiveRow.Cells["idPlazo"].Value);
//			this.txtComision.Value = Funcion.dEscalarSQL(cdsSeteoF, sSQL);
		}

	

		private void uGridBancos_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.uGridBancos.ActiveRow.Cells["idBanco"].Value == DBNull.Value) return;
			
			this.uGridTarjetas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTarjeta, Nombre From TarjetasDeCredito Where idTarjeta = 0");
			this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " Select idPlan, Descripcion From Planes Where idPlan = 0");
			this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlazo, Descripcion, 0 As PorcComision, 0 As PorcInteres, 0 As Sel,  Convert (Datetime,'20181231') Desde ,  Convert (Datetime,'20181231') as Hasta  From Plazos Where idPlazo = 0");

			if ((bool)this.uGridBancos.ActiveRow.Cells["Sel"].Value)
			{
				string sSQL = string.Format("Exec ConsultaMantenimientoTarjetas {0}, 0, 0, 0, {1}", (int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value, (int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value);
				this.uGridTarjetas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);			
			}
		}

		private void uGridBancos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{		
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridBancos.UpdateData();

				string sSQLAct = string.Format("Exec ActualizaMantenimientoTarjetas {0}, 0, 0, 0, 0, 0, 3, {1},'20150101','20180131'", 
					(int)e.Cell.Row.Cells["idBanco"].Value, (int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);				
			}
		}

		private void uGridTarjetas_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.uGridTarjetas.ActiveRow.Cells["idTarjeta"].Value == DBNull.Value) return;

			this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " Select idPlan, Descripcion From Planes Where idPlan = 0");
			this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlazo, Descripcion, 0 As PorcComision, 0 As PorcInteres, 0 As Sel, Convert (Datetime,'20150101') as Desde ,  Convert (Datetime,'20181231') as Hasta  From Plazos Where idPlazo = 0");

			if ((bool)this.uGridTarjetas.ActiveRow.Cells["Sel"].Value)
			{
				string sSQL = string.Format("Exec ConsultaMantenimientoTarjetas {0}, {1}, 0, 1, {2}", 
					(int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value, (int)this.uGridTarjetas.ActiveRow.Cells["idTarjeta"].Value, (int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value);
				this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);				
			}
		}

		private void uGridPlanes_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.uGridPlanes.ActiveRow.Cells["idPlan"].Value == DBNull.Value) return;

			this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlazo, Descripcion, 0 As PorcComision, 0 As PorcInteres, 0 As Sel, Convert (Datetime,'20150101') as Desde ,  Convert (Datetime,'20181231') as Hasta  From Plazos Where idPlazo = 0");

			if ((bool)this.uGridPlanes.ActiveRow.Cells["Sel"].Value)
			{
				string sSQL = string.Format("Exec ConsultaMantenimientoTarjetas {0}, {1}, {2}, 2, {3}", 
					(int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value, (int)this.uGridTarjetas.ActiveRow.Cells["idTarjeta"].Value, 
					(int)this.uGridPlanes.ActiveRow.Cells["idPlan"].Value, (int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value);
				this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);			
			}

      this.uGridPlazos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Activation.AllowEdit;

			if((int)this.uGridPlanes.ActiveRow.Cells["idPlan"].Value == 1 ||(int)this.uGridPlanes.ActiveRow.Cells["idPlan"].Value == 4)
			{
				if ((bool)this.uGridPlanes.ActiveRow.Cells["Sel"].Value)
				{
					string ssSQL = string.Format("Exec ConsultaMantenimientoTarjetas {0}, {1}, {2}, 4, {3}", 
						(int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value, (int)this.uGridTarjetas.ActiveRow.Cells["idTarjeta"].Value, 
						(int)this.uGridPlanes.ActiveRow.Cells["idPlan"].Value, (int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value);
					this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, ssSQL);	
//					this.uGridPlazos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Activation.Disabled;

				}
			}
		}

		private void btnBancos_Click(object sender, System.EventArgs e)
		{
			using(Bancos miBanco = new Bancos())
			{
				if (DialogResult.OK == miBanco.ShowDialog())
				{
					//this.uGridBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Order By Nombre Asc");
				}
			}
		}

		private void btnTarjetas_Click(object sender, System.EventArgs e)
		{
			using(TarjetasDeCredito miTarjeta = new TarjetasDeCredito())
			{
				if (DialogResult.OK == miTarjeta.ShowDialog())
				{
//					string sSQL = string.Format("Exec ListaTarjetasBancos {0}", (int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value);
//					this.uGridTarjetas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
				}
			}
		}


		private void uGridRegion_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.uGridRegion.ActiveRow.Cells["idRegion"].Value == DBNull.Value) return;

			this.uGridBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre, Convert (Bit,0) as Sel From Banco Where idBanco = 0");
			this.uGridTarjetas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTarjeta, Nombre From TarjetasDeCredito Where idTarjeta = 0");
			this.uGridPlanes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlan, Descripcion From Planes Where idPlan = 0");
			this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlazo, Descripcion, 0 As PorcComision, 0 As PorcInteres, 0 As Sel, Convert (Datetime,'20150101') as Desde ,  Convert (Datetime,'20181231') as Hasta  From Plazos Where idPlazo = 0");
			
			if ((bool)this.uGridRegion.ActiveRow.Cells["Sel"].Value)
			{
		
				string sSQL = string.Format("Exec ConsultaMantenimientoTarjetas 0, 0, 0, 3, {0}", (int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value);
				this.uGridBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);		
			}
			
	
		}

		private void uGridRegion_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		 
		}

		private void uGridPlazos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
    	{
//			if((int)this.uGridPlazos.ActiveRow.Cells["idPlan"].Value == 1 ||(int)this.uGridPlanes.ActiveRow.Cells["idPlan"].Value == 4)
//			{
//				if (e.Cell.Column.ToString() == "Hasta" ||e.Cell.Column.ToString() == "Desde" )
//				{
//					if ((DateTime)e.Cell.Row.Cells["Desde"].Value > (DateTime)e.Cell.Row.Cells["Hasta"].Value )
//					{
//						MessageBox.Show("La Fecha Inicial es mayor a la Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						//					e.Cell.Row.Cells["Hasta"].Value = (int)e.Cell.OriginalValue;	
//						this.uGridPlazos.ActiveRow.Selected = true;
//						return;
//					}
//				}
//				
//				if (e.Cell.Column.ToString() == "Hasta" ||e.Cell.Column.ToString() == "Desde" || e.Cell.Column.ToString() == "PorcComision" ||e.Cell.Column.ToString() == "PorcInteres" )
//				{
//					
//					string sSQL = string.Format("Exec ActualizaMantenimientoTarjetas {0}, {1}, {2}, {3}, {4}, {5}, 2, {6}, '{7}', '{8}'", 
//						(int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value, (int)this.uGridTarjetas.ActiveRow.Cells["idTarjeta"].Value, 
//						(int)this.uGridPlanes.ActiveRow.Cells["idPlan"].Value, (int)e.Cell.Row.Cells["idPlazo"].Value, 
//						Convert.ToDecimal(e.Cell.Row.Cells["PorcComision"].Value), Convert.ToDecimal(e.Cell.Row.Cells["PorcInteres"].Value),
//						(int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).ToString("yyyyMMdd"),
//						Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd"));
//					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
//				}
			
//			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			int iCont =0;
			using(frmNuevosPlazos miPlazo = new frmNuevosPlazos())
			{

				if (DialogResult.OK == miPlazo.ShowDialog())
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPlanes.Rows.All)
					{
           iCont = iCont +1;
					}
					if(iCont == 0) return;
					if (this.uGridPlanes.ActiveRow.Cells["idPlan"].Value == DBNull.Value) return;

					this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPlazo, Descripcion, 0 As PorcComision, 0 As PorcInteres, 0 As Sel, Convert (Datetime,'20150101') as Desde ,  Convert (Datetime,'20181231') as Hasta  From Plazos Where idPlazo = 0");

					if ((bool)this.uGridPlanes.ActiveRow.Cells["Sel"].Value)
					{
						string sSQL = string.Format("Exec ConsultaMantenimientoTarjetas {0}, {1}, {2}, 2, {3}", 
							(int)this.uGridBancos.ActiveRow.Cells["idBanco"].Value, (int)this.uGridTarjetas.ActiveRow.Cells["idTarjeta"].Value, 
							(int)this.uGridPlanes.ActiveRow.Cells["idPlan"].Value, (int)this.uGridRegion.ActiveRow.Cells["idRegion"].Value);
						this.uGridPlazos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);			
					}
				}
			}
		}

		private void uGridRegion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
