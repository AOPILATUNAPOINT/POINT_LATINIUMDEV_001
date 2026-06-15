using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmConfigurarPYG.
	/// </summary>
	public class frmConfigurarPYG : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGastos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private C1.Data.C1DataSet c1DataSet1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstructura;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbGrupo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbClase;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnConfiguracion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbBodega;
		private System.Windows.Forms.Button btnExportar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConfigurarPYG()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ActivoPYG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ActivoPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstructuraPYG", -1, "cmbEstructura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoPYG", -1, "cmbGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPYG", -1, "cmbTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaPYG", -1, "cmbCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClasePYG", -1, "cmbClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega", -1, "cmbBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CC");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstructuraPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idClasePYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idClasePYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstructuraPYG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstructuraPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoPYG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPYG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaPYG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaPYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClasePYG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idClasePYG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConfigurarPYG));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.uGridGastos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.c1DataSet1 = new C1.Data.C1DataSet();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstructura = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbClase = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnConfiguracion = new System.Windows.Forms.Button();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btnExportar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataSet1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstructura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClase)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridGastos
			// 
			this.uGridGastos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridGastos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGastos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridGastos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 52;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 130;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Cuenta";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 180;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.Header.Caption = "Gastos";
			ultraGridBand1.HeaderVisible = true;
			this.uGridGastos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGastos.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridGastos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGastos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridGastos.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGastos.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridGastos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGastos.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridGastos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGastos.Location = new System.Drawing.Point(8, 40);
			this.uGridGastos.Name = "uGridGastos";
			this.uGridGastos.Size = new System.Drawing.Size(384, 352);
			this.uGridGastos.TabIndex = 4;
			this.uGridGastos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_CellChange);
			this.uGridGastos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridGastos_AfterCellUpdate);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn2.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource2;
			appearance8.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance8;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 52;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 130;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Cuenta";
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 180;
			ultraGridColumn8.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn8.Header.Caption = "Estructura";
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn8.Width = 150;
			ultraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn9.Header.Caption = "Grupo";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn9.Width = 150;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn10.Header.Caption = "Tipo";
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn10.Width = 150;
			ultraGridColumn11.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn11.Header.Caption = "Cuenta";
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn11.Width = 150;
			ultraGridColumn12.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn12.Header.Caption = "Clase";
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn12.Width = 150;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn13.Header.Caption = "Establecimiento";
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn13.Width = 150;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "Proyecto";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			ultraGridBand2.Header.Caption = "Gastos";
			ultraGridBand2.HeaderVisible = true;
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.uGridDetalle.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(400, 40);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(880, 352);
			this.uGridDetalle.TabIndex = 5;
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14});
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
			// ultraDataSource7
			// 
			ultraDataColumn15.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// c1DataSet1
			// 
			this.c1DataSet1.BindingContextCtrl = this;
			this.c1DataSet1.DataLibrary = "LibFacturacion";
			this.c1DataSet1.DataLibraryUrl = "";
			this.c1DataSet1.DataSetDef = "dsSeteoF";
			this.c1DataSet1.Locale = new System.Globalization.CultureInfo("es-EC");
			this.c1DataSet1.Name = "c1DataSet1";
			this.c1DataSet1.SchemaClassName = "LibFacturacion.DataClass";
			this.c1DataSet1.SchemaDef = null;
			// 
			// ultraDataSource8
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// cmbEstructura
			// 
			this.cmbEstructura.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstructura.DataSource = this.ultraDataSource3;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 140;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbEstructura.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEstructura.DisplayMember = "Nombre";
			this.cmbEstructura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstructura.Location = new System.Drawing.Point(48, 96);
			this.cmbEstructura.Name = "cmbEstructura";
			this.cmbEstructura.Size = new System.Drawing.Size(160, 120);
			this.cmbEstructura.TabIndex = 986;
			this.cmbEstructura.ValueMember = "idEstructuraPYG";
			this.cmbEstructura.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGrupo.DataSource = this.ultraDataSource4;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 140;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbGrupo.DisplayMember = "Nombre";
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(344, 96);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(160, 120);
			this.cmbGrupo.TabIndex = 987;
			this.cmbGrupo.ValueMember = "idGrupoPYG";
			this.cmbGrupo.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// cmbTipo
			// 
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.ultraDataSource5;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 140;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 140;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbTipo.DisplayMember = "Nombre";
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(632, 96);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(160, 120);
			this.cmbTipo.TabIndex = 988;
			this.cmbTipo.ValueMember = "idTipoPYG";
			this.cmbTipo.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource6;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 140;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCuenta.DisplayMember = "Nombre";
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(920, 96);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(152, 120);
			this.cmbCuenta.TabIndex = 989;
			this.cmbCuenta.ValueMember = "idCuentaPYG";
			this.cmbCuenta.Visible = false;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn25.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// cmbClase
			// 
			this.cmbClase.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbClase.DataSource = this.ultraDataSource9;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 140;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbClase.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbClase.DisplayMember = "Nombre";
			this.cmbClase.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbClase.Location = new System.Drawing.Point(48, 224);
			this.cmbClase.Name = "cmbClase";
			this.cmbClase.Size = new System.Drawing.Size(152, 80);
			this.cmbClase.TabIndex = 990;
			this.cmbClase.ValueMember = "idClasePYG";
			this.cmbClase.Visible = false;
			// 
			// ultraDataSource9
			// 
			ultraDataColumn27.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn27,
																																 ultraDataColumn28});
			// 
			// btnConsultar
			// 
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(400, 8);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 991;
			this.btnConsultar.Text = "&Actualizar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1200, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 992;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnConfiguracion
			// 
			this.btnConfiguracion.CausesValidation = false;
			this.btnConfiguracion.Image = ((System.Drawing.Image)(resources.GetObject("btnConfiguracion.Image")));
			this.btnConfiguracion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConfiguracion.Location = new System.Drawing.Point(1112, 8);
			this.btnConfiguracion.Name = "btnConfiguracion";
			this.btnConfiguracion.Size = new System.Drawing.Size(78, 24);
			this.btnConfiguracion.TabIndex = 993;
			this.btnConfiguracion.Text = "&Estructura";
			this.btnConfiguracion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConfiguracion.Click += new System.EventHandler(this.btnConfiguracion_Click);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn29.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn29,
																																	ultraDataColumn30});
			// 
			// cmbBodega
			// 
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource10;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 140;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(564, 143);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(160, 120);
			this.cmbBodega.TabIndex = 994;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.Visible = false;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1032, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 23);
			this.btnExportar.TabIndex = 1012;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// frmConfigurarPYG
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1288, 406);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnConfiguracion);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.cmbClase);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.cmbEstructura);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.uGridGastos);
			this.KeyPreview = true;
			this.Name = "frmConfigurarPYG";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Configuración";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConfigurarPYG_KeyDown);
			this.Load += new System.EventHandler(this.frmConfigurarPYG_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.c1DataSet1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstructura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbClase)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmConfigurarPYG_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200809");

			if (!Funcion.Permiso("1851", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar Configuración", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			string sSQLPa = string.Format(" Select ActivoPYG, idCuenta, Codigo, Descripcion From Cuenta Where Tipo = 5 And Grupo = 0 Order By Codigo ");
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridGastos);

			string sSQLPaE = string.Format(" select idCuenta, Codigo, Descripcion, idEstructuraPYG, idGrupoPYG, idTipoPYG, idCuentaPYG, idClasePYG, Bodega, '' as CC from Cuenta where ActivoPYG = 1 ");
			FuncionesProcedimientos.dsGeneral(sSQLPaE, this.uGridDetalle);

			this.cmbEstructura.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ReporteEstructuraPYG 1 "));
			this.cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ReporteEstructuraPYG 2 "));
			this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ReporteEstructuraPYG 3 "));
			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ReporteEstructuraPYG 4 "));
			this.cmbClase.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ReporteEstructuraPYG 5 "));
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select Bodega, Nombre from Bodega where Factura = 1 "));


		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void frmConfigurarPYG_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
		this.Close ();
		}

		private void uGridGastos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "ActivoPYG")
			{
				if (e.Cell.Row.Cells["idCuenta"].Value != System.DBNull.Value)
				{				
					this.uGridGastos.UpdateData();
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set ActivoPYG = {0} Where idCuenta = {1}", Convert.ToInt32( e.Cell.Row.Cells["ActivoPYG"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
				}
			}
		}

		private void uGridGastos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			string sSQLPaE = string.Format(" select idCuenta, Codigo, Descripcion, idEstructuraPYG, idGrupoPYG, idTipoPYG, idCuentaPYG, idClasePYG, Bodega, '' As CC from Cuenta where ActivoPYG = 1 ");
			FuncionesProcedimientos.dsGeneral(sSQLPaE, this.uGridDetalle);
		}
 bool bActualiza = false;
		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			#region Estructura
			if (e.Cell.Column.ToString() == "idEstructuraPYG")
			{	
				bActualiza = true;

				if ( uGridDetalle.ActiveRow.Cells["idEstructuraPYG"].Value != System.DBNull.Value )
				{								
					if ((int) uGridDetalle.ActiveRow.Cells["idCuenta"].Value > 0 )
					{				
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set idEstructuraPYG = {0} Where idCuenta = {1}", Convert.ToInt32( e.Cell.Row.Cells["idEstructuraPYG"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
					}
				}
				bActualiza = false;
			}
			#endregion Estructura

			#region Grupo
			if (e.Cell.Column.ToString() == "idGrupoPYG")
			{	
				bActualiza = true;

				if ( uGridDetalle.ActiveRow.Cells["idGrupoPYG"].Value != System.DBNull.Value )
				{								
					if ((int) uGridDetalle.ActiveRow.Cells["idCuenta"].Value > 0 )
					{				
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set idGrupoPYG = {0} Where idCuenta = {1}", Convert.ToInt32( e.Cell.Row.Cells["idGrupoPYG"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
					}
				}
				bActualiza = false;
			}
			#endregion Grupo

			#region Tipo
			if (e.Cell.Column.ToString() == "idTipoPYG")
			{	
				bActualiza = true;

				if ( uGridDetalle.ActiveRow.Cells["idTipoPYG"].Value != System.DBNull.Value )
				{								
					if ((int) uGridDetalle.ActiveRow.Cells["idCuenta"].Value > 0 )
					{				
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set idTipoPYG = {0} Where idCuenta = {1}", Convert.ToInt32( e.Cell.Row.Cells["idTipoPYG"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
					}
				}
				bActualiza = false;
			}
			#endregion Tipo

			#region Cuenta
			if (e.Cell.Column.ToString() == "idCuentaPYG")
			{	
				bActualiza = true;

				if ( uGridDetalle.ActiveRow.Cells["idCuentaPYG"].Value != System.DBNull.Value )
				{								
					if ((int) uGridDetalle.ActiveRow.Cells["idCuenta"].Value > 0 )
					{				
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set idCuentaPYG = {0} Where idCuenta = {1}", Convert.ToInt32( e.Cell.Row.Cells["idCuentaPYG"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
					}
				}
				bActualiza = false;
			}
			#endregion Cuenta

			#region Clase
			if (e.Cell.Column.ToString() == "idClasePYG")
			{	
				bActualiza = true;

				if ( uGridDetalle.ActiveRow.Cells["idClasePYG"].Value != System.DBNull.Value )
				{								
					if ((int) uGridDetalle.ActiveRow.Cells["idCuenta"].Value > 0 )
					{				
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set idClasePYG = {0} Where idCuenta = {1}", Convert.ToInt32( e.Cell.Row.Cells["idClasePYG"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
					}
				}
				bActualiza = false;
			}
			#endregion Clase

			#region Bodega
			if (e.Cell.Column.ToString() == "Bodega")
			{	
				bActualiza = true;

				if ( uGridDetalle.ActiveRow.Cells["Bodega"].Value != System.DBNull.Value )
				{								
					if ((int) uGridDetalle.ActiveRow.Cells["idCuenta"].Value > 0 )
					{				
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cuenta Set Bodega = {0} Where idCuenta = {1}", Convert.ToInt32( e.Cell.Row.Cells["Bodega"].Value), (int)e.Cell.Row.Cells["idCuenta"].Value));
					}
				}
				bActualiza = false;
			}
			#endregion Bodega
		}

		private void btnConfiguracion_Click(object sender, System.EventArgs e)
		{
			using (frmFormatoPYG CG = new frmFormatoPYG())
			{				
				if (DialogResult.OK == CG.ShowDialog())
				{
					this.cmbEstructura.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ReporteEstructuraPYG 1 "));
					this.cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ReporteEstructuraPYG 2 "));
					this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ReporteEstructuraPYG 3 "));
					this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ReporteEstructuraPYG 4 "));
					this.cmbClase.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ReporteEstructuraPYG 5 "));

				}
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			using (frmPoryectoGre Viaticos = new frmPoryectoGre( ))
			{
				if (DialogResult.OK == Viaticos.ShowDialog( ))
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Viaticos.uGridGastos.Rows.All)
					{
						if (dr.Cells["Sel"].Value != System.DBNull.Value)
						{
							if ((bool)dr.Cells["Sel"].Value == true)
							{
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dra in this.uGridDetalle.Rows.All)
								{	
									dra.Cells["CC"].Value ="";
                  dra.Cells["CC"].Value = Convert.ToString( dr.Cells["Nombre"].Value);
								}
								FuncionesProcedimientos.ExportaExcel(this.uGridDetalle);
							}
						}
					}
				}
			}
		}
	}
}
