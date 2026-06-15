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
	/// Descripción breve de frmDocsAseguradora.
	/// </summary>
	public class frmDocsAseguradora : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label7;
		public System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optSeleccionado;
		private System.Windows.Forms.Label lblContador;
		public System.Windows.Forms.Button btnImpVerifTerrena;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDocsAseguradora()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Plazo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FondoGarantia");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impreso");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaImpresion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFDG");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Plazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FondoGarantia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaImpresion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFDG");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDocsAseguradora));
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label6 = new System.Windows.Forms.Label();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optSeleccionado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label7 = new System.Windows.Forms.Label();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnImpVerifTerrena = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optSeleccionado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 30;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 80;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 70;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 140;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 200;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Format = "#, ##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 90;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance6;
			ultraGridColumn10.Format = "#, ##0.00";
			ultraGridColumn10.Header.Caption = "Valor Cuotas";
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Width = 90;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellAppearance = appearance7;
			ultraGridColumn11.Header.Caption = "F D G";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Width = 50;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance8;
			ultraGridColumn12.Format = "#, ##0.00";
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Width = 90;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance9;
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance10;
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.Width = 50;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn15.Header.Caption = "Fecha Impresion";
			ultraGridColumn15.Header.VisiblePosition = 13;
			ultraGridColumn15.Width = 110;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
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
																										 ultraGridColumn16});
			ultraGridBand1.SummaryFooterCaption = "Total Caja";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance11;
			this.uGridDocumentos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance13;
			this.uGridDocumentos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridDocumentos.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.None;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 8F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance14;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.uGridDocumentos.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.uGridDocumentos.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.uGridDocumentos.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.uGridDocumentos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(8, 80);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(1000, 480);
			this.uGridDocumentos.TabIndex = 283;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn1.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn4.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn15.DataType = typeof(System.DateTime);
			ultraDataColumn15.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn16.DataType = typeof(int);
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
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 706;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optEstado
			// 
			appearance17.FontData.Name = "Tahoma";
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance17;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 0;
			this.optEstado.ItemAppearance = appearance18;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 2;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Impreso";
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "Pendiente";
			this.optEstado.Items.Add(valueListItem1);
			this.optEstado.Items.Add(valueListItem2);
			this.optEstado.Items.Add(valueListItem3);
			this.optEstado.ItemSpacingVertical = 10;
			this.optEstado.Location = new System.Drawing.Point(192, 8);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(216, 24);
			this.optEstado.TabIndex = 707;
			this.optEstado.Text = "Todos";
			// 
			// dtFecha
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance19;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(120, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 705;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// optSeleccionado
			// 
			appearance20.FontData.Name = "Tahoma";
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.optSeleccionado.Appearance = appearance20;
			this.optSeleccionado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optSeleccionado.CheckedIndex = 1;
			this.optSeleccionado.ItemAppearance = appearance21;
			this.optSeleccionado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = true;
			valueListItem4.DisplayText = "Seleccionado";
			valueListItem5.DataValue = false;
			valueListItem5.DisplayText = "No Seleccionado";
			this.optSeleccionado.Items.Add(valueListItem4);
			this.optSeleccionado.Items.Add(valueListItem5);
			this.optSeleccionado.ItemSpacingVertical = 10;
			this.optSeleccionado.Location = new System.Drawing.Point(416, 8);
			this.optSeleccionado.Name = "optSeleccionado";
			this.optSeleccionado.Size = new System.Drawing.Size(216, 24);
			this.optSeleccionado.TabIndex = 708;
			this.optSeleccionado.Text = "No Seleccionado";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 40);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(40, 16);
			this.label7.TabIndex = 721;
			this.label7.Text = "Cédula";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(824, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(88, 23);
			this.btnImprimir.TabIndex = 723;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(920, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 722;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(56, 40);
			this.txtBuscar.MaxLength = 10;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(120, 19);
			this.txtBuscar.TabIndex = 720;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(192, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 719;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
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
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(312, 40);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 16);
			this.lblContador.TabIndex = 724;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnImpVerifTerrena
			// 
			this.btnImpVerifTerrena.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImpVerifTerrena.Image = ((System.Drawing.Image)(resources.GetObject("btnImpVerifTerrena.Image")));
			this.btnImpVerifTerrena.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImpVerifTerrena.Location = new System.Drawing.Point(824, 40);
			this.btnImpVerifTerrena.Name = "btnImpVerifTerrena";
			this.btnImpVerifTerrena.Size = new System.Drawing.Size(184, 23);
			this.btnImpVerifTerrena.TabIndex = 725;
			this.btnImpVerifTerrena.Text = "&Imprimir Inspeccion de Riesgo";
			this.btnImpVerifTerrena.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImpVerifTerrena.Click += new System.EventHandler(this.btnImpVerifTerrena_Click);
			// 
			// frmDocsAseguradora
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1016, 566);
			this.Controls.Add(this.btnImpVerifTerrena);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.optSeleccionado);
			this.Controls.Add(this.optEstado);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.uGridDocumentos);
			this.Name = "frmDocsAseguradora";
			this.Text = "Documentos Aseguradora";
			this.Load += new System.EventHandler(this.frmDocsAseguradora_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optSeleccionado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}


		private void frmDocsAseguradora_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200627");

			if (!Funcion.Permiso("1956", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Administración DFI", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			string carpeta_1 ="";
			carpeta_1 = @"C:\Latinium\Reportes\Cre_InspecciondeRiesgoCrediticio.rpt";
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{			
			string sSQL = string.Format("Exec Cre_ListaReclamosAseguradora '{0}', {1}, {2}", 
				this.txtBuscar.Text.ToString(), (int)this.optEstado.Value, (bool)this.optSeleccionado.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDocumentos);

			this.lblContador.Text = this.uGridDocumentos.Rows.Count.ToString() + " REGISTROS SELECCIONADOS";
		}

		private void Reportes(int idCompra, int iReporte, int iEstadoCuenta)
		{
			if (iReporte == 0)
			{
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = idCompra;
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				Reporte miRep = new Reporte("Cre_AseguradoraPagosCliente.rpt", "");
				miRep.MdiParent = this.MdiParent;			
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
				return;
			}					
							
			if (iReporte == 1)
			{	
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = idCompra;
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				ParameterField paramFieldCorte = new ParameterField ();
				ParameterDiscreteValue ValIdCorte = new ParameterDiscreteValue ();
				paramFieldCorte.ParameterFieldName = "@Corte";
				ValIdCorte.Value = Convert.ToDateTime(this.dtFecha.Value);
				paramFieldCorte.CurrentValues.Add (ValIdCorte);
				paramFields.Add (paramFieldCorte);

				Reporte miRep = new Reporte("Cre_AseguradoraComunicado.rpt", "");
				miRep.MdiParent = this.MdiParent;			
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();

				return;
			}

			if (iReporte > 1)
			{
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = idCompra;
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				ParameterField paramFieldFecha = new ParameterField ();
				ParameterDiscreteValue ValIdFecha = new ParameterDiscreteValue ();
				paramFieldFecha.ParameterFieldName = "@Fecha";
				ValIdFecha.Value = Convert.ToDateTime(this.dtFecha.Value);
				paramFieldFecha.CurrentValues.Add (ValIdFecha);
				paramFields.Add (paramFieldFecha);

				ParameterField paramFieldEC = new ParameterField ();
				ParameterDiscreteValue ValEC = new ParameterDiscreteValue ();
				paramFieldEC.ParameterFieldName = "@EstadoDeCuenta";
				ValEC.Value = iEstadoCuenta;
				paramFieldEC.CurrentValues.Add (ValEC);
				paramFields.Add (paramFieldEC);

				ParameterField paramFieldCorte = new ParameterField ();
				ParameterDiscreteValue ValIdCorte = new ParameterDiscreteValue ();
				paramFieldCorte.ParameterFieldName = "@Corte";
				ValIdCorte.Value = Convert.ToDateTime(this.dtFecha.Value);
				paramFieldCorte.CurrentValues.Add (ValIdCorte);
				paramFields.Add (paramFieldCorte);

				Reporte miRep = new Reporte("CreAseguradoraEstadoDeCuenta.rpt", "");
				miRep.MdiParent = this.MdiParent;			
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			int idCompra = int.Parse(this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value.ToString());
			int idFDG = int.Parse(this.uGridDocumentos.ActiveRow.Cells["idFDG"].Value.ToString());

			Reportes(idCompra, 0, 0);
			if (idFDG > 0) Reportes(idFDG, 0, 0);

			Reportes(idCompra, 1, 0);
			
			/* TABLA DE AMORTIZACION FIRMADA */
			if (idFDG == 0)	Reportes(idCompra, 2, 0);								
			if (idFDG > 0) Reportes(idFDG, 2, 0);

			/* ESTADO DE CUENTA */
			Reportes(idCompra, 3, 1);
			if (idFDG > 0) Reportes(idFDG, 3, 1);
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDocumentos);
		}

		private void btnImpVerifTerrena_Click(object sender, System.EventArgs e)
		{
			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
				if ((bool)dr.Cells["Sel"].Value) iCont++;

			if (iCont == 0)
			{
				MessageBox.Show("Debe seleccionar al menos un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}


			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{

				if ((bool)dr.Cells["Sel"].Value)
				{
					#region imprecion en pantalla
//								ParameterFields paramFields = new ParameterFields ();
//									
//								ParameterField paramFieldCompra = new ParameterField ();
//								ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
//								paramFieldCompra.ParameterFieldName = "@Cedula";
//								ValIdCompra.Value = dr.Cells["Cedula"].Value.ToString();
//								paramFieldCompra.CurrentValues.Add (ValIdCompra);
//								paramFields.Add (paramFieldCompra);
//					
//								ParameterField paramFieldCorte = new ParameterField ();
//								ParameterDiscreteValue ValIdCorte = new ParameterDiscreteValue ();
//								paramFieldCorte.ParameterFieldName = "@Numero";
//								ValIdCorte.Value = dr.Cells["Numero"].Value.ToString();
//								paramFieldCorte.CurrentValues.Add (ValIdCorte);
//								paramFields.Add (paramFieldCorte);
//					
//					
//								Reporte miRep = new Reporte("Cre_InspecciondeRiesgoCrediticio.rpt", "");
//								miRep.MdiParent = this.MdiParent;			
//								miRep.crVista.ParameterFieldInfo = paramFields;
//								miRep.Show();
					#endregion imprecion en pantalla

					#region Impresion directa
								string sReporte = "Cre_InspecciondeRiesgoCrediticio.rpt";
								ReportDocument Report1 = new ReportDocument();
								Report1.FileName = MenuLatinium.stDirReportes + sReporte;//"RolesIndividuales.rpt";
													
								ConnectionInfo crConnectionInfo = new ConnectionInfo(); 
							{
								string strBaseDato = MenuLatinium.stDirNumero;
								string strUsuario = MenuLatinium.stDirConta;
								string strServidor = MenuLatinium.stDirServidor;
								string strImpresora = MenuLatinium.ImpresoraNombre;
					
								crConnectionInfo.ServerName = strServidor;
								crConnectionInfo.DatabaseName = strBaseDato;
								crConnectionInfo.UserID = "infoelect";
								crConnectionInfo.Password = "Bl45o6$9";
							}										
								
								TableLogOnInfo crTableLogOnInfo; 
								Database crDatabase = Report1.Database; 
								Tables crTables = crDatabase.Tables; 
								Table crTable;
					
								for (int i=0; i< crTables.Count; i++)
								{
									{
										crTable = crTables [i]; 
										crTableLogOnInfo = crTable.LogOnInfo; 
										crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
										crTable.ApplyLogOnInfo(crTableLogOnInfo);
									}
								}
										
								Report1.PrintOptions.PrinterName = MenuLatinium.ImpresoraNombre;
								Report1.SetParameterValue("@Cedula", dr.Cells["Cedula"].Value.ToString());
								Report1.SetParameterValue("@Numero", dr.Cells["Numero"].Value.ToString());
						
								Report1.PrintToPrinter(1, false, 0, 0);

					#endregion Impresion directa

				}
			}
		}
	}
}
