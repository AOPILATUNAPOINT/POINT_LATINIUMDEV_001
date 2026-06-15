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
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteRangoMorosidad.
	/// </summary>
	public class frmReporteRangoMorosidad : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocal;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkLocal;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private System.Windows.Forms.Label label8;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbEntidadFinanciera;
		private System.Windows.Forms.Label lblContador;
		public System.Windows.Forms.Button btnConsultar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtcero;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtuno;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtdos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txttres;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtcuatro;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtcinco;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtseis;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtsiete;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje1;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje7;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optRegion;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optPOrcentajes;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRatioM;
		public Infragistics.Win.UltraWinGrid.UltraGrid ugridcolores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje8;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtOcho;
		private System.Windows.Forms.Label label14;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteRangoMorosidad()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteRangoMorosidad));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Verde");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Amarillo");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Naranja");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rojo");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Verde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Amarillo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Naranja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rojo");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			this.label4 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridLocal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkLocal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnExcel = new System.Windows.Forms.Button();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label8 = new System.Windows.Forms.Label();
			this.cmbEntidadFinanciera = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtcero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtuno = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtdos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txttres = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtcuatro = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtcinco = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtseis = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtsiete = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.txtPorcentaje0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcentaje1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcentaje2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcentaje3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcentaje4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcentaje5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcentaje6 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcentaje7 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.optRegion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optPOrcentajes = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label13 = new System.Windows.Forms.Label();
			this.txtRatioM = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ugridcolores = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtPorcentaje8 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtOcho = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtcero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtuno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txttres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtcuatro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtcinco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtseis)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtsiete)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optPOrcentajes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRatioM)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridcolores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOcho)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 13);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(36, 16);
			this.label4.TabIndex = 744;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 11);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 745;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// uGridLocal
			// 
			this.uGridLocal.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocal.DataSource = this.ultraDataSource6;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridLocal.DisplayLayout.Appearance = appearance2;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.Caption = "Local";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 150;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 28;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand1.SummaryFooterCaption = "TOTALES";
			this.uGridLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridLocal.DisplayLayout.GroupByBox.Hidden = true;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocal.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.uGridLocal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLocal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocal.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocal.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.uGridLocal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridLocal.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridLocal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocal.Location = new System.Drawing.Point(3, 168);
			this.uGridLocal.Name = "uGridLocal";
			this.uGridLocal.Size = new System.Drawing.Size(213, 296);
			this.uGridLocal.TabIndex = 986;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// chkLocal
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkLocal.Appearance = appearance9;
			this.chkLocal.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkLocal.Checked = true;
			this.chkLocal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkLocal.Location = new System.Drawing.Point(184, 168);
			this.chkLocal.Name = "chkLocal";
			this.chkLocal.Size = new System.Drawing.Size(13, 16);
			this.chkLocal.TabIndex = 987;
			this.chkLocal.CheckedChanged += new System.EventHandler(this.chkLocal_CheckedChanged);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1264, 9);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(67, 24);
			this.btnExcel.TabIndex = 994;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance10.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance10;
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(224, 136);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1112, 328);
			this.uGridInformacion.TabIndex = 995;
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(168, 13);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(36, 16);
			this.label8.TabIndex = 999;
			this.label8.Text = "Banco";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEntidadFinanciera
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEntidadFinanciera.Appearance = appearance16;
			this.cmbEntidadFinanciera.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntidadFinanciera.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 247;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbEntidadFinanciera.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEntidadFinanciera.DisplayMember = "Nombre";
			this.cmbEntidadFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntidadFinanciera.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntidadFinanciera.Location = new System.Drawing.Point(216, 11);
			this.cmbEntidadFinanciera.Name = "cmbEntidadFinanciera";
			this.cmbEntidadFinanciera.Size = new System.Drawing.Size(248, 21);
			this.cmbEntidadFinanciera.TabIndex = 998;
			this.cmbEntidadFinanciera.ValueMember = "idEntidadFinanciera";
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(832, 13);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 997;
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(736, 10);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 23);
			this.btnConsultar.TabIndex = 996;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
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
			// label1
			// 
			this.label1.BackColor = System.Drawing.Color.DarkGray;
			this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(96, 48);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(104, 16);
			this.label1.TabIndex = 1000;
			this.label1.Text = "0-0";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label2
			// 
			this.label2.BackColor = System.Drawing.Color.DarkGray;
			this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.ForeColor = System.Drawing.Color.Firebrick;
			this.label2.Location = new System.Drawing.Point(208, 48);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(104, 16);
			this.label2.TabIndex = 1001;
			this.label2.Text = "1- 5";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.BackColor = System.Drawing.Color.DarkGray;
			this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Firebrick;
			this.label3.Location = new System.Drawing.Point(320, 48);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(104, 16);
			this.label3.TabIndex = 1002;
			this.label3.Text = "6-15";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(-57, 33);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1544, 8);
			this.groupBox2.TabIndex = 1008;
			this.groupBox2.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-72, 120);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1544, 8);
			this.groupBox1.TabIndex = 1009;
			this.groupBox1.TabStop = false;
			// 
			// txtcero
			// 
			appearance17.BackColor = System.Drawing.Color.Transparent;
			appearance17.BackColor2 = System.Drawing.Color.Transparent;
			appearance17.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance17.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtcero.Appearance = appearance17;
			this.txtcero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtcero.Enabled = false;
			this.txtcero.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtcero.Location = new System.Drawing.Point(96, 64);
			this.txtcero.MaxValue = ((long)(100000000000));
			this.txtcero.MinValue = 0;
			this.txtcero.Name = "txtcero";
			this.txtcero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtcero.PromptChar = ' ';
			this.txtcero.Size = new System.Drawing.Size(104, 23);
			this.txtcero.TabIndex = 1010;
			// 
			// txtuno
			// 
			appearance18.BackColor = System.Drawing.Color.Transparent;
			appearance18.BackColor2 = System.Drawing.Color.Transparent;
			appearance18.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance18.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtuno.Appearance = appearance18;
			this.txtuno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtuno.Enabled = false;
			this.txtuno.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtuno.Location = new System.Drawing.Point(208, 64);
			this.txtuno.MaxValue = ((long)(100000000000));
			this.txtuno.MinValue = 0;
			this.txtuno.Name = "txtuno";
			this.txtuno.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtuno.PromptChar = ' ';
			this.txtuno.Size = new System.Drawing.Size(104, 23);
			this.txtuno.TabIndex = 1011;
			// 
			// txtdos
			// 
			appearance19.BackColor = System.Drawing.Color.Transparent;
			appearance19.BackColor2 = System.Drawing.Color.Transparent;
			appearance19.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance19.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdos.Appearance = appearance19;
			this.txtdos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdos.Enabled = false;
			this.txtdos.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtdos.Location = new System.Drawing.Point(320, 64);
			this.txtdos.MaxValue = ((long)(100000000000));
			this.txtdos.MinValue = 0;
			this.txtdos.Name = "txtdos";
			this.txtdos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtdos.PromptChar = ' ';
			this.txtdos.Size = new System.Drawing.Size(104, 23);
			this.txtdos.TabIndex = 1012;
			// 
			// txttres
			// 
			appearance20.BackColor = System.Drawing.Color.Transparent;
			appearance20.BackColor2 = System.Drawing.Color.Transparent;
			appearance20.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance20.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txttres.Appearance = appearance20;
			this.txttres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txttres.Enabled = false;
			this.txttres.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txttres.Location = new System.Drawing.Point(432, 64);
			this.txttres.MaxValue = ((long)(100000000000));
			this.txttres.MinValue = 0;
			this.txttres.Name = "txttres";
			this.txttres.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txttres.PromptChar = ' ';
			this.txttres.Size = new System.Drawing.Size(104, 23);
			this.txttres.TabIndex = 1013;
			// 
			// txtcuatro
			// 
			appearance21.BackColor = System.Drawing.Color.Transparent;
			appearance21.BackColor2 = System.Drawing.Color.Transparent;
			appearance21.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance21.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance21.ForeColor = System.Drawing.Color.Black;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtcuatro.Appearance = appearance21;
			this.txtcuatro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtcuatro.Enabled = false;
			this.txtcuatro.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtcuatro.Location = new System.Drawing.Point(552, 64);
			this.txtcuatro.MaxValue = ((long)(100000000000));
			this.txtcuatro.MinValue = 0;
			this.txtcuatro.Name = "txtcuatro";
			this.txtcuatro.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtcuatro.PromptChar = ' ';
			this.txtcuatro.Size = new System.Drawing.Size(104, 23);
			this.txtcuatro.TabIndex = 1014;
			// 
			// txtcinco
			// 
			appearance22.BackColor = System.Drawing.Color.Transparent;
			appearance22.BackColor2 = System.Drawing.Color.Transparent;
			appearance22.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance22.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtcinco.Appearance = appearance22;
			this.txtcinco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtcinco.Enabled = false;
			this.txtcinco.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtcinco.Location = new System.Drawing.Point(664, 64);
			this.txtcinco.MaxValue = ((long)(100000000000));
			this.txtcinco.MinValue = 0;
			this.txtcinco.Name = "txtcinco";
			this.txtcinco.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtcinco.PromptChar = ' ';
			this.txtcinco.Size = new System.Drawing.Size(104, 23);
			this.txtcinco.TabIndex = 1015;
			this.txtcinco.ValueChanged += new System.EventHandler(this.ultraNumericEditor5_ValueChanged);
			// 
			// txtseis
			// 
			appearance23.BackColor = System.Drawing.Color.Transparent;
			appearance23.BackColor2 = System.Drawing.Color.Transparent;
			appearance23.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance23.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtseis.Appearance = appearance23;
			this.txtseis.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtseis.Enabled = false;
			this.txtseis.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtseis.Location = new System.Drawing.Point(776, 64);
			this.txtseis.MaxValue = ((long)(100000000000));
			this.txtseis.MinValue = 0;
			this.txtseis.Name = "txtseis";
			this.txtseis.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtseis.PromptChar = ' ';
			this.txtseis.Size = new System.Drawing.Size(104, 23);
			this.txtseis.TabIndex = 1016;
			// 
			// txtsiete
			// 
			appearance24.BackColor = System.Drawing.Color.Transparent;
			appearance24.BackColor2 = System.Drawing.Color.Transparent;
			appearance24.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance24.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtsiete.Appearance = appearance24;
			this.txtsiete.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtsiete.Enabled = false;
			this.txtsiete.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtsiete.Location = new System.Drawing.Point(1000, 64);
			this.txtsiete.MaxValue = ((long)(100000000000));
			this.txtsiete.MinValue = 0;
			this.txtsiete.Name = "txtsiete";
			this.txtsiete.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtsiete.PromptChar = ' ';
			this.txtsiete.Size = new System.Drawing.Size(104, 23);
			this.txtsiete.TabIndex = 1017;
			// 
			// label5
			// 
			this.label5.BackColor = System.Drawing.Color.DarkGray;
			this.label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(432, 48);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(104, 16);
			this.label5.TabIndex = 1003;
			this.label5.Text = "16 - 30";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label6
			// 
			this.label6.BackColor = System.Drawing.Color.DarkGray;
			this.label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.ForeColor = System.Drawing.Color.Firebrick;
			this.label6.Location = new System.Drawing.Point(552, 48);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(104, 16);
			this.label6.TabIndex = 1004;
			this.label6.Text = "31 - 60";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label7
			// 
			this.label7.BackColor = System.Drawing.Color.DarkGray;
			this.label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.ForeColor = System.Drawing.Color.Firebrick;
			this.label7.Location = new System.Drawing.Point(664, 48);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(104, 16);
			this.label7.TabIndex = 1005;
			this.label7.Text = "61 - 90";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label9
			// 
			this.label9.BackColor = System.Drawing.Color.DarkGray;
			this.label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.Firebrick;
			this.label9.Location = new System.Drawing.Point(776, 48);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(104, 16);
			this.label9.TabIndex = 1006;
			this.label9.Text = "91 - 99";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label10
			// 
			this.label10.BackColor = System.Drawing.Color.DarkGray;
			this.label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.ForeColor = System.Drawing.Color.Firebrick;
			this.label10.Location = new System.Drawing.Point(1000, 48);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(104, 16);
			this.label10.TabIndex = 1007;
			this.label10.Text = "TOTAL";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label11
			// 
			this.label11.BackColor = System.Drawing.Color.DarkGray;
			this.label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.ForeColor = System.Drawing.Color.Firebrick;
			this.label11.Location = new System.Drawing.Point(0, 68);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(88, 16);
			this.label11.TabIndex = 1018;
			this.label11.Text = "Sumatoria";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtPorcentaje0
			// 
			appearance25.BackColor = System.Drawing.Color.Transparent;
			appearance25.BackColor2 = System.Drawing.Color.Transparent;
			appearance25.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance25.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentaje0.Appearance = appearance25;
			this.txtPorcentaje0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje0.Enabled = false;
			this.txtPorcentaje0.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje0.Location = new System.Drawing.Point(96, 88);
			this.txtPorcentaje0.MaxValue = ((long)(100000000000));
			this.txtPorcentaje0.MinValue = 0;
			this.txtPorcentaje0.Name = "txtPorcentaje0";
			this.txtPorcentaje0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje0.PromptChar = ' ';
			this.txtPorcentaje0.Size = new System.Drawing.Size(104, 23);
			this.txtPorcentaje0.TabIndex = 1019;
			// 
			// txtPorcentaje1
			// 
			appearance26.BackColor = System.Drawing.Color.Transparent;
			appearance26.BackColor2 = System.Drawing.Color.Transparent;
			appearance26.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance26.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentaje1.Appearance = appearance26;
			this.txtPorcentaje1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje1.Enabled = false;
			this.txtPorcentaje1.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje1.Location = new System.Drawing.Point(208, 88);
			this.txtPorcentaje1.MaxValue = ((long)(100000000000));
			this.txtPorcentaje1.MinValue = 0;
			this.txtPorcentaje1.Name = "txtPorcentaje1";
			this.txtPorcentaje1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje1.PromptChar = ' ';
			this.txtPorcentaje1.Size = new System.Drawing.Size(104, 23);
			this.txtPorcentaje1.TabIndex = 1020;
			// 
			// txtPorcentaje2
			// 
			appearance27.BackColor = System.Drawing.Color.Transparent;
			appearance27.BackColor2 = System.Drawing.Color.Transparent;
			appearance27.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance27.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentaje2.Appearance = appearance27;
			this.txtPorcentaje2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje2.Enabled = false;
			this.txtPorcentaje2.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje2.Location = new System.Drawing.Point(320, 88);
			this.txtPorcentaje2.MaxValue = ((long)(100000000000));
			this.txtPorcentaje2.MinValue = 0;
			this.txtPorcentaje2.Name = "txtPorcentaje2";
			this.txtPorcentaje2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje2.PromptChar = ' ';
			this.txtPorcentaje2.Size = new System.Drawing.Size(104, 23);
			this.txtPorcentaje2.TabIndex = 1021;
			this.txtPorcentaje2.ValueChanged += new System.EventHandler(this.ultraNumericEditor3_ValueChanged);
			// 
			// txtPorcentaje3
			// 
			appearance28.BackColor = System.Drawing.Color.Transparent;
			appearance28.BackColor2 = System.Drawing.Color.Transparent;
			appearance28.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance28.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentaje3.Appearance = appearance28;
			this.txtPorcentaje3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje3.Enabled = false;
			this.txtPorcentaje3.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje3.Location = new System.Drawing.Point(432, 88);
			this.txtPorcentaje3.MaxValue = ((long)(100000000000));
			this.txtPorcentaje3.MinValue = 0;
			this.txtPorcentaje3.Name = "txtPorcentaje3";
			this.txtPorcentaje3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje3.PromptChar = ' ';
			this.txtPorcentaje3.Size = new System.Drawing.Size(104, 23);
			this.txtPorcentaje3.TabIndex = 1022;
			// 
			// txtPorcentaje4
			// 
			appearance29.BackColor = System.Drawing.Color.Transparent;
			appearance29.BackColor2 = System.Drawing.Color.Transparent;
			appearance29.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance29.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentaje4.Appearance = appearance29;
			this.txtPorcentaje4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje4.Enabled = false;
			this.txtPorcentaje4.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje4.Location = new System.Drawing.Point(552, 88);
			this.txtPorcentaje4.MaxValue = ((long)(100000000000));
			this.txtPorcentaje4.MinValue = 0;
			this.txtPorcentaje4.Name = "txtPorcentaje4";
			this.txtPorcentaje4.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje4.PromptChar = ' ';
			this.txtPorcentaje4.Size = new System.Drawing.Size(104, 23);
			this.txtPorcentaje4.TabIndex = 1023;
			// 
			// txtPorcentaje5
			// 
			appearance30.BackColor = System.Drawing.Color.Transparent;
			appearance30.BackColor2 = System.Drawing.Color.Transparent;
			appearance30.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance30.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentaje5.Appearance = appearance30;
			this.txtPorcentaje5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje5.Enabled = false;
			this.txtPorcentaje5.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje5.Location = new System.Drawing.Point(664, 88);
			this.txtPorcentaje5.MaxValue = ((long)(100000000000));
			this.txtPorcentaje5.MinValue = 0;
			this.txtPorcentaje5.Name = "txtPorcentaje5";
			this.txtPorcentaje5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje5.PromptChar = ' ';
			this.txtPorcentaje5.Size = new System.Drawing.Size(104, 23);
			this.txtPorcentaje5.TabIndex = 1024;
			// 
			// txtPorcentaje6
			// 
			appearance31.BackColor = System.Drawing.Color.Transparent;
			appearance31.BackColor2 = System.Drawing.Color.Transparent;
			appearance31.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance31.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance31.ForeColor = System.Drawing.Color.Black;
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentaje6.Appearance = appearance31;
			this.txtPorcentaje6.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje6.Enabled = false;
			this.txtPorcentaje6.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje6.Location = new System.Drawing.Point(776, 88);
			this.txtPorcentaje6.MaxValue = ((long)(100000000000));
			this.txtPorcentaje6.MinValue = 0;
			this.txtPorcentaje6.Name = "txtPorcentaje6";
			this.txtPorcentaje6.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje6.PromptChar = ' ';
			this.txtPorcentaje6.Size = new System.Drawing.Size(104, 23);
			this.txtPorcentaje6.TabIndex = 1025;
			// 
			// txtPorcentaje7
			// 
			appearance32.BackColor = System.Drawing.Color.Transparent;
			appearance32.BackColor2 = System.Drawing.Color.Transparent;
			appearance32.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance32.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance32.ForeColor = System.Drawing.Color.Black;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentaje7.Appearance = appearance32;
			this.txtPorcentaje7.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje7.Enabled = false;
			this.txtPorcentaje7.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje7.Location = new System.Drawing.Point(1000, 88);
			this.txtPorcentaje7.MaxValue = ((long)(100000000000));
			this.txtPorcentaje7.MinValue = 0;
			this.txtPorcentaje7.Name = "txtPorcentaje7";
			this.txtPorcentaje7.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje7.PromptChar = ' ';
			this.txtPorcentaje7.Size = new System.Drawing.Size(104, 23);
			this.txtPorcentaje7.TabIndex = 1026;
			// 
			// label12
			// 
			this.label12.BackColor = System.Drawing.Color.DarkGray;
			this.label12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.ForeColor = System.Drawing.Color.Firebrick;
			this.label12.Location = new System.Drawing.Point(0, 92);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(88, 16);
			this.label12.TabIndex = 1027;
			this.label12.Text = "Porcentaje";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// optRegion
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.optRegion.Appearance = appearance33;
			this.optRegion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optRegion.CheckedIndex = 0;
			this.optRegion.ItemAppearance = appearance34;
			this.optRegion.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Sierra";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Costa";
			this.optRegion.Items.Add(valueListItem1);
			this.optRegion.Items.Add(valueListItem2);
			this.optRegion.Items.Add(valueListItem3);
			this.optRegion.ItemSpacingVertical = 10;
			this.optRegion.Location = new System.Drawing.Point(16, 136);
			this.optRegion.Name = "optRegion";
			this.optRegion.Size = new System.Drawing.Size(192, 26);
			this.optRegion.TabIndex = 1028;
			this.optRegion.Text = "Todos";
			this.optRegion.ValueChanged += new System.EventHandler(this.optRegion_ValueChanged);
			// 
			// optPOrcentajes
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.optPOrcentajes.Appearance = appearance35;
			this.optPOrcentajes.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optPOrcentajes.CheckedIndex = 0;
			this.optPOrcentajes.ItemAppearance = appearance36;
			this.optPOrcentajes.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "General";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Sumatoria";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Porcentaje";
			this.optPOrcentajes.Items.Add(valueListItem4);
			this.optPOrcentajes.Items.Add(valueListItem5);
			this.optPOrcentajes.Items.Add(valueListItem6);
			this.optPOrcentajes.ItemSpacingVertical = 10;
			this.optPOrcentajes.Location = new System.Drawing.Point(472, 8);
			this.optPOrcentajes.Name = "optPOrcentajes";
			this.optPOrcentajes.Size = new System.Drawing.Size(248, 26);
			this.optPOrcentajes.TabIndex = 1029;
			this.optPOrcentajes.Text = "General";
			this.optPOrcentajes.ValueChanged += new System.EventHandler(this.optPOrcentajes_ValueChanged);
			// 
			// label13
			// 
			this.label13.BackColor = System.Drawing.Color.DarkGray;
			this.label13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.ForeColor = System.Drawing.Color.Firebrick;
			this.label13.Location = new System.Drawing.Point(1120, 48);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(104, 16);
			this.label13.TabIndex = 1030;
			this.label13.Text = "RATIO  M";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtRatioM
			// 
			appearance37.BackColor = System.Drawing.Color.Transparent;
			appearance37.BackColor2 = System.Drawing.Color.Transparent;
			appearance37.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance37.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance37.ForeColor = System.Drawing.Color.Black;
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRatioM.Appearance = appearance37;
			this.txtRatioM.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRatioM.Enabled = false;
			this.txtRatioM.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtRatioM.Location = new System.Drawing.Point(1120, 88);
			this.txtRatioM.MaxValue = ((long)(100000000000));
			this.txtRatioM.MinValue = 0;
			this.txtRatioM.Name = "txtRatioM";
			this.txtRatioM.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRatioM.PromptChar = ' ';
			this.txtRatioM.Size = new System.Drawing.Size(104, 23);
			this.txtRatioM.TabIndex = 1031;
			// 
			// ugridcolores
			// 
			this.ugridcolores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugridcolores.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugridcolores.DataSource = this.ultraDataSource1;
			appearance38.BackColor = System.Drawing.Color.White;
			this.ugridcolores.DisplayLayout.Appearance = appearance38;
			ultraGridBand3.AddButtonCaption = "Serie";
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Width = 60;
			appearance39.BackColor = System.Drawing.Color.Green;
			ultraGridColumn8.CellAppearance = appearance39;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 40;
			appearance40.BackColor = System.Drawing.Color.Yellow;
			ultraGridColumn9.CellAppearance = appearance40;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 50;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(128)), ((System.Byte)(0)));
			ultraGridColumn10.CellAppearance = appearance41;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Width = 45;
			appearance42.BackColor = System.Drawing.Color.Red;
			ultraGridColumn11.CellAppearance = appearance42;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Width = 55;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.SummaryFooterCaption = "";
			this.ugridcolores.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance43.ForeColor = System.Drawing.Color.Black;
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridcolores.DisplayLayout.Override.ActiveRowAppearance = appearance43;
			this.ugridcolores.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugridcolores.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugridcolores.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance44.BackColor = System.Drawing.Color.Transparent;
			this.ugridcolores.DisplayLayout.Override.CardAreaAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance45.FontData.BoldAsString = "True";
			appearance45.FontData.Name = "Arial";
			appearance45.FontData.SizeInPoints = 7.75F;
			appearance45.ForeColor = System.Drawing.Color.White;
			appearance45.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugridcolores.DisplayLayout.Override.HeaderAppearance = appearance45;
			this.ugridcolores.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridcolores.DisplayLayout.Override.RowAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridcolores.DisplayLayout.Override.RowSelectorAppearance = appearance47;
			this.ugridcolores.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugridcolores.Location = new System.Drawing.Point(1064, 136);
			this.ugridcolores.Name = "ugridcolores";
			this.ugridcolores.Size = new System.Drawing.Size(272, 176);
			this.ugridcolores.TabIndex = 1032;
			this.ugridcolores.Visible = false;
			// 
			// ultraDataSource1
			// 
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// txtPorcentaje8
			// 
			appearance48.BackColor = System.Drawing.Color.Transparent;
			appearance48.BackColor2 = System.Drawing.Color.Transparent;
			appearance48.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance48.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance48.ForeColor = System.Drawing.Color.Black;
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentaje8.Appearance = appearance48;
			this.txtPorcentaje8.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje8.Enabled = false;
			this.txtPorcentaje8.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje8.Location = new System.Drawing.Point(888, 88);
			this.txtPorcentaje8.MaxValue = ((long)(100000000000));
			this.txtPorcentaje8.MinValue = 0;
			this.txtPorcentaje8.Name = "txtPorcentaje8";
			this.txtPorcentaje8.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje8.PromptChar = ' ';
			this.txtPorcentaje8.Size = new System.Drawing.Size(104, 23);
			this.txtPorcentaje8.TabIndex = 1035;
			// 
			// txtOcho
			// 
			appearance49.BackColor = System.Drawing.Color.Transparent;
			appearance49.BackColor2 = System.Drawing.Color.Transparent;
			appearance49.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance49.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance49.ForeColor = System.Drawing.Color.Black;
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOcho.Appearance = appearance49;
			this.txtOcho.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOcho.Enabled = false;
			this.txtOcho.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtOcho.Location = new System.Drawing.Point(888, 64);
			this.txtOcho.MaxValue = ((long)(100000000000));
			this.txtOcho.MinValue = 0;
			this.txtOcho.Name = "txtOcho";
			this.txtOcho.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtOcho.PromptChar = ' ';
			this.txtOcho.Size = new System.Drawing.Size(104, 23);
			this.txtOcho.TabIndex = 1034;
			// 
			// label14
			// 
			this.label14.BackColor = System.Drawing.Color.DarkGray;
			this.label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.ForeColor = System.Drawing.Color.Firebrick;
			this.label14.Location = new System.Drawing.Point(888, 48);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(104, 16);
			this.label14.TabIndex = 1033;
			this.label14.Text = "100 - 180";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmReporteRangoMorosidad
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1344, 470);
			this.Controls.Add(this.txtPorcentaje8);
			this.Controls.Add(this.txtOcho);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.ugridcolores);
			this.Controls.Add(this.txtRatioM);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.optPOrcentajes);
			this.Controls.Add(this.optRegion);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtPorcentaje7);
			this.Controls.Add(this.txtPorcentaje6);
			this.Controls.Add(this.txtPorcentaje5);
			this.Controls.Add(this.txtPorcentaje4);
			this.Controls.Add(this.txtPorcentaje3);
			this.Controls.Add(this.txtPorcentaje2);
			this.Controls.Add(this.txtPorcentaje1);
			this.Controls.Add(this.txtPorcentaje0);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.txtsiete);
			this.Controls.Add(this.txtseis);
			this.Controls.Add(this.txtcinco);
			this.Controls.Add(this.txtcuatro);
			this.Controls.Add(this.txttres);
			this.Controls.Add(this.txtdos);
			this.Controls.Add(this.txtuno);
			this.Controls.Add(this.txtcero);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.cmbEntidadFinanciera);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.chkLocal);
			this.Controls.Add(this.uGridLocal);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.Name = "frmReporteRangoMorosidad";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Rango Morosidad";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteRangoMorosidad_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteRangoMorosidad_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtcero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtuno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txttres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtcuatro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtcinco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtseis)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtsiete)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optPOrcentajes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRatioM)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridcolores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOcho)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void chkLocal_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
			{
				dr.Cells["Sel"].Value = this.chkLocal.Checked;
			}
		}
		private void frmReporteRangoMorosidad_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200602");

			if (!Funcion.Permiso("1228", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Rango Morosidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar) this.btnExcel.Enabled =  true;

	
			this.dtFecha.Value = DateTime.Today ;

			string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegion.Value);		
			
			this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);


//			this.uGridLocal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre as Local , Convert(Bit, 1) As Sel, Region From Bodega Where Activo = 1 AND Factura = 1 Order By Local");
			this.cmbEntidadFinanciera.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera  where idEntidadFinanciera not in (6,2) Order By idEntidadFinanciera");
			this.ugridcolores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select Rango  , Verde  , Amarillo  ,Naranja  , Rojo  from RangoColoresMorosidad");
			this.ugridcolores.Visible = false;


		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
		  FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.txtcero.Value = 0.00m;;
			this.txtuno.Value = 0.00m;;
			this.txtdos.Value = 0.00m;;
			this.txttres.Value = 0.00m;;
			this.txtcuatro.Value = 0.00m;;
			this.txtcinco.Value = 0.00m;;
		  this.txtseis.Value = 0.00m;;
			this.txtsiete.Value = 0.00m;;
			this.txtOcho.Value = 0.00m;;
			this.txtPorcentaje8.Value = 0.00m;;


			this.txtPorcentaje0.Value = 0.00m;;
		  this.txtPorcentaje1.Value = 0.00m;;
			this.txtPorcentaje2.Value = 0.00m;;
			this.txtPorcentaje3.Value = 0.00m;;
			this.txtPorcentaje4.Value = 0.00m;;
			this.txtPorcentaje5.Value = 0.00m;;
			this.txtPorcentaje6.Value = 0.00m;;
			this.txtPorcentaje7.Value = 0.00m;;

			this.uGridInformacion.DataSource = null;

			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			#region Valida Local

			string sLocal= "";
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocal.Length == 0) sLocal = dr.Cells["Bodega"].Value.ToString();
					else if (sLocal.Length > 0) sLocal = sLocal + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			if (sLocal.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Local ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			#endregion Valida Local

			int idEntidadFinanciera  = 0;

			if (this.cmbEntidadFinanciera.ActiveRow != null) idEntidadFinanciera = (int)this.cmbEntidadFinanciera.Value;


				string sSQLPa = string.Format("Exec ReporteReporteRangoMorosidad '{0}', {1}, '{2}', {3}", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), idEntidadFinanciera, sLocal , (int) this.optPOrcentajes.Value);
				this.uGridInformacion.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
			
			
			
			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";


			#region Consulta Individual
			string ssql= string.Format ("Exec ReporteReporteRangoMorosidad '{0}', {1}, '{2}', {3} ", 
				     Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), idEntidadFinanciera, sLocal , 3); 

			SqlDataReader dra = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dra.Read();
			if (dra.HasRows)
			{	
			
				if (dra.GetValue(0) != System.DBNull.Value)this.txtcero.Value = dra.GetDecimal(0);
				if (dra.GetValue(1) != System.DBNull.Value)this.txtuno.Value = dra.GetDecimal(1);
				if (dra.GetValue(2) != System.DBNull.Value)this.txtdos.Value = dra.GetDecimal(2);
				if (dra.GetValue(3) != System.DBNull.Value)this.txttres.Value = dra.GetDecimal(3);
				if (dra.GetValue(4) != System.DBNull.Value)this.txtcuatro.Value = dra.GetDecimal(4);
				if (dra.GetValue(5) != System.DBNull.Value)this.txtcinco.Value = dra.GetDecimal(5);
				if (dra.GetValue(6) != System.DBNull.Value)this.txtseis.Value = dra.GetDecimal(6);
				if (dra.GetValue(7) != System.DBNull.Value)this.txtOcho.Value = dra.GetDecimal(7);
				if (dra.GetValue(8) != System.DBNull.Value)this.txtsiete.Value = dra.GetDecimal(8);
			}
			dra.Close();
			#endregion Consulta Individual
      
			if((decimal)this.txtsiete.Value > 0)
			{
				this.txtPorcentaje0.Value = (((decimal) this.txtcero.Value * 100)/(decimal)this.txtsiete.Value);
				this.txtPorcentaje1.Value = (((decimal) this.txtuno.Value * 100)/(decimal)this.txtsiete.Value);
				this.txtPorcentaje2.Value = (((decimal) this.txtdos.Value * 100)/(decimal)this.txtsiete.Value);
				this.txtPorcentaje3.Value = (((decimal) this.txttres.Value * 100)/(decimal)this.txtsiete.Value);
				this.txtPorcentaje4.Value = (((decimal) this.txtcuatro.Value * 100)/(decimal)this.txtsiete.Value);
				this.txtPorcentaje5.Value = (((decimal) this.txtcinco.Value * 100)/(decimal)this.txtsiete.Value);
				this.txtPorcentaje6.Value = (((decimal) this.txtseis.Value * 100)/(decimal)this.txtsiete.Value);
				this.txtPorcentaje7.Value = (((decimal) this.txtsiete.Value * 100)/(decimal)this.txtsiete.Value);
				this.txtPorcentaje8.Value = (((decimal) this.txtOcho.Value * 100)/(decimal)this.txtsiete.Value);

			}
	
			if ( (int) this.optPOrcentajes.Value == 0)
			{
				this.DiseñoGridTurno(this.uGridInformacion);
				this.ugridcolores.Visible = false;
								this.uGridInformacion.Size = new Size(1112, 386);
		
			}
			if ( (int) this.optPOrcentajes.Value == 2)
			{
				this.DiseñoGridColor(this.uGridInformacion);
					this.ugridcolores.Visible = false;
					this.DiseñoGridRatioM(this.uGridInformacion);

			}
			if ( (int) this.optPOrcentajes.Value == 1)
			{
				this.DiseñoGridSumatoria(this.uGridInformacion);
					this.ugridcolores.Visible = false;
							this.uGridInformacion.Size = new Size(1112, 386);
			}
			this.txtRatioM.Value = 		((decimal) this.txtPorcentaje3.Value +(decimal) this.txtPorcentaje4.Value +(decimal) this.txtPorcentaje5.Value +	(decimal) this.txtPorcentaje6.Value +(decimal) this.txtPorcentaje8.Value );
			

		}
		private void DiseñoGridRatioM(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
			{
				if (dr.Cells["16-30"].Value.ToString().Trim().Length == 0)  dr.Cells["16-30"].Value = 0;
				if (dr.Cells["31-60"].Value.ToString().Trim().Length == 0) dr.Cells["31-60"].Value = 0;
				if (dr.Cells["61-90"].Value.ToString().Trim().Length == 0) dr.Cells["61-90"].Value = 0;
				if (dr.Cells["91-99"].Value.ToString().Trim().Length == 0)  dr.Cells["91-99"].Value = 0;
				if (dr.Cells["100-180"].Value.ToString().Trim().Length == 0)  dr.Cells["100-180"].Value = 0;

		            dr.Cells["Ratio M"].Value =  Convert.ToDecimal(dr.Cells["16-30"].Value) +  Convert.ToDecimal(dr.Cells["31-60"].Value)  + 
					      Convert.ToDecimal(dr.Cells["61-90"].Value) +Convert.ToDecimal(dr.Cells["91-99"].Value) + Convert.ToDecimal(dr.Cells["100-180"].Value) ;
			
			}
		}
  
		private void DiseñoGridSumatoria(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["Bodega"].Hidden = true;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["Almacen"].Width = 145;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["Ratio M"].Hidden = true;
			if (this.uGridInformacion.Rows.Count > 0)
			{				
				int iContColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;					
					
				string nombrCol = "";

				for (int j = 2; j < iContColumnas; j++)
				{
					nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;	
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 130;
					//						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0.00";
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;	
				}
			}

		}

		private void DiseñoGridTurno(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			try
			{				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Bodega"].Hidden = true;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Almacen"].Width = 145;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Ratio M"].Hidden = true;
				
				if (this.uGridInformacion.Rows.Count > 0)
				{				
					int iContColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;					
					
					string nombrCol = "";

					for (int j = 2; j < iContColumnas; j++)
					{
						nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;	
//						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left;
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 130;
						//						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0.00";
						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;	
					}
				}
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void DiseñoGridColor(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			
			this.uGridInformacion.Size = new Size(832, 386);
			this.ugridcolores.Visible = false;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["Bodega"].Hidden = true;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["Almacen"].Width = 145;
			this.uGridInformacion.DisplayLayout.Bands[0].Columns["TOTAL"].Hidden = true;

			string sColumna = "";
			int iContColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;	

			if (this.uGridInformacion.Rows.Count > 0)
			{				
//				int iContColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;					
					
				string nombrCol = "";

				for (int j = 2; j < iContColumnas; j++)
				{
					nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;	
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 73;
					//						this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0.00";
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;	
				}
			}
					
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
			{
				for (int i = 2; i < iContColumnas; i++)
				{

					sColumna = this.uGridInformacion.DisplayLayout.Bands[0].Columns[i].Header.Caption;			

					if (dr.Cells["0-0"].Value.ToString().Trim().Length == 0)  dr.Cells["0-0"].Value = 0;
					if (dr.Cells["1-5"].Value.ToString().Trim().Length == 0) dr.Cells["1-5"].Value = 0;
					if (dr.Cells["6-15"].Value.ToString().Trim().Length == 0) dr.Cells["6-15"].Value = 0;
					if (dr.Cells["16-30"].Value.ToString().Trim().Length == 0)  dr.Cells["16-30"].Value = 0;

					if (dr.Cells["31-60"].Value.ToString().Trim().Length == 0)  dr.Cells["31-60"].Value = 0;
					if (dr.Cells["61-90"].Value.ToString().Trim().Length == 0)  dr.Cells["61-90"].Value = 0;
					if (dr.Cells["91-99"].Value.ToString().Trim().Length == 0)  dr.Cells["91-99"].Value = 0;
					if (dr.Cells["100-180"].Value.ToString().Trim().Length == 0)  dr.Cells["100-180"].Value = 0;

					double dSumC=  Convert.ToDouble(dr.Cells["0-0"].Value) +  Convert.ToDouble(dr.Cells["1-5"].Value)  + 
						Convert.ToDouble(dr.Cells["6-15"].Value) +Convert.ToDouble(dr.Cells["16-30"].Value) ;
					#region Rango 0-0
					if( sColumna == "0-0")
					{
						if (dr.Cells[sColumna].Value.ToString().Trim().Length == 0) 
						{
							dr.Cells[sColumna].Appearance.BackColor = Color.White;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.White;								
						}
						else if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
						{
							//							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
							double dNumero = dSumC;
								
							if(dNumero >= 0 &&  dNumero <= 85)
							{
					
								dr.Cells[sColumna].Appearance.BackColor = Color.Red;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
							}
							if(dNumero >= 85.01 &&  dNumero <= 88.99)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
							}

							if(dNumero >= 89 &&  dNumero <= 94.99)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Yellow;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Yellow;
							}
							if(dNumero >= 95)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Green;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
							}
												
						}
					}
					#endregion Rango 0-0

					#region Rango 1-15
					if( sColumna == "1-5")
					{
						if (dr.Cells[sColumna].Value.ToString().Trim().Length == 0) 
						{
							dr.Cells[sColumna].Appearance.BackColor = Color.White;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.White;								
						}
						else if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
						{
							decimal dNumero = Convert.ToDecimal ( dr.Cells[sColumna].Value.ToString ());
								
							if(dNumero >= 0 &&  dNumero <= 2)
							{
					
								dr.Cells[sColumna].Appearance.BackColor = Color.Green;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
							}
							if(dNumero > 2 &&  dNumero <= 3)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Yellow;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Yellow;
							}

							if(dNumero > 3 &&  dNumero <= 4)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
							}
							if(dNumero > 4)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Red;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
							}
												
						}
					}
					#endregion Rango #region Rango 1-15

					#region Rango 6-15
					if( sColumna == "6-15")
					{
						if (dr.Cells[sColumna].Value.ToString().Trim().Length == 0) 
						{
							dr.Cells[sColumna].Appearance.BackColor = Color.White;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.White;								
						}
						else if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
						{
							decimal dNumero = Convert.ToDecimal ( dr.Cells[sColumna].Value.ToString ());
								
							if(dNumero >= 0 &&  dNumero <= 3)
							{
					
								dr.Cells[sColumna].Appearance.BackColor = Color.Green;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
							}
							if(dNumero > 3 &&  dNumero <= 4)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Yellow;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Yellow;
							}

							if(dNumero > 4 &&  dNumero <= 5)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
							}
							if(dNumero > 5)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Red;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
							}
												
						}
					}
					#endregion Rango 6-15

					#region Rango 16-30
					if( sColumna == "16-30")
					{
						if (dr.Cells[sColumna].Value.ToString().Trim().Length == 0) 
						{
							dr.Cells[sColumna].Appearance.BackColor = Color.White;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.White;								
						}
						else if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
						{
							decimal dNumero = Convert.ToDecimal ( dr.Cells[sColumna].Value.ToString ());
								
							if(dNumero >= 0 &&  dNumero <= 4)
							{
					
								dr.Cells[sColumna].Appearance.BackColor = Color.Green;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
							}
							if(dNumero > 4 &&  dNumero <= 5)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Yellow;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Yellow;
							}

							if(dNumero > 5 &&  dNumero <= 6)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
							}
							if(dNumero > 6)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Red;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
							}
												
						}
					}
					#endregion Rango #region Rango 16-30

					#region Rango 31-60
					if( sColumna == "31-60")
					{
						if (dr.Cells[sColumna].Value.ToString().Trim().Length == 0) 
						{
							dr.Cells[sColumna].Appearance.BackColor = Color.White;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.White;								
						}
						else if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
						{
							decimal dNumero = Convert.ToDecimal ( dr.Cells[sColumna].Value.ToString ());
								
							if(dNumero >= 0 &&  dNumero <= 3)
							{
					
								dr.Cells[sColumna].Appearance.BackColor = Color.Green;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
							}
							if(dNumero > 3 &&  dNumero <= 4)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Yellow;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Yellow;
							}

							if(dNumero > 4 &&  dNumero <= 5)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
							}
							if(dNumero > 5)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Red;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
							}
												
						}
					}
					#endregion Rango #region Rango 31-60

					#region Rango 61-90
					if( sColumna == "61-90")
					{
						if (dr.Cells[sColumna].Value.ToString().Trim().Length == 0) 
						{
							dr.Cells[sColumna].Appearance.BackColor = Color.White;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.White;								
						}
						else if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
						{
							decimal dNumero = Convert.ToDecimal ( dr.Cells[sColumna].Value.ToString ());
								
							if(dNumero >= 0 &&  dNumero <= 2)
							{
					
								dr.Cells[sColumna].Appearance.BackColor = Color.Green;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
							}
							if(dNumero > 2 &&  dNumero <= 3)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Yellow;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Yellow;
							}

							if(dNumero > 3 &&  dNumero <= 4)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
							}
							if(dNumero > 4)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Red;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
							}
												
						}
					}
					#endregion Rango 61-90

					#region Rango 91-99
					if( sColumna == "91-99")
					{
						if (dr.Cells[sColumna].Value.ToString().Trim().Length == 0) 
						{
							dr.Cells[sColumna].Appearance.BackColor = Color.White;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.White;								
						}
						else if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
						{
							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
								
							if(dNumero >= 0 &&  dNumero <= 1)
							{
					
								dr.Cells[sColumna].Appearance.BackColor = Color.Green;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Green;
							}
							if(dNumero > 1 &&  dNumero <= 1.5)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Yellow;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Yellow;
							}

							if(dNumero > 1.51 &&  dNumero <= 2)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Orange;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Orange;
							}
							if(dNumero > 2)
							{
										
								dr.Cells[sColumna].Appearance.BackColor = Color.Red;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
							}
												
						}
					}
					#endregion Rango 91-99

					
					#region Rango 99-180
					if( sColumna == "100-180")
					{
						if (dr.Cells[sColumna].Value.ToString().Trim().Length == 0) 
						{
							dr.Cells[sColumna].Appearance.BackColor = Color.White;
							dr.Cells[sColumna].Appearance.BackColor2 = Color.White;								
						}
						else if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
						{
							double dNumero = Convert.ToDouble ( dr.Cells[sColumna].Value.ToString ());
								
							if(dNumero >= 0 )
							{
					
								dr.Cells[sColumna].Appearance.BackColor = Color.Red;
								dr.Cells[sColumna].Appearance.BackColor2 = Color.Red;
							}
												
						}
					}
					#endregion Rango 99-180


//					if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
//					{
//						decimal dNumero = Convert.ToDecimal ( dr.Cells[sColumna].Value.ToString ());
//								
//						if(dNumero == 0)
//						{
//							dr.Cells[sColumna].Appearance.BackColor = Color.White;
//							dr.Cells[sColumna].Appearance.BackColor2 = Color.White;
//						}
//					}
				}
			}

			

		}
		private void ultraNumericEditor5_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraNumericEditor3_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void frmReporteRangoMorosidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();	
		}

		private void optRegion_ValueChanged(object sender, System.EventArgs e)
		{
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
			{
				dr.Cells["Sel"].Value = false;
			}
			
			if ((int)this.optRegion.Value == 0 )
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegion.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 1)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}


			if ((int)this.optRegion.Value == 1)
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegion.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 1)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}

			if ((int)this.optRegion.Value == 2)
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegion.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 2)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}

		}

		private void optPOrcentajes_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnConsultar_Click (sender , e);
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.Key;
			int iDesde = 0;
			int iHasta = 0 ;

			if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value)
			{
				MessageBox.Show(string.Format("No hay  registrado en la Columna : {0}", sColumna), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
						#region hola
			if (e.Cell.Column.ToString() == sColumna)
			{
				if (sColumna == "0-0")
				{
					iDesde = 0;
					iHasta = 0 ;
						
				}
				if (sColumna == "1-5")
				{
					iDesde = 1;
					iHasta = 5 ;
						
				}
				if (sColumna == "6-15")
				{
					iDesde = 6;
					iHasta = 15 ;
						
				}
				if (sColumna == "16-30")
				{
					iDesde = 16;
					iHasta = 30;
				}
				if (sColumna == "31-60")
				{
					iDesde = 31;
					iHasta = 60 ;
						
				}
				if (sColumna == "61-90")
				{
					iDesde = 61;
					iHasta = 90 ;
						
				}
				if (sColumna == "91-99")
				{
					iDesde = 91;
					iHasta = 99 ;
						
				}
				if (sColumna == "100-180")
				{
					iDesde = 100;
					iHasta = 180 ;
						
				}
				#endregion hola
//				int Numcuotas = Convert.ToInt32(sColumna);
				int idEntidadFinanciera  = 0;

				if (this.cmbEntidadFinanciera.ActiveRow != null) idEntidadFinanciera = (int)this.cmbEntidadFinanciera.Value;
//				int iNumCuota = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" select 99 - {0} ",Numcuotas));	
														
				using (frmReporteFDGClientes CCJ = new frmReporteFDGClientes ((DateTime)this.dtFecha.Value, idEntidadFinanciera,
								 (int)e.Cell.Row.Cells["Bodega"].Value, iDesde, iHasta, true, (string)e.Cell.Row.Cells["Almacen"].Value, miAcceso.BExportar))
				{
		
					if (DialogResult.OK == CCJ.ShowDialog())
					{
						
					}
				}		
			}
			}

		}
		
	}

