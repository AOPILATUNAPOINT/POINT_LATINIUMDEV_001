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
	/// Descripción breve de frmCre_ResumenRetirosDeMercaderia.
	/// </summary>
	public class frmCre_ResumenRetirosDeMercaderia : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridNotasDeCredito;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridRetiros;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridRemates;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDebitos;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.ContextMenu mnuExportar;
		private System.Windows.Forms.MenuItem mnuExportarNDC;
		private System.Windows.Forms.MenuItem mnuExportarRetiros;
		private System.Windows.Forms.MenuItem mnuExportarRemates;
		private System.Windows.Forms.MenuItem mnuExportarDebitos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCobros;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPrecancelaciones;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridFDG;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_ResumenRetirosDeMercaderia()
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
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_ResumenRetirosDeMercaderia));
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			this.uGridNotasDeCredito = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnVer = new System.Windows.Forms.Button();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridRetiros = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridRemates = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridDebitos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExportar = new System.Windows.Forms.Button();
			this.mnuExportar = new System.Windows.Forms.ContextMenu();
			this.mnuExportarNDC = new System.Windows.Forms.MenuItem();
			this.mnuExportarRetiros = new System.Windows.Forms.MenuItem();
			this.mnuExportarRemates = new System.Windows.Forms.MenuItem();
			this.mnuExportarDebitos = new System.Windows.Forms.MenuItem();
			this.uGridCobros = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridPrecancelaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridFDG = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.uGridNotasDeCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRetiros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRemates)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDebitos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCobros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPrecancelaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridFDG)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridNotasDeCredito
			// 
			this.uGridNotasDeCredito.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridNotasDeCredito.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridNotasDeCredito.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.Header.Caption = "Notas De Credito";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridNotasDeCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridNotasDeCredito.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridNotasDeCredito.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridNotasDeCredito.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.uGridNotasDeCredito.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridNotasDeCredito.DisplayLayout.Override.CellAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 7.75F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridNotasDeCredito.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridNotasDeCredito.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridNotasDeCredito.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridNotasDeCredito.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.uGridNotasDeCredito.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridNotasDeCredito.Location = new System.Drawing.Point(8, 48);
			this.uGridNotasDeCredito.Name = "uGridNotasDeCredito";
			this.uGridNotasDeCredito.Size = new System.Drawing.Size(1056, 168);
			this.uGridNotasDeCredito.TabIndex = 736;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(384, 6);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 741;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(200, 10);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(34, 16);
			this.label17.TabIndex = 740;
			this.label17.Text = "Hasta";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 10);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(37, 16);
			this.label18.TabIndex = 739;
			this.label18.Text = "Desde";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance7;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 738;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance8;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 737;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// uGridRetiros
			// 
			this.uGridRetiros.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridRetiros.Cursor = System.Windows.Forms.Cursors.Default;
			appearance9.BackColor = System.Drawing.Color.White;
			this.uGridRetiros.DisplayLayout.Appearance = appearance9;
			ultraGridBand2.Header.Caption = "Retiros De Mercaderia Pichincha - Solidario - Point DATOS FALSOS + POR INCONFORMI" +
				"DAD + PAGO SOLIDARIO POR RETIRO INCOMPLETO DE PRODUCTOS + POR FALTA DE PAGO";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridRetiros.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridRetiros.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridRetiros.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridRetiros.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridRetiros.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridRetiros.DisplayLayout.Override.CellAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 7.75F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRetiros.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.uGridRetiros.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRetiros.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRetiros.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridRetiros.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRetiros.Location = new System.Drawing.Point(8, 216);
			this.uGridRetiros.Name = "uGridRetiros";
			this.uGridRetiros.Size = new System.Drawing.Size(1056, 72);
			this.uGridRetiros.TabIndex = 742;
			// 
			// uGridRemates
			// 
			this.uGridRemates.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridRemates.Cursor = System.Windows.Forms.Cursors.Default;
			appearance15.BackColor = System.Drawing.Color.White;
			this.uGridRemates.DisplayLayout.Appearance = appearance15;
			ultraGridBand3.Header.Caption = "Ventas De Remate Consumidor Final - Mena Dos";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridRemates.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridRemates.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridRemates.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridRemates.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridRemates.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridRemates.DisplayLayout.Override.CellAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 7.75F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRemates.DisplayLayout.Override.HeaderAppearance = appearance18;
			this.uGridRemates.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRemates.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRemates.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.uGridRemates.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRemates.Location = new System.Drawing.Point(8, 288);
			this.uGridRemates.Name = "uGridRemates";
			this.uGridRemates.Size = new System.Drawing.Size(1056, 88);
			this.uGridRemates.TabIndex = 743;
			// 
			// uGridDebitos
			// 
			this.uGridDebitos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDebitos.Cursor = System.Windows.Forms.Cursors.Default;
			appearance21.BackColor = System.Drawing.Color.White;
			this.uGridDebitos.DisplayLayout.Appearance = appearance21;
			ultraGridBand4.Header.Caption = "Debitos Bancarios Retiro Mercaderia Pichincha";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridDebitos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGridDebitos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDebitos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDebitos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.uGridDebitos.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDebitos.DisplayLayout.Override.CellAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance24.FontData.BoldAsString = "True";
			appearance24.FontData.Name = "Arial";
			appearance24.FontData.SizeInPoints = 7.75F;
			appearance24.ForeColor = System.Drawing.Color.White;
			appearance24.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDebitos.DisplayLayout.Override.HeaderAppearance = appearance24;
			this.uGridDebitos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDebitos.DisplayLayout.Override.RowSelectorAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDebitos.DisplayLayout.Override.SelectedRowAppearance = appearance26;
			this.uGridDebitos.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDebitos.Location = new System.Drawing.Point(8, 536);
			this.uGridDebitos.Name = "uGridDebitos";
			this.uGridDebitos.Size = new System.Drawing.Size(1056, 80);
			this.uGridDebitos.TabIndex = 744;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1088, 8);
			this.groupBox1.TabIndex = 745;
			this.groupBox1.TabStop = false;
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
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(976, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 746;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// mnuExportar
			// 
			this.mnuExportar.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuExportarNDC,
																																								this.mnuExportarRetiros,
																																								this.mnuExportarRemates,
																																								this.mnuExportarDebitos});
			// 
			// mnuExportarNDC
			// 
			this.mnuExportarNDC.Index = 0;
			this.mnuExportarNDC.Text = "Notas De Credito";
			this.mnuExportarNDC.Click += new System.EventHandler(this.mnuExportarNDC_Click);
			// 
			// mnuExportarRetiros
			// 
			this.mnuExportarRetiros.Index = 1;
			this.mnuExportarRetiros.Text = "Retiros";
			this.mnuExportarRetiros.Click += new System.EventHandler(this.mnuExportarRetiros_Click);
			// 
			// mnuExportarRemates
			// 
			this.mnuExportarRemates.Index = 2;
			this.mnuExportarRemates.Text = "Ventas De Remate";
			this.mnuExportarRemates.Click += new System.EventHandler(this.mnuExportarRemates_Click);
			// 
			// mnuExportarDebitos
			// 
			this.mnuExportarDebitos.Index = 3;
			this.mnuExportarDebitos.Text = "Debitos Banco Por Retiro";
			this.mnuExportarDebitos.Click += new System.EventHandler(this.mnuExportarDebitos_Click);
			// 
			// uGridCobros
			// 
			this.uGridCobros.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCobros.Cursor = System.Windows.Forms.Cursors.Default;
			appearance27.BackColor = System.Drawing.Color.White;
			this.uGridCobros.DisplayLayout.Appearance = appearance27;
			ultraGridBand5.Header.Caption = "Cobros Cartera Propia";
			ultraGridBand5.HeaderVisible = true;
			ultraGridBand5.SummaryFooterCaption = "";
			this.uGridCobros.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.uGridCobros.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCobros.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCobros.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance28.BackColor = System.Drawing.Color.Transparent;
			this.uGridCobros.DisplayLayout.Override.CardAreaAppearance = appearance28;
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCobros.DisplayLayout.Override.CellAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance30.FontData.BoldAsString = "True";
			appearance30.FontData.Name = "Arial";
			appearance30.FontData.SizeInPoints = 7.75F;
			appearance30.ForeColor = System.Drawing.Color.White;
			appearance30.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCobros.DisplayLayout.Override.HeaderAppearance = appearance30;
			this.uGridCobros.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCobros.DisplayLayout.Override.RowSelectorAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCobros.DisplayLayout.Override.SelectedRowAppearance = appearance32;
			this.uGridCobros.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCobros.Location = new System.Drawing.Point(8, 376);
			this.uGridCobros.Name = "uGridCobros";
			this.uGridCobros.Size = new System.Drawing.Size(1056, 80);
			this.uGridCobros.TabIndex = 747;
			// 
			// uGridPrecancelaciones
			// 
			this.uGridPrecancelaciones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPrecancelaciones.Cursor = System.Windows.Forms.Cursors.Default;
			appearance33.BackColor = System.Drawing.Color.White;
			this.uGridPrecancelaciones.DisplayLayout.Appearance = appearance33;
			ultraGridBand6.Header.Caption = "Debitos Bancarios Precancelaciones Solidario";
			ultraGridBand6.HeaderVisible = true;
			ultraGridBand6.SummaryFooterCaption = "";
			this.uGridPrecancelaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.uGridPrecancelaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPrecancelaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPrecancelaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance34.BackColor = System.Drawing.Color.Transparent;
			this.uGridPrecancelaciones.DisplayLayout.Override.CardAreaAppearance = appearance34;
			appearance35.ForeColor = System.Drawing.Color.Black;
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPrecancelaciones.DisplayLayout.Override.CellAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance36.FontData.BoldAsString = "True";
			appearance36.FontData.Name = "Arial";
			appearance36.FontData.SizeInPoints = 7.75F;
			appearance36.ForeColor = System.Drawing.Color.White;
			appearance36.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPrecancelaciones.DisplayLayout.Override.HeaderAppearance = appearance36;
			this.uGridPrecancelaciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPrecancelaciones.DisplayLayout.Override.RowSelectorAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPrecancelaciones.DisplayLayout.Override.SelectedRowAppearance = appearance38;
			this.uGridPrecancelaciones.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPrecancelaciones.Location = new System.Drawing.Point(8, 616);
			this.uGridPrecancelaciones.Name = "uGridPrecancelaciones";
			this.uGridPrecancelaciones.Size = new System.Drawing.Size(1056, 80);
			this.uGridPrecancelaciones.TabIndex = 749;
			// 
			// uGridFDG
			// 
			this.uGridFDG.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridFDG.Cursor = System.Windows.Forms.Cursors.Default;
			appearance39.BackColor = System.Drawing.Color.White;
			this.uGridFDG.DisplayLayout.Appearance = appearance39;
			ultraGridBand7.Header.Caption = "Fondo De Garantia Pichincha + Solidario";
			ultraGridBand7.HeaderVisible = true;
			ultraGridBand7.SummaryFooterCaption = "";
			this.uGridFDG.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.uGridFDG.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridFDG.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridFDG.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance40.BackColor = System.Drawing.Color.Transparent;
			this.uGridFDG.DisplayLayout.Override.CardAreaAppearance = appearance40;
			appearance41.ForeColor = System.Drawing.Color.Black;
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridFDG.DisplayLayout.Override.CellAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance42.FontData.BoldAsString = "True";
			appearance42.FontData.Name = "Arial";
			appearance42.FontData.SizeInPoints = 7.75F;
			appearance42.ForeColor = System.Drawing.Color.White;
			appearance42.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridFDG.DisplayLayout.Override.HeaderAppearance = appearance42;
			this.uGridFDG.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance43.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFDG.DisplayLayout.Override.RowSelectorAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFDG.DisplayLayout.Override.SelectedRowAppearance = appearance44;
			this.uGridFDG.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridFDG.Location = new System.Drawing.Point(8, 456);
			this.uGridFDG.Name = "uGridFDG";
			this.uGridFDG.Size = new System.Drawing.Size(1056, 80);
			this.uGridFDG.TabIndex = 748;
			this.uGridFDG.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridFDG_ClickCellButton);
			// 
			// frmCre_ResumenRetirosDeMercaderia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1072, 702);
			this.Controls.Add(this.uGridPrecancelaciones);
			this.Controls.Add(this.uGridFDG);
			this.Controls.Add(this.uGridCobros);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridDebitos);
			this.Controls.Add(this.uGridRemates);
			this.Controls.Add(this.uGridRetiros);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridNotasDeCredito);
			this.Name = "frmCre_ResumenRetirosDeMercaderia";
			this.Text = "Resumen De Retiros De Mercaderia";
			this.Load += new System.EventHandler(this.frmCre_ResumenRetirosDeMercaderia_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridNotasDeCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRetiros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRemates)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDebitos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCobros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPrecancelaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridFDG)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void frmCre_ResumenRetirosDeMercaderia_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "081217");

			if (!Funcion.Permiso("943", cdsSeteoF))
			{				
				MessageBox.Show("No tiene Acceso RESUMEN RETIROS", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if(miAcceso.BExportar)this.btnExportar.Enabled = true;

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Parse("01/01/" + DateTime.Today.Year.ToString());
			this.dtHasta.Value = DateTime.Today;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if(!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la Fecha Inicial", false,"", cdsSeteoF)) return;
			if(!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la Fecha Final", false,"", cdsSeteoF)) return;

			if(!Validacion.vbComparaFechas((DateTime) this.dtDesde.Value, (DateTime) this.dtHasta.Value, this.dtHasta,"La fecha inicial no puede ser posterior a la fecha final",false))return;
			#endregion Valida Fechas

			this.uGridNotasDeCredito.DataSource = null;
			this.uGridRetiros.DataSource = null;
			this.uGridRemates.DataSource = null;
			this.uGridDebitos.DataSource = null;
			this.uGridCobros.DataSource = null;
			this.uGridFDG.DataSource = null;
			this.uGridPrecancelaciones.DataSource = null;

			string sSQL = string.Format("Exec ResumenNotasDeCredito '{0}', '{1}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridNotasDeCredito);

			this.DiseñoGrid(this.uGridNotasDeCredito, false, true, "Motivo");

			sSQL = string.Format("Exec ResumenRetirosMercaderia '{0}', '{1}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridRetiros);

			this.DiseñoGrid(this.uGridRetiros, true, false, "");

			sSQL = string.Format("Exec ResumenVentasDeRemate '{0}', '{1}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridRemates);

			this.DiseñoGrid(this.uGridRemates, false, true, "Almacen");

			sSQL = string.Format("Exec ResumenDebitosRetiroMercaderia '{0}', '{1}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDebitos);

			this.DiseñoGrid(this.uGridDebitos, true, false, "");

			sSQL = string.Format("Exec ResumenAbonos '{0}', '{1}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridCobros);

			this.DiseñoGrid(this.uGridCobros, true, false, "Motivo");

			sSQL = string.Format("Exec ResumenFondoGarantia '{0}', '{1}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridFDG);

			this.DiseñoGrid(this.uGridFDG, true, false, "");

			sSQL = string.Format("Exec ResumenPrecancelaciones '{0}', '{1}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridPrecancelaciones);

			this.DiseñoGrid(this.uGridPrecancelaciones, true, false, "");
			this.DiseñoGridTurno();

		}
		private void DiseñoGridTurno()
		{
			try
			{				
//				this.uGridFDG.DisplayLayout.Bands[0].Columns["bodega"].Hidden = true;
//				this.uGridFDG.DisplayLayout.Bands[0].Columns["Local"].Header.Caption = "Caja";
//				this.uGridFDG.DisplayLayout.Bands[0].Columns["Local"].Width = 150;
//				this.uGridFDG.DisplayLayout.Bands[0].Columns["Local"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				
				int iContCerrado = 0;
				int iContAbierto = 0;

				if (this.uGridFDG.Rows.Count > 0)
				{				
					int iContColumnas = (int)this.uGridFDG.DisplayLayout.Bands[0].Columns.Count;					
					
					string nombrCol = "";

					for (int j = 0; j < iContColumnas; j++)
					{
						nombrCol = this.uGridFDG.DisplayLayout.Bands[0].Columns[j].Header.Caption;	
						this.uGridFDG.DisplayLayout.Bands[0].Columns[nombrCol].Style = ColumnStyle.Button;
						this.uGridFDG.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
						this.uGridFDG.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;
						this.uGridFDG.DisplayLayout.Bands[0].Columns[nombrCol].Format = "#,##0.00";
						this.uGridFDG.DisplayLayout.Bands[0].Columns[nombrCol].Header.Caption = nombrCol;	

					}
					
				}

//				this.lblContador.Text = iContAbierto.ToString()+ " Cajas Abiertas - " + iContCerrado.ToString() + " Cajas Cerradas " ;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void DiseñoGrid(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, bool bIniciaMes, bool bSumatorias, string sColumna)
		{
			#region Diseño Grid
			if (uGrid.Rows.Count > 0)
			{
				int iColIni = 0;

				if (!bIniciaMes)
				{
					iColIni = 1;

					uGrid.DisplayLayout.Bands[0].Columns[sColumna].Width = 200;
					uGrid.DisplayLayout.Bands[0].Columns[sColumna].CellActivation = Activation.ActivateOnly;
				}
				
				if (uGrid.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)uGrid.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = iColIni; j < iColumnas; j++)
					{
						nombrCol = uGrid.DisplayLayout.Bands[0].Columns[j].Header.Caption;
												
						uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						uGrid.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
						uGrid.DisplayLayout.Bands[0].Columns[nombrCol].Width = 80;

						if (bSumatorias)
						{
							uGrid.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, uGrid.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
							uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
							uGrid.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";
						}
					}					
				}				
			}			
			#endregion Diseño Grid
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void mnuExportarNDC_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridNotasDeCredito);
		}

		private void mnuExportarRetiros_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridRetiros);
		}

		private void mnuExportarRemates_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridRemates);
		}

		private void mnuExportarDebitos_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDebitos);
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			this.mnuExportar.Show(this.btnExportar, new Point(1, 25));			
		}

		private void uGridFDG_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
					
				string sColumna = e.Cell.Column.Key;
			
			
				if (e.Cell.Column.ToString() == sColumna)
				{
//					if (e.Cell.Row.Cells[sColumna].Appearance.BackColor == Color.White)
//					{
//						return;
//					}
					int iMes = 0 ;
          int iPeriodo = 0;

					string ssql= string.Format ("Exec RetornaMesAnoRetiroMercaderia '{0}'",sColumna); 
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
					dr.Read();
					iPeriodo = dr.GetInt32 (0);
					iMes  = dr.GetInt32 (1);
					dr.Close();

					DateTime dtInicio = DateTime.Parse("01/" + iMes + "/" + iPeriodo);
					DateTime dtFinalPeriodos = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
			
//					DateTime dtFecha = DateTime.Parse (sColumna + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString());
									
					using (frmCre_ResumenClientesFondoDeGarantia CCJ = new frmCre_ResumenClientesFondoDeGarantia (dtInicio,dtFinalPeriodos, sColumna))
					{
						if (DialogResult.OK == CCJ.ShowDialog())
						{

						}
					}				 				
				}
			
		}
	}
}
