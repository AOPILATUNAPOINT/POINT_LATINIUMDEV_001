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
	/// Descripción breve de frmTiemposHabitaciones.
	/// </summary>
	public class frmTiemposHabitaciones : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ugHabitaciones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnGrabar;
		private C1.Data.C1DataSet cdsBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblBodega;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridHoras;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ChkHoras;
		private Infragistics.Win.UltraWinGrid.UltraGrid uDiasSemana;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ChkHabitaciones;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ChkDias;

		int laBodega = 0;

		public frmTiemposHabitaciones()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorasHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroHabitacion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHorasHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seleccion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmTiemposHabitaciones));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHorasHabitaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTmpHabitaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorTiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seleccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDiaHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seleccion");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorasHabitaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTmpHabitaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorTiempo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccion");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDiaHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccion");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			this.ugHabitaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGrabar = new System.Windows.Forms.Button();
			this.cdsBodega = new C1.Data.C1DataSet();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblBodega = new System.Windows.Forms.Label();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridHoras = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ChkHoras = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ChkHabitaciones = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uDiasSemana = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ChkDias = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridHoras)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uDiasSemana)).BeginInit();
			this.SuspendLayout();
			// 
			// ugHabitaciones
			// 
			this.ugHabitaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugHabitaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugHabitaciones.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugHabitaciones.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 117;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 109;
			ultraGridColumn3.Header.Caption = "Habitación";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 341;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn4.Header.Caption = "Nombre";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 276;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.Caption = "Num";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 45;
			ultraGridColumn6.Header.Caption = "Sel";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			ultraGridBand1.Header.Caption = "HABITACIONES";
			ultraGridBand1.HeaderVisible = true;
			this.ugHabitaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugHabitaciones.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.ugHabitaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugHabitaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugHabitaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.CardAreaAppearance = appearance4;
			this.ugHabitaciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ugHabitaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugHabitaciones.Location = new System.Drawing.Point(8, 40);
			this.ugHabitaciones.Name = "ugHabitaciones";
			this.ugHabitaciones.Size = new System.Drawing.Size(400, 368);
			this.ugHabitaciones.TabIndex = 19;
			this.ugHabitaciones.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugHabitaciones_CellChange);
			this.ugHabitaciones.Click += new System.EventHandler(this.ugHabitaciones_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(bool);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// btnGrabar
			// 
			this.btnGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGrabar.CausesValidation = false;
			this.btnGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btnGrabar.Image")));
			this.btnGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGrabar.Location = new System.Drawing.Point(8, 416);
			this.btnGrabar.Name = "btnGrabar";
			this.btnGrabar.Size = new System.Drawing.Size(72, 32);
			this.btnGrabar.TabIndex = 20;
			this.btnGrabar.Text = "Grabar";
			this.btnGrabar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGrabar.Click += new System.EventHandler(this.btnGrabar_Click);
			// 
			// cdsBodega
			// 
			this.cdsBodega.BindingContextCtrl = this;
			this.cdsBodega.DataLibrary = "LibFacturacion";
			this.cdsBodega.DataLibraryUrl = "";
			this.cdsBodega.DataSetDef = "dsBodega";
			this.cdsBodega.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsBodega.Name = "cdsBodega";
			this.cdsBodega.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsBodega.SchemaDef = null;
			this.cdsBodega.BeforeFill += new C1.Data.FillEventHandler(this.cdsBodega_BeforeFill);
			// 
			// ultraDataSource4
			// 
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource4;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Width = 30;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 30;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 30;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(56, 7);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(152, 21);
			this.cmbBodega.TabIndex = 132;
			this.cmbBodega.ValueMember = "idBodega";
			// 
			// ultraDataSource3
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(bool);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(16, 9);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 131;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btBuscar
			// 
			this.btBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btBuscar.Image")));
			this.btBuscar.Location = new System.Drawing.Point(216, 6);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(72, 22);
			this.btBuscar.TabIndex = 133;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(bool);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23});
			// 
			// uGridHoras
			// 
			this.uGridHoras.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridHoras.DataSource = this.ultraDataSource3;
			appearance9.BackColor = System.Drawing.Color.White;
			this.uGridHoras.DisplayLayout.Appearance = appearance9;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 53;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 76;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 35;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 39;
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 28;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn16.Width = 190;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellAppearance = appearance10;
			ultraGridColumn17.Format = "#,##0";
			ultraGridColumn17.Header.Caption = "Valor";
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn17.Width = 30;
			ultraGridColumn18.Header.Caption = "Sel";
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn18.Width = 30;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			ultraGridBand3.Header.Caption = "HORAS";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance11.FontData.SizeInPoints = 7.25F;
			ultraGridBand3.Override.HeaderAppearance = appearance11;
			this.uGridHoras.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHoras.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.uGridHoras.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridHoras.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.uGridHoras.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 10F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridHoras.DisplayLayout.Override.HeaderAppearance = appearance14;
			appearance15.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance15.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHoras.DisplayLayout.Override.RowAlternateAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHoras.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHoras.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.uGridHoras.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridHoras.Location = new System.Drawing.Point(696, 40);
			this.uGridHoras.Name = "uGridHoras";
			this.uGridHoras.Size = new System.Drawing.Size(280, 368);
			this.uGridHoras.TabIndex = 284;
			// 
			// ChkHoras
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.ChkHoras.Appearance = appearance18;
			this.ChkHoras.BackColor = System.Drawing.Color.Transparent;
			this.ChkHoras.CausesValidation = false;
			this.ChkHoras.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.ChkHoras.Location = new System.Drawing.Point(944, 40);
			this.ChkHoras.Name = "ChkHoras";
			this.ChkHoras.Size = new System.Drawing.Size(16, 16);
			this.ChkHoras.TabIndex = 308;
			this.ChkHoras.CheckedChanged += new System.EventHandler(this.ChkHoras_CheckedChanged);
			// 
			// ChkHabitaciones
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.ChkHabitaciones.Appearance = appearance19;
			this.ChkHabitaciones.BackColor = System.Drawing.Color.Transparent;
			this.ChkHabitaciones.CausesValidation = false;
			this.ChkHabitaciones.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.ChkHabitaciones.Location = new System.Drawing.Point(360, 40);
			this.ChkHabitaciones.Name = "ChkHabitaciones";
			this.ChkHabitaciones.Size = new System.Drawing.Size(16, 21);
			this.ChkHabitaciones.TabIndex = 312;
			this.ChkHabitaciones.CheckedChanged += new System.EventHandler(this.ChkHabitaciones_CheckedChanged);
			// 
			// uDiasSemana
			// 
			this.uDiasSemana.Cursor = System.Windows.Forms.Cursors.Default;
			this.uDiasSemana.DataSource = this.ultraDataSource2;
			appearance20.BackColor = System.Drawing.Color.White;
			this.uDiasSemana.DisplayLayout.Appearance = appearance20;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 75;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 35;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 46;
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Width = 220;
			ultraGridColumn23.Header.Caption = "Sel";
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn23.Width = 30;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			ultraGridBand4.Header.Caption = "DIAS SEMANA";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance21.FontData.SizeInPoints = 7.25F;
			ultraGridBand4.Override.HeaderAppearance = appearance21;
			this.uDiasSemana.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.uDiasSemana.DisplayLayout.Override.ActiveRowAppearance = appearance22;
			this.uDiasSemana.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uDiasSemana.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance23.BackColor = System.Drawing.Color.Transparent;
			this.uDiasSemana.DisplayLayout.Override.CardAreaAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance24.FontData.BoldAsString = "True";
			appearance24.FontData.Name = "Arial";
			appearance24.FontData.SizeInPoints = 10F;
			appearance24.ForeColor = System.Drawing.Color.White;
			appearance24.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uDiasSemana.DisplayLayout.Override.HeaderAppearance = appearance24;
			appearance25.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance25.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uDiasSemana.DisplayLayout.Override.RowAlternateAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uDiasSemana.DisplayLayout.Override.RowSelectorAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uDiasSemana.DisplayLayout.Override.SelectedRowAppearance = appearance27;
			this.uDiasSemana.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uDiasSemana.Location = new System.Drawing.Point(416, 40);
			this.uDiasSemana.Name = "uDiasSemana";
			this.uDiasSemana.Size = new System.Drawing.Size(280, 368);
			this.uDiasSemana.TabIndex = 315;
			this.uDiasSemana.Click += new System.EventHandler(this.uDiasSemana_Click);
			// 
			// ChkDias
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.ChkDias.Appearance = appearance28;
			this.ChkDias.BackColor = System.Drawing.Color.Transparent;
			this.ChkDias.CausesValidation = false;
			this.ChkDias.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.ChkDias.Location = new System.Drawing.Point(664, 40);
			this.ChkDias.Name = "ChkDias";
			this.ChkDias.Size = new System.Drawing.Size(16, 16);
			this.ChkDias.TabIndex = 316;
			this.ChkDias.CheckedChanged += new System.EventHandler(this.ChkDias_CheckedChanged);
			// 
			// frmTiemposHabitaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(984, 453);
			this.Controls.Add(this.ChkDias);
			this.Controls.Add(this.uDiasSemana);
			this.Controls.Add(this.ChkHabitaciones);
			this.Controls.Add(this.ChkHoras);
			this.Controls.Add(this.uGridHoras);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.btnGrabar);
			this.Controls.Add(this.ugHabitaciones);
			this.Name = "frmTiemposHabitaciones";
			this.Text = "Tiempos habitaciones";
			this.Load += new System.EventHandler(this.frmTiemposHabitaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridHoras)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uDiasSemana)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsBodega_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsBodega.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		bool activo = false;
		bool actv = false;

		private void frmTiemposHabitaciones_Load(object sender, System.EventArgs e)
		{
			string stExecBodega = string.Format("Select idBodega, Bodega, Codigo, Nombre From Bodega Where Activo = 1");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsBodega, stExecBodega);
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			laBodega = (int)this.cmbBodega.Value;
		  string ugSQL = string.Format("Exec Busca_Habitaciones {0}",laBodega.ToString());
			ugHabitaciones.DataSource = Funcion.dvProcedimiento(cdsBodega, ugSQL);
			//string dugSQL = string.Format("Exec Dias_Habitaciones {0}",laBodega.ToString());
			//uGridDias.DataSource = Funcion.dvProcedimiento(cdsBodega, dugSQL);
			//string hugSQL = string.Format("Exec Horas_Habitaciones {0}",laBodega.ToString());
			//uGridHoras.DataSource = Funcion.dvProcedimiento(cdsBodega, hugSQL);
		}

		private void btnGrabar_Click(object sender, System.EventArgs e)
		{
			#region habitaciones
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
				
				try
				{							
					// Graba habitaciones
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugHabitaciones.Rows.All)
					{
						string sSQLDetalle = string.Format("Exec Graba_Habitaciones_Tiempos {0}, {1}, {2}",
							(int)dr.Cells["Seleccion"].Value, 
							(int)dr.Cells["idHorasHabitacion"].Value, 
							(int)dr.Cells["idHabitacion"].Value);
						oCmdActualiza.CommandText = sSQLDetalle;
						oCmdActualiza.ExecuteNonQuery();
					}
					oTransaction.Commit();
				}
				catch (Exception ex)
				{
					try
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
			#endregion habitaciones
			// Visuliza información
			MessageBox.Show("Información grabada correctamente...!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
		}

		private void ugHabitaciones_Click(object sender, System.EventArgs e)
		{
			try
			{
				int idHabtcn = (int)this.ugHabitaciones.ActiveRow.Cells["idHabitacion"].Value;
				//string hugSQL = string.Format("Exec Horas_Habitaciones {0}",idHabtcn);
				//uGridHoras.DataSource = Funcion.dvProcedimiento(cdsBodega, hugSQL);
				string hugSQL = string.Format("Exec Dias_Habitaciones {0}",idHabtcn);
				uDiasSemana.DataSource = Funcion.dvProcedimiento(cdsBodega, hugSQL);
				activo = true;
				this.ChkDias.Checked = false;
				activo = false;
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void ChkHoras_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!actv)
			{
				int idHabtcn = (int)this.ugHabitaciones.ActiveRow.Cells["idHabitacion"].Value;
				int idDiHatcn = (int)this.uDiasSemana.ActiveRow.Cells["idDia"].Value;
				int idTmpHab = (int)this.uDiasSemana.ActiveRow.Cells["idDiaHabitacion"].Value;
				if (this.ChkHoras.Checked)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridHoras.Rows.All)
					{
						dr.Cells["Seleccion"].Value = true;
						string SP_Insert = string.Format("Exec Graba_Horas_Habitaciones {0},{1},{2},{3},{4}", 
							(int)dr.Cells["idHorasHabitaciones"].Value,
							idHabtcn,idDiHatcn,
							(int)dr.Cells["idTiempo"].Value,idTmpHab);
						Funcion.EjecutaSQL(cdsBodega,SP_Insert);
					}
				}
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridHoras.Rows.All)
					{
						dr.Cells["Seleccion"].Value = false;
						string SP_Insert = string.Format("Exec Graba_Horas_Habitaciones {0},{1},{2},{3},{4}", 
							(int)dr.Cells["idHorasHabitaciones"].Value,
							idHabtcn,idDiHatcn,
							(int)dr.Cells["idTiempo"].Value,idTmpHab);
						Funcion.EjecutaSQL(cdsBodega,SP_Insert);
					}
				}
			}
		}

		private void ChkHabitaciones_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.ChkHabitaciones.Checked)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugHabitaciones.Rows.All)
				{
					dr.Cells["Seleccion"].Value = true;
					string SP_Insert = string.Format("Exec Graba_Dias_Habitaciones {0},{1}", 1,(int)dr.Cells["idHabitacion"].Value);
					Funcion.EjecutaSQL(cdsBodega,SP_Insert);
				}
			}
			else
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugHabitaciones.Rows.All)
				{
					dr.Cells["Seleccion"].Value = false;
					string SP_Insert = string.Format("Exec Graba_Dias_Habitaciones {0},{1}", 0,(int)dr.Cells["idHabitacion"].Value);
					Funcion.EjecutaSQL(cdsBodega,SP_Insert);
				}
			}		
		}

		private void ugHabitaciones_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			//Procesa información al seleccionar el check
		}

		private void ChkDias_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!activo)
			{
				int idHabtcn = (int)this.ugHabitaciones.ActiveRow.Cells["idHabitacion"].Value;
				if (this.ChkDias.Checked)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uDiasSemana.Rows.All)
					{
						dr.Cells["Seleccion"].Value = true;
						string SP_Insert = string.Format("Exec Grb_Dias_Habitaciones {0},{1},{2}", 
							(int)dr.Cells["idDiaHabitacion"].Value,
							idHabtcn,
							(int)dr.Cells["idDia"].Value);
						Funcion.EjecutaSQL(cdsBodega,SP_Insert);
					}
				}
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uDiasSemana.Rows.All)
					{
						dr.Cells["Seleccion"].Value = false;
						string SP_Insert = string.Format("Exec Grb_Dias_Habitaciones {0},{1},{2}", 
							(int)dr.Cells["idDiaHabitacion"].Value,
							idHabtcn,
							(int)dr.Cells["idDia"].Value);
						Funcion.EjecutaSQL(cdsBodega,SP_Insert);
					}
				}
			}
		}

		private void uDiasSemana_Click(object sender, System.EventArgs e)
		{
			try
			{
				int idHabtcn = (int)this.uDiasSemana.ActiveRow.Cells["idHabitacion"].Value;
				int idDiaHabitcn = (int)this.uDiasSemana.ActiveRow.Cells["idDia"].Value;
				//string hugSQL = string.Format("Exec Horas_Habitaciones {0}",idHabtcn);
				//uGridHoras.DataSource = Funcion.dvProcedimiento(cdsBodega, hugSQL);
				string hugSQL = string.Format("Exec Horas_Habitaciones {0},{1}",idHabtcn,idDiaHabitcn);
				uGridHoras.DataSource = Funcion.dvProcedimiento(cdsBodega, hugSQL);
				actv = true;
				this.ChkHoras.Checked = false;
				actv = false;

			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
