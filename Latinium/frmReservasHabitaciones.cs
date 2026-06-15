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
	/// Descripción breve de frmReservasHabitaciones.
	/// </summary>
	public class frmReservasHabitaciones : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugHabitaciones;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFeFin;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFeIni;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private DevExpress.XtraEditors.SimpleButton btBuscar;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.TextBox txtBskArticulo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugArticulos;
		private System.Windows.Forms.Label label4;
		private C1.Data.C1DataSet cdsBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugDetalleReserva;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSubTotal;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumReserva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempo;
		private System.Windows.Forms.Label label14;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReservasHabitaciones()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHabitacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroHabitacion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tematica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PReserva");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn1", 0);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReservasHabitaciones));
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSucursal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tematica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PReserva");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn1", 0);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio2");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleReserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleReserva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idReserva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.ugHabitaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtFeFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFeIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btBuscar = new DevExpress.XtraEditors.SimpleButton();
			this.label12 = new System.Windows.Forms.Label();
			this.txtBskArticulo = new System.Windows.Forms.TextBox();
			this.ugArticulos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.cdsBodega = new C1.Data.C1DataSet();
			this.ugDetalleReserva = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSubTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtNumReserva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTiempo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDetalleReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 30;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 30;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 30;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(16, 24);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(184, 21);
			this.cmbBodega.TabIndex = 327;
			this.cmbBodega.ValueMember = "idBodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(16, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(39, 19);
			this.label5.TabIndex = 326;
			this.label5.Text = "Local";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 210;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn8.CellAppearance = appearance2;
			ultraGridColumn8.Header.Caption = "#";
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 30;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Width = 125;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance3;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Precio";
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Width = 70;
			ultraGridColumn11.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			ultraGridColumn11.CellButtonAppearance = appearance4;
			ultraGridColumn11.Header.Caption = "...";
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn11.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.ugHabitaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugHabitaciones.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.ugHabitaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugHabitaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugHabitaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.CardAreaAppearance = appearance6;
			this.ugHabitaciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugHabitaciones.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugHabitaciones.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ugHabitaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugHabitaciones.Location = new System.Drawing.Point(8, 144);
			this.ugHabitaciones.Name = "ugHabitaciones";
			this.ugHabitaciones.Size = new System.Drawing.Size(496, 328);
			this.ugHabitaciones.TabIndex = 325;
			this.ugHabitaciones.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugHabitaciones_ClickCellButton);
			this.ugHabitaciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugHabitaciones_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// dtFeFin
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFeFin.Appearance = appearance11;
			this.dtFeFin.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.dtFeFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFeFin.DateButtons.Add(dateButton1);
			this.dtFeFin.Format = "dd/MM/yyyy HH:mm";
			this.dtFeFin.Location = new System.Drawing.Point(360, 24);
			this.dtFeFin.Name = "dtFeFin";
			this.dtFeFin.NonAutoSizeHeight = 23;
			this.dtFeFin.ReadOnly = true;
			this.dtFeFin.Size = new System.Drawing.Size(136, 21);
			this.dtFeFin.SpinButtonsVisible = true;
			this.dtFeFin.TabIndex = 331;
			this.dtFeFin.Value = ((object)(resources.GetObject("dtFeFin.Value")));
			// 
			// dtFeIni
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFeIni.Appearance = appearance12;
			this.dtFeIni.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.dtFeIni.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFeIni.DateButtons.Add(dateButton2);
			this.dtFeIni.Format = "dd/MM/yyyy HH:mm";
			this.dtFeIni.Location = new System.Drawing.Point(224, 24);
			this.dtFeIni.Name = "dtFeIni";
			this.dtFeIni.NonAutoSizeHeight = 23;
			this.dtFeIni.Size = new System.Drawing.Size(136, 21);
			this.dtFeIni.SpinButtonsVisible = true;
			this.dtFeIni.TabIndex = 330;
			this.dtFeIni.Value = ((object)(resources.GetObject("dtFeIni.Value")));
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(360, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(63, 16);
			this.label2.TabIndex = 329;
			this.label2.Text = "Hora Salida";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(224, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(71, 16);
			this.label1.TabIndex = 328;
			this.label1.Text = "Hora Entrada";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(72, 48);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(121, 16);
			this.label3.TabIndex = 333;
			this.label3.Text = "Descripción RESERVA";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.AcceptsReturn = true;
			this.txtDescripcion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance13;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(72, 64);
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(432, 72);
			this.txtDescripcion.TabIndex = 332;
			// 
			// btBuscar
			// 
			this.btBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btBuscar.Image")));
			this.btBuscar.Location = new System.Drawing.Point(736, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(80, 32);
			this.btBuscar.TabIndex = 336;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// label12
			// 
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(512, 8);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(64, 16);
			this.label12.TabIndex = 334;
			this.label12.Text = "ARTICULO";
			// 
			// txtBskArticulo
			// 
			this.txtBskArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBskArticulo.Location = new System.Drawing.Point(512, 24);
			this.txtBskArticulo.Name = "txtBskArticulo";
			this.txtBskArticulo.Size = new System.Drawing.Size(216, 20);
			this.txtBskArticulo.TabIndex = 335;
			this.txtBskArticulo.Text = "";
			// 
			// ugArticulos
			// 
			this.ugArticulos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugArticulos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugArticulos.DataSource = this.ultraDataSource2;
			appearance14.BackColor = System.Drawing.Color.White;
			this.ugArticulos.DisplayLayout.Appearance = appearance14;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 30;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Width = 305;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance15;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Precio";
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn16.Width = 70;
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance16.Image = ((object)(resources.GetObject("appearance16.Image")));
			ultraGridColumn18.CellButtonAppearance = appearance16;
			ultraGridColumn18.Header.Caption = "...";
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 20;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.ugArticulos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugArticulos.DisplayLayout.Override.ActiveRowAppearance = appearance17;
			this.ugArticulos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugArticulos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugArticulos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.ugArticulos.DisplayLayout.Override.CardAreaAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 10F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugArticulos.DisplayLayout.Override.HeaderAppearance = appearance19;
			appearance20.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance20.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.RowAlternateAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.ugArticulos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugArticulos.Location = new System.Drawing.Point(512, 48);
			this.ugArticulos.Name = "ugArticulos";
			this.ugArticulos.Size = new System.Drawing.Size(432, 144);
			this.ugArticulos.TabIndex = 337;
			this.ugArticulos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugArticulos_ClickCellButton);
			this.ugArticulos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugArticulos_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// label4
			// 
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(512, 200);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(168, 16);
			this.label4.TabIndex = 339;
			this.label4.Text = "DETALLE Reserva";
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
			// ugDetalleReserva
			// 
			this.ugDetalleReserva.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugDetalleReserva.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugDetalleReserva.DataSource = this.ultraDataSource3;
			appearance23.BackColor = System.Drawing.Color.White;
			this.ugDetalleReserva.DisplayLayout.Appearance = appearance23;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Width = 250;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance24;
			ultraGridColumn23.Format = "#,##0";
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.PromptChar = ' ';
			ultraGridColumn23.Width = 30;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance25;
			ultraGridColumn24.Format = "#,##0.00000";
			ultraGridColumn24.Header.VisiblePosition = 5;
			ultraGridColumn24.PromptChar = ' ';
			ultraGridColumn24.Width = 60;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance26;
			ultraGridColumn25.Format = "#,##0.00000";
			ultraGridColumn25.Header.Caption = "Subtotal";
			ultraGridColumn25.Header.VisiblePosition = 6;
			ultraGridColumn25.PromptChar = ' ';
			ultraGridColumn25.Width = 60;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			ultraGridBand4.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ugDetalleReserva.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugDetalleReserva.DisplayLayout.Override.ActiveRowAppearance = appearance27;
			this.ugDetalleReserva.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugDetalleReserva.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugDetalleReserva.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance28.BackColor = System.Drawing.Color.Transparent;
			this.ugDetalleReserva.DisplayLayout.Override.CardAreaAppearance = appearance28;
			this.ugDetalleReserva.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance29.FontData.BoldAsString = "True";
			appearance29.FontData.Name = "Arial";
			appearance29.FontData.SizeInPoints = 10F;
			appearance29.ForeColor = System.Drawing.Color.White;
			appearance29.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugDetalleReserva.DisplayLayout.Override.HeaderAppearance = appearance29;
			appearance30.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance30.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleReserva.DisplayLayout.Override.RowAlternateAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleReserva.DisplayLayout.Override.RowSelectorAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleReserva.DisplayLayout.Override.SelectedRowAppearance = appearance32;
			this.ugDetalleReserva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugDetalleReserva.Location = new System.Drawing.Point(512, 216);
			this.ugDetalleReserva.Name = "ugDetalleReserva";
			this.ugDetalleReserva.Size = new System.Drawing.Size(432, 168);
			this.ugDetalleReserva.TabIndex = 341;
			this.ugDetalleReserva.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ugDetalleReserva_BeforeRowsDeleted);
			this.ugDetalleReserva.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugDetalleReserva_KeyDown);
			this.ugDetalleReserva.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugDetalleReserva_AfterCellUpdate);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance33;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(864, 440);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(80, 31);
			this.txtTotal.TabIndex = 347;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance34;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(864, 416);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 21);
			this.txtIva.TabIndex = 346;
			// 
			// txtSubTotal
			// 
			this.txtSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSubTotal.Appearance = appearance35;
			this.txtSubTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSubTotal.Enabled = false;
			this.txtSubTotal.FormatString = "#,##0.00";
			this.txtSubTotal.Location = new System.Drawing.Point(864, 390);
			this.txtSubTotal.Name = "txtSubTotal";
			this.txtSubTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSubTotal.PromptChar = ' ';
			this.txtSubTotal.Size = new System.Drawing.Size(80, 21);
			this.txtSubTotal.TabIndex = 345;
			// 
			// label8
			// 
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(776, 447);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 16);
			this.label8.TabIndex = 344;
			this.label8.Text = "TOTAL";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(776, 418);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 16);
			this.label7.TabIndex = 343;
			this.label7.Text = "I.V.A.";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(776, 392);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(80, 16);
			this.label6.TabIndex = 342;
			this.label6.Text = "SUB TOTAL";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(520, 392);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(88, 19);
			this.label9.TabIndex = 348;
			this.label9.Text = "# RESERVA:";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumReserva
			// 
			this.txtNumReserva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumReserva.Appearance = appearance36;
			this.txtNumReserva.BorderStyle = Infragistics.Win.UIElementBorderStyle.None;
			appearance37.BorderColor = System.Drawing.Color.White;
			appearance37.BorderColor3DBase = System.Drawing.Color.White;
			this.txtNumReserva.ButtonAppearance = appearance37;
			this.txtNumReserva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumReserva.Enabled = false;
			this.txtNumReserva.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumReserva.FormatString = "#,##0";
			this.txtNumReserva.Location = new System.Drawing.Point(616, 391);
			this.txtNumReserva.Name = "txtNumReserva";
			this.txtNumReserva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNumReserva.PromptChar = ' ';
			this.txtNumReserva.ReadOnly = true;
			this.txtNumReserva.Size = new System.Drawing.Size(80, 20);
			this.txtNumReserva.TabIndex = 349;
			// 
			// txtTiempo
			// 
			this.txtTiempo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempo.Appearance = appearance38;
			this.txtTiempo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempo.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTiempo.FormatString = "#0";
			this.txtTiempo.Location = new System.Drawing.Point(16, 64);
			this.txtTiempo.Name = "txtTiempo";
			this.txtTiempo.PromptChar = ' ';
			this.txtTiempo.Size = new System.Drawing.Size(40, 24);
			this.txtTiempo.TabIndex = 351;
			// 
			// label14
			// 
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(16, 48);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(40, 16);
			this.label14.TabIndex = 350;
			this.label14.Text = "Tmpo";
			// 
			// frmReservasHabitaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(952, 477);
			this.Controls.Add(this.txtTiempo);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtNumReserva);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.txtBskArticulo);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtSubTotal);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.ugDetalleReserva);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.ugArticulos);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.dtFeFin);
			this.Controls.Add(this.dtFeIni);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.ugHabitaciones);
			this.KeyPreview = true;
			this.Name = "frmReservasHabitaciones";
			this.Text = "Reservas de Habitaciones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReservasHabitaciones_KeyDown);
			this.Load += new System.EventHandler(this.frmReservasHabitaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugHabitaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDetalleReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSubTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region variables
    int idBodg = 0;
		decimal PorcentajeIva = 0.00m;
		int idRes = 0;
		int idnHbtcn = 0;
		#endregion variables

		private void cdsBodega_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsBodega.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmReservasHabitaciones_Load(object sender, System.EventArgs e)
		{
			/* Tiempo de habitación */
			string nHoras = "SELECT dbo.Tmpo_Habitaciones()";
			this.txtTiempo.Value = Funcion.iEscalarSQL(cdsBodega, nHoras); 
			/* Porcentaje del IVA */
			//string lcPiva = "SELECT Iva FROM SeteoF WHERE idSeteoF=1";
			string lcPiva = "Select dbo.RetornaIVAVigenteLote()";
			PorcentajeIva = Funcion.decEscalarSQL(cdsBodega, lcPiva); 
			this.dtFeIni.Value = DateTime.Now;
			this.dtFeFin.Value = DateTime.Now;
			string stExecBodega = string.Format("Select idBodega, Bodega, Codigo, Nombre From Bodega Where Activo=1 And Almacen=1");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsBodega, stExecBodega);
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			idBodg = (int)this.cmbBodega.Value;
			string ugSQL = string.Format("Exec HabitacionesReserva {0}",idBodg);
			ugHabitaciones.DataSource = Funcion.dvProcedimiento(cdsBodega, ugSQL);
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			string stBusca = txtBskArticulo.Text.Trim();
			stBusca = stBusca.Replace("'", "''");
				
			//iBodega = (int)this.cmbBodega.Value;			

			DateTime dtFecha = DateTime.Today;

			string sSQL = string.Format("Exec ConsultaInv_Lotte '{0}'", stBusca);
			ugArticulos.DataSource = Funcion.dvProcedimiento(cdsBodega, sSQL);
					
			ugArticulos.Focus();
			this.Cursor = Cursors.Default;
		}

		private void ugHabitaciones_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Registrar articulo en detalle de factura
			try
			{
				idnHbtcn = (int)ugHabitaciones.ActiveRow.Cells["idHabitacion"].Value;
				DateTime FeInicia = (DateTime)this.dtFeIni.Value;

				string lcQry = string.Format("Select dbo.VerificaReserva("+idnHbtcn.ToString()+","+"'"+FeInicia.ToString("yyyyMMdd HH:mm")+"',"+idBodg.ToString()+")");
				string tpres = Funcion.sEscalarSQL(cdsBodega,lcQry);

				if(tpres.Equals("Existe"))
				{
					MessageBox.Show("Error: Ya existe una reserva para la HABITACION.?", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				else
				{
					string SP_Insert = string.Format("Exec Reserva_Habitacion {0},{1},'{2}','{3}',{4}", 
						idnHbtcn, idBodg,
						FeInicia.ToString("yyyyMMdd HH:mm"),
						this.txtDescripcion.Value.ToString(),
						this.txtTiempo.Value);
					idRes = Funcion.iEscalarSQL(cdsBodega,SP_Insert);
					string QryDetFc = "SELECT a.idDetalleReserva,a.idReserva,a.idArticulo,b.Articulo,a.Cantidad,a.Precio,a.Cantidad*a.Precio as SubTotal ";
					QryDetFc += "From DetalleReserva a ";
					QryDetFc += "Left Join Articulo b on a.idArticulo=b.idArticulo ";
					QryDetFc += "Where a.idReserva="+idRes.ToString();
					ugDetalleReserva.DataSource = Funcion.dvProcedimiento(cdsBodega, QryDetFc);
					txtNumReserva.Value = idRes;
					Total();
				}
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void Total()
		{
			this.txtIva.Value = 0.0000m;
			this.txtSubTotal.Value = 0.0000m;
			this.txtTotal.Value = 0.00m;

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugDetalleReserva.Rows.All)
			{
				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 4);
					this.txtSubTotal.Value = (decimal)this.txtSubTotal.Value + Math.Round(Convert.ToDecimal(dr.Cells["SubTotal"].Value), 4);
					//Actualiza la cantidad de producto
					string Updt = string.Format("Exec Actualiza_DetReserva_Lotte {0},{1}", dr.Cells["idDetalleReserva"].Value.ToString(),dr.Cells["Cantidad"].Value.ToString());
					Funcion.EjecutaSQL(cdsBodega,Updt);
				}
			}
			this.txtIva.Value = Math.Round((Convert.ToDecimal(this.txtSubTotal.Value) * PorcentajeIva /100.00m), 4);

			this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtSubTotal.Value), 4)
				+ Math.Round(Convert.ToDecimal(this.txtIva.Value), 4);
			#endregion Desplazamiento de grilla			
		}

		private void ugArticulos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			// Registrar articulo en detalle de factura
			try
			{
				if (idRes != 0)
				{
					int idArtFc = (int)ugArticulos.ActiveRow.Cells["idArticulo"].Value;
					decimal preFc = 0;
					preFc = (decimal)ugArticulos.ActiveRow.Cells["Precio1"].Value;
				
					/*Detalle reserva*/
					string SP_Insert = string.Format("Exec GrabaReserva_Habitacion {0},{1},{2},{3}", idRes,idArtFc,1,preFc);
					Funcion.EjecutaSQL(cdsBodega,SP_Insert);
					string QryDetFc = "SELECT a.idDetalleReserva,a.idReserva,a.idArticulo,b.Articulo,a.Cantidad,a.Precio,a.Cantidad*a.Precio as SubTotal ";
					QryDetFc += "From DetalleReserva a ";
					QryDetFc += "Left Join Articulo b on a.idArticulo=b.idArticulo ";
					QryDetFc += "Where a.idReserva="+idRes.ToString();
					ugDetalleReserva.DataSource = Funcion.dvProcedimiento(cdsBodega, QryDetFc);
					Total();
				}
				else 
				{
					MessageBox.Show ("Debe registrar habitación de reserva para ingresar artículos...!!!", "Aviso", 
						MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
				}
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}		
		}

		private void frmReservasHabitaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaReservasLotte br = new BuscaReservasLotte(idBodg))
				{
					if (DialogResult.OK == br.ShowDialog())
					{
						//Recupera información de reserva
						idRes = br.idResr;
						idnHbtcn = br.idHbta;
						dtFeIni.Value = br.FIni;
						dtFeFin.Value = br.FFin;
						txtDescripcion.Value = br.DesResrv;
						string QryDetFc = "SELECT a.idDetalleReserva,a.idReserva,a.idArticulo,b.Articulo,a.Cantidad,a.Precio,a.Cantidad*a.Precio as SubTotal ";
						QryDetFc += "From DetalleReserva a ";
						QryDetFc += "Left Join Articulo b on a.idArticulo=b.idArticulo ";
						QryDetFc += "Where a.idReserva="+idRes.ToString();
						ugDetalleReserva.DataSource = Funcion.dvProcedimiento(cdsBodega, QryDetFc);
						txtNumReserva.Value = idRes;
						Total();
					}
				}
			}
		}

		private void ugDetalleReserva_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			for (int i = 0; i < e.Rows.Length; i++)
			{
				if ((int)e.Rows[i].Cells["idDetalleReserva"].Value > 0)
				{	
					if (DialogResult.Yes ==	MessageBox.Show("Desea Eliminar Registro.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
					{
						Funcion.EjecutaSQL(cdsBodega, string.Format("Delete From DetalleReserva Where idDetalleReserva = {0}", (int)e.Rows[i].Cells["idDetalleReserva"].Value));
						e.DisplayPromptMsg = false;
						Total();
					}
					else e.Cancel = true;
				}
			}
		}

		private void ugDetalleReserva_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ugDetalleReserva.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{
							#region Enter
						case (int) Keys.Enter:
							ugDetalleReserva.PerformAction(UltraGridAction.ExitEditMode, false, false);
							ugDetalleReserva.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							ugDetalleReserva.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
							#endregion Enter
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error del sistema.");
				}
			}
		}

		private void ugDetalleReserva_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Cantidad
			if (e.Cell.Column.ToString() == "Cantidad")
			{
				if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value) e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
				if (true)
				{
					if (Funcion.bEscalarSQL(cdsBodega, string.Format("Declare @Res Bit = 0 If ((Select idTipoGrupo From Articulo Where idArticulo = {0}) In (1, 6) And (Select idGrupoArticulo From Articulo Where idArticulo = {0}) <> 26) Set @Res = 1 Select @Res", (int)e.Cell.Row.Cells["idArticulo"].Value), true))
					{
						int idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
						int Existencia = Funcion.iEscalarSQL(cdsBodega, string.Format("Select dbo.ExistenciasInventarioLotte({0}, {1})", idArticulo, idBodg));
						int iCantidad = 0;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugDetalleReserva.Rows.All)
						{
							if (idArticulo == (int)dr.Cells["idArticulo"].Value) iCantidad = iCantidad + (int)dr.Cells["Cantidad"].Value;
						}

            if (e.Cell.Column.ToString() == "Cantidad") Total();
					}
				}
			}
			#endregion Cantidad
		}

		private void ugHabitaciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ugArticulos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
