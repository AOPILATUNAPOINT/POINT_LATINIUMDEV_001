
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinToolbars;
//using Tulpep.NotificationWindow;
//using Infragistics.Win.

//using CustomUIControls;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_VistaRevisiondeCedulas.
	/// </summary>
	public class frmCre_VistaRevisiondeCedulas : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnSeleccion;
		public Infragistics.Win.UltraWinGrid.UltraGrid dgvBodegas;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAutomatico;
		private System.Windows.Forms.Button btnRefrescar;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraGrid dgvDatos;
		private System.Windows.Forms.Timer timer1;
		private System.Windows.Forms.Timer timer2;
		private System.Windows.Forms.NotifyIcon notifyIcon1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblMensaje;
		private System.ComponentModel.IContainer components;

		public frmCre_VistaRevisiondeCedulas()
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
		bool bSeleccion = false;
		int iContador = 0;
//		TaskbarNotifier taskbarNotifier;
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_VistaRevisiondeCedulas));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_Solicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSolicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Trabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCliente", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMensaje", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Solicitud", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedulas", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pagare", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Amortizacion", 5);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GastosDeCobranza", 6);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LugarDePago", 7);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngresosGastos", 8);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ReservaDeDominio", 9);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_Solicitud");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSolicitud");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaNacimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Trabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnSeleccion = new System.Windows.Forms.Button();
			this.dgvBodegas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dgvDatos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkAutomatico = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnRefrescar = new System.Windows.Forms.Button();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.timer2 = new System.Windows.Forms.Timer(this.components);
			this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblMensaje = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dgvBodegas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dgvDatos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm:ss";
			this.dtFecha.Location = new System.Drawing.Point(16, 16);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(152, 21);
			this.dtFecha.TabIndex = 779;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// btnSeleccion
			// 
			this.btnSeleccion.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSeleccion.CausesValidation = false;
			this.btnSeleccion.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSeleccion.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSeleccion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSeleccion.ImageIndex = 62;
			this.btnSeleccion.Location = new System.Drawing.Point(8, 56);
			this.btnSeleccion.Name = "btnSeleccion";
			this.btnSeleccion.Size = new System.Drawing.Size(104, 23);
			this.btnSeleccion.TabIndex = 826;
			this.btnSeleccion.Text = "Seleccionar Todo";
			this.btnSeleccion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSeleccion.Click += new System.EventHandler(this.btnSeleccion_Click);
			// 
			// dgvBodegas
			// 
			this.dgvBodegas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.dgvBodegas.Cursor = System.Windows.Forms.Cursors.Default;
			this.dgvBodegas.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dgvBodegas.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 50;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 175;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.dgvBodegas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dgvBodegas.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.dgvBodegas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.dgvBodegas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.dgvBodegas.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 7.75F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.dgvBodegas.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgvBodegas.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgvBodegas.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgvBodegas.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.dgvBodegas.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dgvBodegas.Location = new System.Drawing.Point(8, 88);
			this.dgvBodegas.Name = "dgvBodegas";
			this.dgvBodegas.Size = new System.Drawing.Size(280, 280);
			this.dgvBodegas.TabIndex = 827;
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
			// dgvDatos
			// 
			this.dgvDatos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.dgvDatos.Cursor = System.Windows.Forms.Cursors.Default;
			this.dgvDatos.DataSource = this.ultraDataSource2;
			appearance8.BackColor = System.Drawing.Color.White;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dgvDatos.DisplayLayout.Appearance = appearance8;
			ultraGridColumn5.Header.VisiblePosition = 15;
			ultraGridColumn6.Header.VisiblePosition = 16;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 150;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Width = 300;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Width = 150;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance9;
			ultraGridColumn13.Header.Caption = "Tiempo (min)";
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.Header.Caption = "Estado Cliente";
			ultraGridColumn15.Header.VisiblePosition = 21;
			ultraGridColumn16.Header.Caption = "F. Nacimiento";
			ultraGridColumn16.Header.VisiblePosition = 22;
			ultraGridColumn17.Header.Caption = "F. Vivienda";
			ultraGridColumn17.Header.VisiblePosition = 23;
			ultraGridColumn18.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn18.Header.Caption = "Tipo de Vivienda";
			ultraGridColumn18.Header.VisiblePosition = 24;
			ultraGridColumn18.Width = 200;
			ultraGridColumn19.Header.Caption = "F. Trabajo";
			ultraGridColumn19.Header.VisiblePosition = 25;
			ultraGridColumn20.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn20.Header.VisiblePosition = 26;
			ultraGridColumn20.Width = 200;
			ultraGridColumn21.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn21.Header.Caption = "Notas del Local";
			ultraGridColumn21.Header.VisiblePosition = 8;
			ultraGridColumn21.Width = 300;
			ultraGridColumn22.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn22.Header.Caption = "Respuesta";
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn22.Width = 55;
			ultraGridColumn23.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn23.Header.Caption = "Mensaje";
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn23.Width = 55;
			ultraGridColumn24.DataType = typeof(bool);
			ultraGridColumn24.Header.VisiblePosition = 9;
			ultraGridColumn24.Width = 90;
			ultraGridColumn25.DataType = typeof(bool);
			ultraGridColumn25.Header.VisiblePosition = 10;
			ultraGridColumn25.Width = 90;
			ultraGridColumn26.DataType = typeof(bool);
			ultraGridColumn26.Header.VisiblePosition = 12;
			ultraGridColumn26.Width = 90;
			ultraGridColumn27.DataType = typeof(bool);
			ultraGridColumn27.Header.VisiblePosition = 14;
			ultraGridColumn27.Width = 90;
			ultraGridColumn28.DataType = typeof(bool);
			ultraGridColumn28.Header.Caption = "Cobranza";
			ultraGridColumn28.Header.VisiblePosition = 17;
			ultraGridColumn28.Width = 90;
			ultraGridColumn29.DataType = typeof(bool);
			ultraGridColumn29.Header.Caption = "Lugar Pago";
			ultraGridColumn29.Header.VisiblePosition = 18;
			ultraGridColumn29.Width = 90;
			ultraGridColumn30.DataType = typeof(bool);
			ultraGridColumn30.Header.Caption = "I/G";
			ultraGridColumn30.Header.VisiblePosition = 19;
			ultraGridColumn30.Width = 90;
			ultraGridColumn31.DataType = typeof(bool);
			ultraGridColumn31.Header.Caption = "R De Dominio";
			ultraGridColumn31.Header.VisiblePosition = 20;
			ultraGridColumn31.Width = 90;
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
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
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
																										 ultraGridColumn31});
			this.dgvDatos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dgvDatos.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.dgvDatos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.dgvDatos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.dgvDatos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.dgvDatos.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 7.75F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.dgvDatos.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgvDatos.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgvDatos.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.dgvDatos.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.dgvDatos.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dgvDatos.Location = new System.Drawing.Point(296, 88);
			this.dgvDatos.Name = "dgvDatos";
			this.dgvDatos.Size = new System.Drawing.Size(904, 280);
			this.dgvDatos.TabIndex = 828;
			this.dgvDatos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.dgvDatos_ClickCellButton);
			this.dgvDatos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.dgvDatos_InitializeLayout);
			this.dgvDatos.SizeChanged += new System.EventHandler(this.dgvDatos_SizeChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(System.DateTime);
			ultraDataColumn17.DataType = typeof(System.DateTime);
			ultraDataColumn19.DataType = typeof(System.DateTime);
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
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// chkAutomatico
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAutomatico.Appearance = appearance16;
			this.chkAutomatico.BackColor = System.Drawing.Color.Transparent;
			this.chkAutomatico.CausesValidation = false;
			this.chkAutomatico.Location = new System.Drawing.Point(1016, 16);
			this.chkAutomatico.Name = "chkAutomatico";
			this.chkAutomatico.Size = new System.Drawing.Size(80, 22);
			this.chkAutomatico.TabIndex = 829;
			this.chkAutomatico.Text = "Automatico";
			this.chkAutomatico.CheckedChanged += new System.EventHandler(this.chkAutomatico_CheckedChanged);
			// 
			// btnRefrescar
			// 
			this.btnRefrescar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnRefrescar.CausesValidation = false;
			this.btnRefrescar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnRefrescar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnRefrescar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRefrescar.ImageIndex = 62;
			this.btnRefrescar.Location = new System.Drawing.Point(1112, 16);
			this.btnRefrescar.Name = "btnRefrescar";
			this.btnRefrescar.Size = new System.Drawing.Size(80, 23);
			this.btnRefrescar.TabIndex = 830;
			this.btnRefrescar.Text = "    Refrescar";
			this.btnRefrescar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRefrescar.Click += new System.EventHandler(this.btnRefrescar_Click);
			// 
			// optEstados
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance17;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 1;
			this.optEstados.ItemAppearance = appearance18;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendientes";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Respondido";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(936, 48);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(256, 32);
			this.optEstados.TabIndex = 831;
			this.optEstados.Text = "Pendientes";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// timer1
			// 
			this.timer1.Interval = 1000;
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// timer2
			// 
			this.timer2.Enabled = true;
			this.timer2.Interval = 1000;
			this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
			// 
			// notifyIcon1
			// 
			this.notifyIcon1.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon1.Icon")));
			this.notifyIcon1.Text = "notifyIcon1";
			this.notifyIcon1.Visible = true;
			this.notifyIcon1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseDown);
			this.notifyIcon1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseMove);
			this.notifyIcon1.DoubleClick += new System.EventHandler(this.notifyIcon1_DoubleClick);
			this.notifyIcon1.Click += new System.EventHandler(this.notifyIcon1_Click);
			this.notifyIcon1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseUp);
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
			// lblMensaje
			// 
			this.lblMensaje.AutoSize = true;
			this.lblMensaje.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblMensaje.ForeColor = System.Drawing.Color.Firebrick;
			this.lblMensaje.Location = new System.Drawing.Point(200, 16);
			this.lblMensaje.Name = "lblMensaje";
			this.lblMensaje.Size = new System.Drawing.Size(0, 17);
			this.lblMensaje.TabIndex = 13;
			this.lblMensaje.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmCre_VistaRevisiondeCedulas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1208, 382);
			this.Controls.Add(this.lblMensaje);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.btnRefrescar);
			this.Controls.Add(this.chkAutomatico);
			this.Controls.Add(this.dgvDatos);
			this.Controls.Add(this.dgvBodegas);
			this.Controls.Add(this.btnSeleccion);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCre_VistaRevisiondeCedulas";
			this.Text = "Revisión de Cédulas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_VistaRevisiondeCedulas_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmCre_VistaRevisiondeCedulas_Closing);
			this.Load += new System.EventHandler(this.frmCre_VistaRevisiondeCedulas_Load);
			this.Move += new System.EventHandler(this.frmCre_VistaRevisiondeCedulas_Move);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dgvBodegas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dgvDatos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnSeleccion_Click(object sender, System.EventArgs e)
		{
			bSeleccion = !bSeleccion;

			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow  dr in this.dgvBodegas.Rows)
			{
				dr.Cells["Sel"].Value = bSeleccion;
			}

			if (!bSeleccion) this.btnSeleccion.Text = "Seleccionar Todo"; else this.btnSeleccion.Text = "Quitar Seleccion"; 
		}

		private void frmCre_VistaRevisiondeCedulas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083401");

			if (!Funcion.Permiso("981", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Recepción de Cedulas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.timer1.Enabled = true;
//			string sSQLOSol = string.Format("Exec [ValidaLocalesRevisionCedula]{0} ");
//			this.dgvBodegas.DataSource = FuncionesProcedimientos.dtGeneral(sSQLOSol);

			string sSQL = string.Format("Exec [ValidaLocalesRevisionCedula] {0}", MenuLatinium.IdUsuario );
			this.dgvBodegas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);


		}

		private void Consulta()
		{
			#region Bodegas
			string sLocales = "";

			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.dgvBodegas.Rows)
			{
				if (Convert.ToBoolean(dr.Cells["Sel"].Value))
				{
					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString();
					else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}
			#endregion Bodegas

			#region Estados
//			int iEstado = 0;
//			if (this.rdbPendientes.Checked) iEstado = 1;
//			if (this.rdbRespondido.Checked) iEstado = 2;                        
			#endregion Estados

			#region Consulta de cedulas enviadas
      int iContador = 0;
			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.dgvBodegas.Rows)
			{
				if (Convert.ToBoolean(dr.Cells["Sel"].Value)) iContador++;
			}

			if (iContador > 0)
			{
				string sSQLB = string.Format("Exec Cre_ListaClientesRevisionCedulas '{0}','{1}',{2} ", sLocales,Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.optEstados.Value  );
				this.dgvDatos.DataSource = FuncionesProcedimientos.dtGeneral(sSQLB);


				#region Mensaje
				int iPendientes = 0;
				int iRespondido = 0;
				int iAprobados = 0;
				int iNegados = 0;

				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.dgvDatos.Rows)
				{
					if (Convert.ToInt32(dr.Cells["iEstado"].Value) == 1) iPendientes++;
					if (Convert.ToInt32(dr.Cells["iEstado"].Value) == 2) iRespondido++;
					if (Convert.ToInt32(dr.Cells["idEstado"].Value) == 1) iAprobados++;
					if (Convert.ToInt32(dr.Cells["idEstado"].Value) == 3) iNegados++;
				}
                                
   			this.lblMensaje.Text = "- " + this.dgvDatos.Rows.Count + " CLIENTES - " + iPendientes + " PENDIENTES - " + iRespondido + " RESPONDIDOS - " + iAprobados + " APROBADOS - " + iNegados + " NEGADOS";                
				#endregion Mensaje

				#region Colores
				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.dgvDatos.Rows)
				{
					if (Convert.ToInt32(row.Cells["iEstado"].Value) == 1)
					{
						row.Appearance.BackColor = Color.Red;
						row.Appearance.ForeColor = Color.White;
//             this.dgvDatos..Appearance.BackColor = System.Drawing.Color.White;
//					this.dgvDatos.Rows.Appearance.BackColor = Color.White;
					
					}
					if (Convert.ToInt32(row.Cells["iEstado"].Value) == 2)
					{
						row.Appearance.BackColor = Color.Orange;
						row.Appearance.ForeColor = Color.White;
//						this.dgvDatos.Rows.Appearance.BackColor = Color.Orange;
//						this.dgvDatos.Rows.Appearance.BackColor = Color.White;

				
					}
					if (Convert.ToInt32(row.Cells["iEstado"].Value) == 3)
					{
						row.Appearance.BackColor = Color.Green;
						row.Appearance.ForeColor = Color.White;
//						this.dgvDatos.Rows.Appearance.BackColor = Color.Green;
//					this.dgvDatos.Rows.Appearance.BackColor = Color.White;

			
					}
					if (Convert.ToInt32(row.Cells["iEstado"].Value) == 4)
					{
						row.Appearance.BackColor = Color.Blue;
						row.Appearance.ForeColor = Color.White;
//						this.dgvDatos.Rows.Appearance.BackColor = Color.Blue;
//						this.dgvDatos.Rows.Appearance.BackColor = Color.White;
				
					}
				}
				#endregion Colores

				#region Verifica nuevas consultas
				if (this.chkAutomatico.Checked)
				{
					int iContNuevas = 0;

					
         iContNuevas = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Cre_UltimasConsultas '{0}'", sLocales));		
//					nCre_Solicitudes.Cre_UltimasConsultas(sLocales);

					if (iContNuevas > 0)
					{
//						this.niInformacion.BalloonTipIcon = ToolTipIcon.Warning;
//						this.niInformacion.BalloonTipTitle = Application.ProductName;
//						this.niInformacion.BalloonTipText = string.Format("Tiene {0} Cedula(s) Por Revisar", iContNuevas);
//						this.niInformacion.ShowBalloonTip(8);
//						Infragistics.Win.UltraWinToolbars.UltraToolbarsManager TB = new Infragistics.Win.UltraWinToolbars.UltraToolbarsManager();
//
//						TB.ShowPopup(string.Format("Tiene {0} Cedula(s) Por Revisar", iContNuevas)); 
//
//
////						this.Activate();
//						this.WindowState = FormWindowState.Normal;
//						this.notifyIcon1.InitializeLifetimeService();
//						this.notifyIcon1.Text = "hola";
////						this.notifyIcon1.
//// 						MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional del Conyuge"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
////            MessageBox.Show( TopMost = true , "Mensaje", "Titulo", MessageBoxButtons.OK);
//						  showMessage("Este es mi mensaje", 10000);  /* 1 segundo = 1000 */

//						SystemSounds.Exclamation.Play();                   
					}
				}
				#endregion Verifica nuevas consultas


			
//				}
			}
			#endregion Consulta de cedulas enviadas
		}
		private void showMessage(string msg, int duration)
		{
//			using (Timer t = new Timer())
//			{
////				Timer time = new Timer();
////				time.Interval = duration;
//////				time.Tick += 	new ElapsedEventHandler(time); 
////				time.Tick += 	new EventHandler(time.Tick); 
////				/* Evento enlazado */
////
////				time.Start();
//
////				Timer t = new Timer();
//				t.Interval = duration;
//				t.Tick += new EventHandler(t_Elapsed);
//				t.Start();
//
//				/* Muestras el texto en el MB */
//				MessageBox.Show(msg);
//			} 
		}

		private void btnRefrescar_Click(object sender, System.EventArgs e)
		{
			 Consulta();
		}
		

	

		private void timer1_Tick(object sender, System.EventArgs e)
		{
			if ( this.chkAutomatico.Checked )
			{
				Consulta();
			}
//			(sender as Timer).Stop();  /* Detiene el Timer */
//			SendKeys.Send("{ESC}"); 
		}

		private void timer2_Tick(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Now.ToLongTimeString ();

//			 this.dtFecha.Value = DateTime.Now;
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			 this.Consulta();
		}

		private void notifyIcon1_DoubleClick(object sender, System.EventArgs e)
		{
		
		}

		private void frmCre_VistaRevisiondeCedulas_Move(object sender, System.EventArgs e)
		{

		}

		private void frmCre_VistaRevisiondeCedulas_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			

		}

		private void dgvDatos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

			string sUsuarioidCuponesWeb ="";
			sUsuarioidCuponesWeb = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT Nombre FROM Usuario WHERE idUsuario = {0}",
				MenuLatinium.IdUsuario ));	



			if (e.Cell.Column.ToString() == "btnCliente")
			{
				
					if ((int)e.Cell.Row.Cells["iEstado"].Value == 1)
					{

						if (this.chkAutomatico.Checked) this.timer1.Enabled = false;

						using (frmCre_RespuestaCedulas Revision = new frmCre_RespuestaCedulas())
						{
							if (DialogResult.OK == Revision.ShowDialog())
							{
								
								#region Guarda Cuota
								using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
								{
									oConexion.Open();

									SqlTransaction oTransaction = null;
									SqlCommand oCmdActualiza = oConexion.CreateCommand();
									oCmdActualiza.Connection = oConexion;
									oCmdActualiza.CommandType = CommandType.Text;
									oCmdActualiza.CommandTimeout = 120;
						
									try
									{
										oTransaction = oConexion.BeginTransaction();
										oCmdActualiza.Transaction = oTransaction;
								

										string sSQLSol = string.Format("Exec Cre_RespuestaHistorialRevisionCedulas {0},{1},'{2}',{3},{4},'{5}','{6}',{7}, {8} ",
											(int) e.Cell.Row.Cells["idCre_Solicitud"].Value ,
											(int) e.Cell.Row.Cells["iBodega"].Value ,

											Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),
										(int)Revision.optEstados.Value,
											0, 
											sUsuarioidCuponesWeb, 
											Revision.txtObservaciones.Text,
											2, 
											false
											);
					
										oCmdActualiza.CommandText = sSQLSol;
										oCmdActualiza.ExecuteNonQuery();	
	
										oTransaction.Commit();
					
									}
									catch (Exception ex)
									{
										oTransaction.Rollback();

										MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);	
										return;
									}
									finally
									{
										oConexion.Close();
												this.Consulta();
									}
								}
								#endregion Guarda Cuota
								
						
							}
						}

						if (this.chkAutomatico.Checked) this.timer1.Enabled = true;
				}
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void chkAutomatico_CheckedChanged(object sender, System.EventArgs e)
		{
			int iContador = 0;
			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.dgvBodegas.Rows)
			{
				if (Convert.ToBoolean(dr.Cells["Sel"].Value)) iContador++;
			}


			if (this.chkAutomatico.Checked)
			{
				if (iContador > 0)
				{
					Consulta();
					this.timer1.Enabled = true;
				}
			}
			else
			{
				this.timer1.Enabled = false;
			}
		}

		private void dgvDatos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void notifyIcon1_Click(object sender, System.EventArgs e)
		{
			this.Visible = true;

			this.WindowState = FormWindowState.Normal;
//			this.WindowState.ToString () = "Hola";
		
			this.notifyIcon1.InitializeLifetimeService();
				this.notifyIcon1.Text = "Credito Nuevo";
//				this.niInformacion.BalloonTipText = string.Format("Tiene {0} Cedula(s) Por Revisar", iContNuevas);
		}

		private void dgvDatos_SizeChanged(object sender, System.EventArgs e)
		{
//			if (this.WindowState == FormWindowState.Minimized)
//
//			{
//
//				this.Visible = false;
//
//				this.notifyIcon1.Visible = true;
//
//			}
		}

		private void notifyIcon1_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			
		}

		private void notifyIcon1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
		
		}

		private void notifyIcon1_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			
		}

		private void frmCre_VistaRevisiondeCedulas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}
	}
}
