using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteCognoConsultas.
	/// </summary>
	public class frmReporteCognoConsultas : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkLocal;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocal;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optRegiones;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteCognoConsultas()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IDENTIFICACION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEC_NACIMIENTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CORREO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ACT_LABORAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EST_LABORAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TELEFONO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEGMENTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RESULTADO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VAL_CUOTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NUM_REQ");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RESPUESTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEC_INGRESO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BODEGA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEGMENTACION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RAN_ING");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CAP_PAG");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SCO_INC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SCORING");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SCO_SOB_END");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOT_VEN");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CAR_CAS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EST_SOLICITUD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EST_CIVIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NIV_EDUCACION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PAI_NACIMIENTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PRO_NACIMIENTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CAN_NACIMIENTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PARR_NACIMIENTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("COD_POSTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEC_RESPUESTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("APRODADOR");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CAR_CAS_HIS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CAR_VEN_HIS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ETAPA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EST_REQUERIMIENTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MOT_RECHAZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("USUARIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TIPOCONSULTA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteCognoConsultas));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IDENTIFICACION");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOMBRE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEC_NACIMIENTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GENERO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CORREO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ACT_LABORAL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EST_LABORAL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TELEFONO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEGMENTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RESULTADO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VAL_CUOTA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NUM_REQ");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RESPUESTA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEC_INGRESO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BODEGA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEGMENTACION");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RAN_ING");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CAP_PAG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SCO_INC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SCORING");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SCO_SOB_END");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOT_VEN");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CAR_CAS");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EST_SOLICITUD");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EST_CIVIL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NIV_EDUCACION");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PAI_NACIMIENTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PRO_NACIMIENTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CAN_NACIMIENTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PARR_NACIMIENTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("COD_POSTAL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEC_RESPUESTA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("APRODADOR");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CAR_CAS_HIS");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CAR_VEN_HIS");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ETAPA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EST_REQUERIMIENTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MOT_RECHAZO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("USUARIO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TIPOCONSULTA");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExcel = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.chkLocal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridLocal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.optRegiones = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(System.DateTime);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.DateTime);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn35.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40});
			// 
			// ultraDataSource6
			// 
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(bool);
			ultraDataColumn44.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44});
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
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1168, 16);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(67, 24);
			this.btnExcel.TabIndex = 1012;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(696, 16);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 1011;
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 75;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 60;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 60;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 150;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 100;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 120;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 100;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 100;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 60;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 60;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 60;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 60;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 60;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 40;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 60;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 90;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 100;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Width = 100;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Width = 100;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Width = 100;
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Width = 100;
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.Width = 80;
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn35.Width = 80;
			ultraGridColumn36.Header.VisiblePosition = 35;
			ultraGridColumn36.Width = 100;
			ultraGridColumn37.Header.VisiblePosition = 36;
			ultraGridColumn37.Width = 120;
			ultraGridColumn38.Header.VisiblePosition = 37;
			ultraGridColumn38.Width = 100;
			ultraGridColumn39.Header.VisiblePosition = 38;
			ultraGridColumn39.Width = 100;
			ultraGridColumn40.Header.Caption = "TIPO_CONSULTA";
			ultraGridColumn40.Header.VisiblePosition = 39;
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
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40});
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 8F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(232, 48);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1016, 352);
			this.uGridInformacion.TabIndex = 1010;
			// 
			// chkLocal
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkLocal.Appearance = appearance7;
			this.chkLocal.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkLocal.Checked = true;
			this.chkLocal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkLocal.Location = new System.Drawing.Point(184, 56);
			this.chkLocal.Name = "chkLocal";
			this.chkLocal.Size = new System.Drawing.Size(13, 16);
			this.chkLocal.TabIndex = 1009;
			this.chkLocal.CheckedChanged += new System.EventHandler(this.chkLocal_CheckedChanged);
			// 
			// uGridLocal
			// 
			this.uGridLocal.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocal.DataSource = this.ultraDataSource6;
			appearance8.BackColor = System.Drawing.Color.White;
			this.uGridLocal.DisplayLayout.Appearance = appearance8;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn41.Header.Caption = "Local";
			ultraGridColumn41.Header.VisiblePosition = 0;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.Width = 150;
			ultraGridColumn43.Header.Caption = "...";
			ultraGridColumn43.Header.VisiblePosition = 2;
			ultraGridColumn43.Width = 28;
			ultraGridColumn44.Header.VisiblePosition = 3;
			ultraGridColumn44.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44});
			ultraGridBand2.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			ultraGridBand2.SummaryFooterCaption = "TOTALES";
			this.uGridLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridLocal.DisplayLayout.GroupByBox.Hidden = true;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocal.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			this.uGridLocal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLocal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocal.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocal.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.uGridLocal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance12.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance12.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.RowAlternateAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocal.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridLocal.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridLocal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocal.Location = new System.Drawing.Point(8, 56);
			this.uGridLocal.Name = "uGridLocal";
			this.uGridLocal.Size = new System.Drawing.Size(208, 352);
			this.uGridLocal.TabIndex = 1008;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-104, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1496, 8);
			this.groupBox1.TabIndex = 1007;
			this.groupBox1.TabStop = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(176, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 1003;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 16);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(37, 16);
			this.label5.TabIndex = 1002;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optRegiones
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.optRegiones.Appearance = appearance15;
			this.optRegiones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optRegiones.CheckedIndex = 0;
			this.optRegiones.ItemAppearance = appearance16;
			this.optRegiones.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Costa";
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Sierra";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Oriente";
			this.optRegiones.Items.Add(valueListItem1);
			this.optRegiones.Items.Add(valueListItem2);
			this.optRegiones.Items.Add(valueListItem3);
			this.optRegiones.Items.Add(valueListItem4);
			this.optRegiones.ItemSpacingVertical = 10;
			this.optRegiones.Location = new System.Drawing.Point(336, 16);
			this.optRegiones.Name = "optRegiones";
			this.optRegiones.Size = new System.Drawing.Size(256, 24);
			this.optRegiones.TabIndex = 1006;
			this.optRegiones.Text = "Todos";
			this.optRegiones.Visible = false;
			this.optRegiones.ValueChanged += new System.EventHandler(this.optRegiones_ValueChanged);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(624, 16);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 1004;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtHasta
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance17;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(224, 16);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1001;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance18;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 16);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1000;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// frmReporteCognoConsultas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1248, 414);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.chkLocal);
			this.Controls.Add(this.uGridLocal);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.optRegiones);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Name = "frmReporteCognoConsultas";
			this.Text = "Reporte Cogno Consultas";
			this.Load += new System.EventHandler(this.frmReporteCognoConsultas_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmReporteCognoConsultas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200625");

			if (!Funcion.Permiso("1933", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar al Reporte Cogno Consultas ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportarExcel) this.btnExcel.Enabled = true;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.Today ;

			string sSQLb = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQLb);		

			if (iBodega == 0) 			
				this.optRegiones.Visible = true;						

			string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1}, '{2}'", MenuLatinium.IdUsuario, (int)this.optRegiones.Value, Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
		}

		private void optRegiones_ValueChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				dr.Cells["Sel"].Value = false;

			#region Todos
			//if ((int)this.optRegiones.Value == 0)
			//{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1}, '{2}'", MenuLatinium.IdUsuario, (int)this.optRegiones.Value, Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
					if ( (int) dr.Cells["Region"].Value == 1)
						dr.Cells["Sel"].Value = true;
			//}
			#endregion Todos

			/*
			#region Sierra 
			if ((int)this.optRegiones.Value == 1)
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegiones.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 1)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}
			#endregion Sierra 

			#region Costa 
			if ((int)this.optRegiones.Value == 2)
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegiones.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 2)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}
			#endregion Costa 

			#region Amazonia  
			if ((int)this.optRegiones.Value == 3)
			{
				string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0}, {1} ", MenuLatinium.IdUsuario, (int)this.optRegiones.Value);		
				this.uGridLocal.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
				{
					if ( (int) dr.Cells["Region"].Value == 2)
					{
						dr.Cells["Sel"].Value = true;
					}
				}
			}
			#endregion Amazonia 
*/
		}

		private void chkLocal_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocal.Rows.All)
			{
				dr.Cells["Sel"].Value = this.chkLocal.Checked;
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			#endregion Valida Fechas

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

		
			this.lblContador.Text = "";
			string sSQLPa = string.Format("Exec ReporteCognoConsultas '{0}', '{1}', '{2}'",Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocal);

			FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.uGridInformacion);

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
					
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}
	}
}
