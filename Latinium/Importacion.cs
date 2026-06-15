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
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Importacion.
	/// </summary>
	public class Importacion : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsImportacion;
		private System.Windows.Forms.Button btImportacion;
		private System.Windows.Forms.Button btAsiento;
		private C1.Data.C1DataSet cdsImportacionNombre;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoImportacion;
		private C1.Data.C1DataSet cdsTipoImportacion;
		private System.Windows.Forms.Button btImportacionTipo;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecio1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor8;
		private System.Windows.Forms.Button btCalcular;
		private System.Windows.Forms.Button btCrearAsientos;
		private System.Windows.Forms.Button btFactura;
		private System.Windows.Forms.Button btProgramacion;
		private System.Windows.Forms.Button btAprobacion;
		private System.Windows.Forms.Label lbl1;
		private System.Windows.Forms.Label lbl2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label lblAdvalorem;
		private System.Windows.Forms.Label lblFodinfa;
		private System.Windows.Forms.Label lblIce;
		private System.Windows.Forms.Label lblIvaImp;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor11;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesAduanizacion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtLlegadaPuerto;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEmbarque;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaFin;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtIngresoBodega;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasDeProduccion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasViaje;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAdValorem;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFodInfa;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIce;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSalidaCapitales;
		private System.Windows.Forms.Button btGrabar;
		private System.Windows.Forms.Button btEditar;
		private System.Windows.Forms.Button btDeshacer;
		private System.Windows.Forms.Button btSalir;
		private System.Windows.Forms.Button btResumenImp;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoCalculo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
   C1.Data.C1DataRow drImportacion;
		public Importacion()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ImportacionNombre", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFlete");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_ImportacionDet", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionDet", -1, "cmbTipoImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiento");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoImportacion", -1, "cmbTipoImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Adicional");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalS");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoObligatorio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalculo", -1, "cmbTipoCalculo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ImportacionTipo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Orden");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ivaimportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("salvaguardia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Exterior");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton6 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalculo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calculo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalculo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Calculo");
			this.label1 = new System.Windows.Forms.Label();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsImportacion = new C1.Data.C1DataSet();
			this.btImportacion = new System.Windows.Forms.Button();
			this.btAsiento = new System.Windows.Forms.Button();
			this.cdsImportacionNombre = new C1.Data.C1DataSet();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbTipoImportacion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsTipoImportacion = new C1.Data.C1DataSet();
			this.btImportacionTipo = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.txtPrecio1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtAdValorem = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFodInfa = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor8 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btCalcular = new System.Windows.Forms.Button();
			this.btCrearAsientos = new System.Windows.Forms.Button();
			this.btFactura = new System.Windows.Forms.Button();
			this.btProgramacion = new System.Windows.Forms.Button();
			this.btAprobacion = new System.Windows.Forms.Button();
			this.lbl1 = new System.Windows.Forms.Label();
			this.lbl2 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtIce = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblAdvalorem = new System.Windows.Forms.Label();
			this.lblFodinfa = new System.Windows.Forms.Label();
			this.lblIce = new System.Windows.Forms.Label();
			this.lblIvaImp = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.txtDiasViaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDiasDeProduccion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.dtDesAduanizacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label16 = new System.Windows.Forms.Label();
			this.dtLlegadaPuerto = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label9 = new System.Windows.Forms.Label();
			this.dtFechaEmbarque = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label10 = new System.Windows.Forms.Label();
			this.dtFechaFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label8 = new System.Windows.Forms.Label();
			this.dtIngresoBodega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label20 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraNumericEditor11 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label21 = new System.Windows.Forms.Label();
			this.txtSalidaCapitales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label22 = new System.Windows.Forms.Label();
			this.btGrabar = new System.Windows.Forms.Button();
			this.btEditar = new System.Windows.Forms.Button();
			this.btDeshacer = new System.Windows.Forms.Button();
			this.btSalir = new System.Windows.Forms.Button();
			this.btResumenImp = new System.Windows.Forms.Button();
			this.cmbTipoCalculo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsImportacionNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdValorem)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFodInfa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIce)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasViaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasDeProduccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesAduanizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLlegadaPuerto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEmbarque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIngresoBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalidaCapitales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCalculo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(78, 16);
			this.label1.TabIndex = 89;
			this.label1.Text = "IMPORTACIÓN";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.CausesValidation = false;
			this.barraDato1.DataMember = "_Importacion";
			this.barraDato1.DataNombreId = "idImportacion";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsImportacion;
			this.barraDato1.DataTabla = "Importacion";
			this.barraDato1.DataTablaHija = "ImportacionDet";
			this.barraDato1.Location = new System.Drawing.Point(8, 600);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(176, 26);
			this.barraDato1.TabIndex = 1;
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = false;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = false;
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cdsImportacion
			// 
			this.cdsImportacion.BindingContextCtrl = this;
			this.cdsImportacion.DataLibrary = "LibFacturacion";
			this.cdsImportacion.DataLibraryUrl = "";
			this.cdsImportacion.DataSetDef = "dsImportacion2";
			this.cdsImportacion.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsImportacion.Name = "cdsImportacion";
			this.cdsImportacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsImportacion.SchemaDef = null;
			this.cdsImportacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsImportacion_BeforeFill);
			// 
			// btImportacion
			// 
			this.btImportacion.Location = new System.Drawing.Point(864, 600);
			this.btImportacion.Name = "btImportacion";
			this.btImportacion.Size = new System.Drawing.Size(96, 24);
			this.btImportacion.TabIndex = 157;
			this.btImportacion.Text = "Importar Datos";
			this.btImportacion.Visible = false;
			this.btImportacion.Click += new System.EventHandler(this.btImportacion_Click);
			// 
			// btAsiento
			// 
			this.btAsiento.Location = new System.Drawing.Point(944, 8);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(120, 24);
			this.btAsiento.TabIndex = 19;
			this.btAsiento.Text = "Asiento ";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// cdsImportacionNombre
			// 
			this.cdsImportacionNombre.BindingContextCtrl = this;
			this.cdsImportacionNombre.DataLibrary = "LibFacturacion";
			this.cdsImportacionNombre.DataLibraryUrl = "";
			this.cdsImportacionNombre.DataSetDef = "dsImportacionNombre";
			this.cdsImportacionNombre.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsImportacionNombre.Name = "cdsImportacionNombre";
			this.cdsImportacionNombre.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsImportacionNombre.SchemaDef = null;
			this.cdsImportacionNombre.BeforeFill += new C1.Data.FillEventHandler(this.cdsImportacionNombre_BeforeFill);
			// 
			// cmbProyecto
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance1;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "Importacion.idImportacion"));
			this.cmbProyecto.DataMember = "ImportacionNombre";
			this.cmbProyecto.DataSource = this.cdsImportacionNombre;
			ultraGridBand1.AddButtonCaption = "SubProyecto";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 408;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(96, 8);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(320, 21);
			this.cmbProyecto.TabIndex = 0;
			this.cmbProyecto.ValueMember = "idImportacionNombre";
			this.cmbProyecto.Validated += new System.EventHandler(this.cmbProyecto_Validated);
			this.cmbProyecto.RowSelected += new Infragistics.Win.UltraWinGrid.RowSelectedEventHandler(this.cmbProyecto_RowSelected);
			this.cmbProyecto.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbProyecto_InitializeLayout);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_Importacion.Importacion - ImportacionDet";
			this.ultraGrid1.DataSource = this.cdsImportacion;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn4.Width = 64;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 61;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 51;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 53;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 57;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance3;
			ultraGridColumn9.Header.Caption = "Código";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 72;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 78;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance4;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn11.PromptChar = ' ';
			ultraGridColumn11.Width = 72;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance5;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn12.PromptChar = ' ';
			ultraGridColumn12.Width = 44;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance6;
			ultraGridColumn13.Header.Caption = "Asiento";
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.PromptChar = ' ';
			ultraGridColumn13.Width = 76;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.VisiblePosition = 8;
			ultraGridColumn14.Width = 190;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 7;
			ultraGridColumn15.Width = 72;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 63;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 69;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.Caption = "Tipo";
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.PromptChar = ' ';
			ultraGridColumn18.Width = 166;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance7;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 17;
			ultraGridColumn19.PromptChar = ' ';
			ultraGridColumn19.Width = 74;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance8;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 6;
			ultraGridColumn20.PromptChar = ' ';
			ultraGridColumn20.Width = 82;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance9;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Subtotal";
			ultraGridColumn21.Header.VisiblePosition = 18;
			ultraGridColumn21.PromptChar = ' ';
			ultraGridColumn21.Width = 74;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn22.Header.Caption = "...";
			ultraGridColumn22.Header.VisiblePosition = 19;
			ultraGridColumn22.PromptChar = ' ';
			ultraGridColumn22.Width = 30;
			ultraGridColumn23.Header.Caption = "Calculo";
			ultraGridColumn23.Header.VisiblePosition = 14;
			ultraGridColumn23.Width = 83;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn23});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 10F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 120);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1056, 344);
			this.ultraGrid1.TabIndex = 10;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cmbTipoImportacion
			// 
			this.cmbTipoImportacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoImportacion.DataMember = "ImportacionTipo";
			this.cmbTipoImportacion.DataSource = this.cdsTipoImportacion;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 579;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 3;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 4;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 5;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 6;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30});
			this.cmbTipoImportacion.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipoImportacion.DisplayMember = "Nombre";
			this.cmbTipoImportacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoImportacion.Location = new System.Drawing.Point(136, 280);
			this.cmbTipoImportacion.Name = "cmbTipoImportacion";
			this.cmbTipoImportacion.Size = new System.Drawing.Size(312, 104);
			this.cmbTipoImportacion.TabIndex = 163;
			this.cmbTipoImportacion.ValueMember = "idImportacionTipo";
			this.cmbTipoImportacion.Visible = false;
			// 
			// cdsTipoImportacion
			// 
			this.cdsTipoImportacion.BindingContextCtrl = this;
			this.cdsTipoImportacion.DataLibrary = "LibFacturacion";
			this.cdsTipoImportacion.DataLibraryUrl = "";
			this.cdsTipoImportacion.DataSetDef = "dsTipoImportacion";
			this.cdsTipoImportacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoImportacion.Name = "cdsTipoImportacion";
			this.cdsTipoImportacion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTipoImportacion.SchemaDef = null;
			this.cdsTipoImportacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsTipoImportacion_BeforeFill);
			// 
			// btImportacionTipo
			// 
			this.btImportacionTipo.Location = new System.Drawing.Point(688, 8);
			this.btImportacionTipo.Name = "btImportacionTipo";
			this.btImportacionTipo.Size = new System.Drawing.Size(120, 24);
			this.btImportacionTipo.TabIndex = 17;
			this.btImportacionTipo.Text = "Tipo de Importación";
			this.btImportacionTipo.Click += new System.EventHandler(this.btImportacionTipo_Click);
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label2.Location = new System.Drawing.Point(62, 568);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(24, 17);
			this.label2.TabIndex = 166;
			this.label2.Text = "CIF";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(378, 472);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(68, 16);
			this.label3.TabIndex = 167;
			this.label3.Text = "ADVALOREM";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(378, 496);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(50, 16);
			this.label4.TabIndex = 168;
			this.label4.Text = "FODINFA";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(378, 544);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(22, 16);
			this.label6.TabIndex = 169;
			this.label6.Text = "IVA";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label7.Location = new System.Drawing.Point(776, 544);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(91, 17);
			this.label7.TabIndex = 170;
			this.label7.Text = "TOTAL GASTOS";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label7.Click += new System.EventHandler(this.label7_Click);
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(62, 472);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(25, 16);
			this.label12.TabIndex = 176;
			this.label12.Text = "FOB";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(62, 496);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(51, 16);
			this.label13.TabIndex = 177;
			this.label13.Text = "FOB EXW";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(62, 520);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(114, 16);
			this.label14.TabIndex = 178;
			this.label14.Text = "FLETE DEL EXTERIOR";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(62, 544);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(47, 16);
			this.label15.TabIndex = 179;
			this.label15.Text = "SEGURO";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPrecio1
			// 
			this.txtPrecio1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecio1.Appearance = appearance15;
			this.txtPrecio1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.Fob"));
			this.txtPrecio1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecio1.FormatString = "#,##0.00";
			this.txtPrecio1.Location = new System.Drawing.Point(194, 472);
			this.txtPrecio1.Name = "txtPrecio1";
			this.txtPrecio1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecio1.PromptChar = ' ';
			this.txtPrecio1.Size = new System.Drawing.Size(96, 22);
			this.txtPrecio1.TabIndex = 184;
			// 
			// ultraNumericEditor1
			// 
			this.ultraNumericEditor1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor1.Appearance = appearance16;
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.FobExw"));
			this.ultraNumericEditor1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor1.FormatString = "#,##0.00";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(194, 496);
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(96, 22);
			this.ultraNumericEditor1.TabIndex = 185;
			// 
			// ultraNumericEditor2
			// 
			this.ultraNumericEditor2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor2.Appearance = appearance17;
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.FleteExt"));
			this.ultraNumericEditor2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor2.FormatString = "#,##0.00";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(194, 520);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(96, 22);
			this.ultraNumericEditor2.TabIndex = 186;
			// 
			// ultraNumericEditor3
			// 
			this.ultraNumericEditor3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor3.Appearance = appearance18;
			this.ultraNumericEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.Seguro"));
			this.ultraNumericEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor3.FormatString = "#,##0.00";
			this.ultraNumericEditor3.Location = new System.Drawing.Point(194, 544);
			this.ultraNumericEditor3.Name = "ultraNumericEditor3";
			this.ultraNumericEditor3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor3.PromptChar = ' ';
			this.ultraNumericEditor3.Size = new System.Drawing.Size(96, 22);
			this.ultraNumericEditor3.TabIndex = 187;
			// 
			// ultraNumericEditor4
			// 
			this.ultraNumericEditor4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor4.Appearance = appearance19;
			this.ultraNumericEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.Cif"));
			this.ultraNumericEditor4.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor4.FormatString = "#,##0.00";
			this.ultraNumericEditor4.Location = new System.Drawing.Point(194, 568);
			this.ultraNumericEditor4.Name = "ultraNumericEditor4";
			this.ultraNumericEditor4.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor4.PromptChar = ' ';
			this.ultraNumericEditor4.Size = new System.Drawing.Size(96, 22);
			this.ultraNumericEditor4.TabIndex = 188;
			// 
			// txtAdValorem
			// 
			this.txtAdValorem.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAdValorem.Appearance = appearance20;
			this.txtAdValorem.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.Advalorem"));
			this.txtAdValorem.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAdValorem.FormatString = "#,##0.00";
			this.txtAdValorem.Location = new System.Drawing.Point(514, 472);
			this.txtAdValorem.Name = "txtAdValorem";
			this.txtAdValorem.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAdValorem.PromptChar = ' ';
			this.txtAdValorem.Size = new System.Drawing.Size(96, 22);
			this.txtAdValorem.TabIndex = 11;
			this.txtAdValorem.Click += new System.EventHandler(this.txtAdValorem_Click);
			this.txtAdValorem.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAdValorem_KeyPress);
			this.txtAdValorem.Validated += new System.EventHandler(this.txtAdValorem_Validated);
			this.txtAdValorem.ValueChanged += new System.EventHandler(this.txtAdValorem_ValueChanged);
			this.txtAdValorem.Enter += new System.EventHandler(this.txtAdValorem_Enter);
			// 
			// txtFodInfa
			// 
			this.txtFodInfa.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFodInfa.Appearance = appearance21;
			this.txtFodInfa.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.Fodinfa"));
			this.txtFodInfa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFodInfa.FormatString = "#,##0.00";
			this.txtFodInfa.Location = new System.Drawing.Point(514, 496);
			this.txtFodInfa.Name = "txtFodInfa";
			this.txtFodInfa.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFodInfa.PromptChar = ' ';
			this.txtFodInfa.Size = new System.Drawing.Size(96, 22);
			this.txtFodInfa.TabIndex = 12;
			this.txtFodInfa.Click += new System.EventHandler(this.txtFodInfa_Click);
			this.txtFodInfa.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtFodInfa_KeyPress);
			this.txtFodInfa.Validated += new System.EventHandler(this.txtFodInfa_Validated);
			this.txtFodInfa.ValueChanged += new System.EventHandler(this.txtFodInfa_ValueChanged);
			this.txtFodInfa.Enter += new System.EventHandler(this.txtFodInfa_Enter);
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance22;
			this.txtIva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.IvaImpor"));
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(514, 544);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(96, 22);
			this.txtIva.TabIndex = 14;
			this.txtIva.Click += new System.EventHandler(this.txtIva_Click);
			this.txtIva.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIva_KeyPress);
			this.txtIva.Validated += new System.EventHandler(this.txtIva_Validated);
			this.txtIva.ValueChanged += new System.EventHandler(this.txtIva_ValueChanged);
			this.txtIva.Enter += new System.EventHandler(this.txtIva_Enter);
			// 
			// ultraNumericEditor8
			// 
			this.ultraNumericEditor8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor8.Appearance = appearance23;
			this.ultraNumericEditor8.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.TotalG"));
			this.ultraNumericEditor8.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor8.FormatString = "#,##0.00";
			this.ultraNumericEditor8.Location = new System.Drawing.Point(912, 544);
			this.ultraNumericEditor8.Name = "ultraNumericEditor8";
			this.ultraNumericEditor8.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor8.PromptChar = ' ';
			this.ultraNumericEditor8.Size = new System.Drawing.Size(96, 22);
			this.ultraNumericEditor8.TabIndex = 192;
			// 
			// btCalcular
			// 
			this.btCalcular.Location = new System.Drawing.Point(968, 600);
			this.btCalcular.Name = "btCalcular";
			this.btCalcular.Size = new System.Drawing.Size(96, 24);
			this.btCalcular.TabIndex = 193;
			this.btCalcular.Text = "Calcular Valores";
			this.btCalcular.Visible = false;
			this.btCalcular.Click += new System.EventHandler(this.btCalcular_Click);
			// 
			// btCrearAsientos
			// 
			this.btCrearAsientos.Location = new System.Drawing.Point(760, 600);
			this.btCrearAsientos.Name = "btCrearAsientos";
			this.btCrearAsientos.Size = new System.Drawing.Size(96, 24);
			this.btCrearAsientos.TabIndex = 194;
			this.btCrearAsientos.Text = "Crear Asientos";
			this.btCrearAsientos.Visible = false;
			this.btCrearAsientos.Click += new System.EventHandler(this.btCrearAsientos_Click);
			// 
			// btFactura
			// 
			this.btFactura.Location = new System.Drawing.Point(816, 8);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(120, 24);
			this.btFactura.TabIndex = 18;
			this.btFactura.Text = "Factura";
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// btProgramacion
			// 
			this.btProgramacion.Location = new System.Drawing.Point(560, 8);
			this.btProgramacion.Name = "btProgramacion";
			this.btProgramacion.Size = new System.Drawing.Size(120, 24);
			this.btProgramacion.TabIndex = 16;
			this.btProgramacion.Text = "Programación";
			this.btProgramacion.Click += new System.EventHandler(this.btProgramacion_Click);
			// 
			// btAprobacion
			// 
			this.btAprobacion.Location = new System.Drawing.Point(656, 600);
			this.btAprobacion.Name = "btAprobacion";
			this.btAprobacion.Size = new System.Drawing.Size(96, 24);
			this.btAprobacion.TabIndex = 197;
			this.btAprobacion.Text = "Aprobación";
			this.btAprobacion.Visible = false;
			this.btAprobacion.Click += new System.EventHandler(this.btAprobacion_Click);
			// 
			// lbl1
			// 
			this.lbl1.AutoSize = true;
			this.lbl1.BackColor = System.Drawing.Color.Transparent;
			this.lbl1.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lbl1.Location = new System.Drawing.Point(672, 88);
			this.lbl1.Name = "lbl1";
			this.lbl1.Size = new System.Drawing.Size(112, 19);
			this.lbl1.TabIndex = 198;
			this.lbl1.Text = "TOTAL GASTOS:";
			this.lbl1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lbl2
			// 
			this.lbl2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lbl2.AutoSize = true;
			this.lbl2.BackColor = System.Drawing.Color.Transparent;
			this.lbl2.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lbl2.Location = new System.Drawing.Point(632, 568);
			this.lbl2.Name = "lbl2";
			this.lbl2.Size = new System.Drawing.Size(112, 19);
			this.lbl2.TabIndex = 199;
			this.lbl2.Text = "TOTAL GASTOS:";
			this.lbl2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(378, 520);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(22, 16);
			this.label5.TabIndex = 200;
			this.label5.Text = "ICE";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIce
			// 
			this.txtIce.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIce.Appearance = appearance24;
			this.txtIce.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.Corpei"));
			this.txtIce.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIce.FormatString = "#,##0.00";
			this.txtIce.Location = new System.Drawing.Point(514, 520);
			this.txtIce.Name = "txtIce";
			this.txtIce.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIce.PromptChar = ' ';
			this.txtIce.Size = new System.Drawing.Size(96, 22);
			this.txtIce.TabIndex = 13;
			this.txtIce.Click += new System.EventHandler(this.txtIce_Click);
			this.txtIce.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIce_KeyPress);
			this.txtIce.Validated += new System.EventHandler(this.txtIce_Validated);
			this.txtIce.ValueChanged += new System.EventHandler(this.txtIce_ValueChanged);
			this.txtIce.Enter += new System.EventHandler(this.txtIce_Enter);
			// 
			// lblAdvalorem
			// 
			this.lblAdvalorem.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblAdvalorem.BackColor = System.Drawing.Color.Transparent;
			this.lblAdvalorem.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsImportacion, "_Importacion.NAdvalorem1"));
			this.lblAdvalorem.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblAdvalorem.ForeColor = System.Drawing.Color.Black;
			this.lblAdvalorem.Location = new System.Drawing.Point(626, 472);
			this.lblAdvalorem.Name = "lblAdvalorem";
			this.lblAdvalorem.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.lblAdvalorem.Size = new System.Drawing.Size(88, 16);
			this.lblAdvalorem.TabIndex = 202;
			this.lblAdvalorem.Text = "ADVALOREM";
			this.lblAdvalorem.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblAdvalorem.Visible = false;
			// 
			// lblFodinfa
			// 
			this.lblFodinfa.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblFodinfa.BackColor = System.Drawing.Color.Transparent;
			this.lblFodinfa.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsImportacion, "_Importacion.NFodinfa"));
			this.lblFodinfa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblFodinfa.ForeColor = System.Drawing.Color.Black;
			this.lblFodinfa.Location = new System.Drawing.Point(626, 496);
			this.lblFodinfa.Name = "lblFodinfa";
			this.lblFodinfa.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.lblFodinfa.Size = new System.Drawing.Size(88, 16);
			this.lblFodinfa.TabIndex = 203;
			this.lblFodinfa.Text = "FODINFA";
			this.lblFodinfa.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblFodinfa.Visible = false;
			this.lblFodinfa.Click += new System.EventHandler(this.lblFodinfa_Click);
			// 
			// lblIce
			// 
			this.lblIce.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblIce.BackColor = System.Drawing.Color.Transparent;
			this.lblIce.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsImportacion, "_Importacion.NCorpei"));
			this.lblIce.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblIce.ForeColor = System.Drawing.Color.Black;
			this.lblIce.Location = new System.Drawing.Point(626, 520);
			this.lblIce.Name = "lblIce";
			this.lblIce.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.lblIce.Size = new System.Drawing.Size(88, 16);
			this.lblIce.TabIndex = 204;
			this.lblIce.Text = "ICE";
			this.lblIce.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblIce.Visible = false;
			this.lblIce.Click += new System.EventHandler(this.label9_Click);
			// 
			// lblIvaImp
			// 
			this.lblIvaImp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblIvaImp.BackColor = System.Drawing.Color.Transparent;
			this.lblIvaImp.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsImportacion, "_Importacion.NDesconsolidacion"));
			this.lblIvaImp.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblIvaImp.ForeColor = System.Drawing.Color.Black;
			this.lblIvaImp.Location = new System.Drawing.Point(626, 544);
			this.lblIvaImp.Name = "lblIvaImp";
			this.lblIvaImp.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.lblIvaImp.Size = new System.Drawing.Size(88, 16);
			this.lblIvaImp.TabIndex = 205;
			this.lblIvaImp.Text = "IVA IMPORT";
			this.lblIvaImp.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblIvaImp.Visible = false;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(412, 64);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(79, 16);
			this.label19.TabIndex = 247;
			this.label19.Text = "DIAS DE VIAJE";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(412, 40);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(74, 16);
			this.label17.TabIndex = 246;
			this.label17.Text = "PRODUCCIÓN";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasViaje
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasViaje.Appearance = appearance25;
			this.txtDiasViaje.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtDiasViaje.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.DiasViaje"));
			this.txtDiasViaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasViaje.FormatString = "";
			this.txtDiasViaje.Location = new System.Drawing.Point(540, 64);
			this.txtDiasViaje.Name = "txtDiasViaje";
			this.txtDiasViaje.PromptChar = ' ';
			this.txtDiasViaje.Size = new System.Drawing.Size(80, 20);
			this.txtDiasViaje.TabIndex = 6;
			this.txtDiasViaje.Click += new System.EventHandler(this.txtDiasViaje_Click);
			this.txtDiasViaje.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDiasViaje_KeyPress);
			this.txtDiasViaje.Validating += new System.ComponentModel.CancelEventHandler(this.txtDiasViaje_Validating);
			this.txtDiasViaje.Validated += new System.EventHandler(this.txtDiasViaje_Validated);
			this.txtDiasViaje.ValueChanged += new System.EventHandler(this.txtDiasViaje_ValueChanged);
			this.txtDiasViaje.Enter += new System.EventHandler(this.txtDiasViaje_Enter);
			// 
			// txtDiasDeProduccion
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasDeProduccion.Appearance = appearance26;
			this.txtDiasDeProduccion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtDiasDeProduccion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.DiasProd"));
			this.txtDiasDeProduccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasDeProduccion.FormatString = "";
			this.txtDiasDeProduccion.Location = new System.Drawing.Point(540, 40);
			this.txtDiasDeProduccion.Name = "txtDiasDeProduccion";
			this.txtDiasDeProduccion.PromptChar = ' ';
			this.txtDiasDeProduccion.Size = new System.Drawing.Size(80, 20);
			this.txtDiasDeProduccion.TabIndex = 3;
			this.txtDiasDeProduccion.Click += new System.EventHandler(this.txtDiasDeProduccion_Click);
			this.txtDiasDeProduccion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDiasDeProduccion_KeyPress);
			this.txtDiasDeProduccion.Validating += new System.ComponentModel.CancelEventHandler(this.txtDiasDeProduccion_Validating);
			this.txtDiasDeProduccion.Validated += new System.EventHandler(this.txtDiasDeProduccion_Validated);
			this.txtDiasDeProduccion.ValueChanged += new System.EventHandler(this.txtDiasDeProduccion_ValueChanged);
			this.txtDiasDeProduccion.Enter += new System.EventHandler(this.txtDiasDeProduccion_Enter);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(116, 88);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(105, 16);
			this.label11.TabIndex = 243;
			this.label11.Text = "DESADUANIZACIÓN";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesAduanizacion
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesAduanizacion.Appearance = appearance27;
			this.dtDesAduanizacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtDesAduanizacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.FechaDesAduaniz"));
			dateButton1.Caption = "Today";
			this.dtDesAduanizacion.DateButtons.Add(dateButton1);
			this.dtDesAduanizacion.Format = "dd/MM/yyyy";
			this.dtDesAduanizacion.Location = new System.Drawing.Point(236, 88);
			this.dtDesAduanizacion.Name = "dtDesAduanizacion";
			this.dtDesAduanizacion.NonAutoSizeHeight = 23;
			this.dtDesAduanizacion.Size = new System.Drawing.Size(112, 21);
			this.dtDesAduanizacion.SpinButtonsVisible = true;
			this.dtDesAduanizacion.TabIndex = 8;
			this.dtDesAduanizacion.Value = new System.DateTime(2013, 6, 13, 0, 0, 0, 0);
			this.dtDesAduanizacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtDesAduanizacion_KeyPress);
			this.dtDesAduanizacion.Validating += new System.ComponentModel.CancelEventHandler(this.dtDesAduanizacion_Validating);
			this.dtDesAduanizacion.Validated += new System.EventHandler(this.dtDesAduanizacion_Validated);
			this.dtDesAduanizacion.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtDesAduanizacion_BeforeDropDown);
			this.dtDesAduanizacion.ValueChanged += new System.EventHandler(this.dtDesAduanizacion_ValueChanged);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(724, 64);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(106, 16);
			this.label16.TabIndex = 241;
			this.label16.Text = "LLEGADA A PUERTO";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtLlegadaPuerto
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtLlegadaPuerto.Appearance = appearance28;
			this.dtLlegadaPuerto.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtLlegadaPuerto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.FechaLlegPuerto"));
			dateButton2.Caption = "Today";
			this.dtLlegadaPuerto.DateButtons.Add(dateButton2);
			this.dtLlegadaPuerto.Format = "dd/MM/yyyy";
			this.dtLlegadaPuerto.Location = new System.Drawing.Point(844, 64);
			this.dtLlegadaPuerto.Name = "dtLlegadaPuerto";
			this.dtLlegadaPuerto.NonAutoSizeHeight = 23;
			this.dtLlegadaPuerto.Size = new System.Drawing.Size(112, 21);
			this.dtLlegadaPuerto.SpinButtonsVisible = true;
			this.dtLlegadaPuerto.TabIndex = 7;
			this.dtLlegadaPuerto.Value = new System.DateTime(2013, 6, 13, 0, 0, 0, 0);
			this.dtLlegadaPuerto.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtLlegadaPuerto_BeforeDropDown);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(116, 64);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(62, 16);
			this.label9.TabIndex = 239;
			this.label9.Text = "EMBARQUE";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaEmbarque
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEmbarque.Appearance = appearance29;
			this.dtFechaEmbarque.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaEmbarque.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.FechaEmb"));
			dateButton3.Caption = "Today";
			this.dtFechaEmbarque.DateButtons.Add(dateButton3);
			this.dtFechaEmbarque.Format = "dd/MM/yyyy";
			this.dtFechaEmbarque.Location = new System.Drawing.Point(236, 64);
			this.dtFechaEmbarque.Name = "dtFechaEmbarque";
			this.dtFechaEmbarque.NonAutoSizeHeight = 23;
			this.dtFechaEmbarque.Size = new System.Drawing.Size(112, 21);
			this.dtFechaEmbarque.SpinButtonsVisible = true;
			this.dtFechaEmbarque.TabIndex = 5;
			this.dtFechaEmbarque.Value = new System.DateTime(2013, 6, 13, 0, 0, 0, 0);
			this.dtFechaEmbarque.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechaEmbarque_KeyPress);
			this.dtFechaEmbarque.Validating += new System.ComponentModel.CancelEventHandler(this.dtFechaEmbarque_Validating);
			this.dtFechaEmbarque.Validated += new System.EventHandler(this.dtFechaEmbarque_Validated);
			this.dtFechaEmbarque.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaEmbarque_BeforeDropDown);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(724, 40);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(22, 16);
			this.label10.TabIndex = 237;
			this.label10.Text = "FIN";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaFin
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaFin.Appearance = appearance30;
			this.dtFechaFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFechaFin.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.FechaFin"));
			dateButton4.Caption = "Today";
			this.dtFechaFin.DateButtons.Add(dateButton4);
			this.dtFechaFin.Format = "dd/MM/yyyy";
			this.dtFechaFin.Location = new System.Drawing.Point(844, 40);
			this.dtFechaFin.Name = "dtFechaFin";
			this.dtFechaFin.NonAutoSizeHeight = 23;
			this.dtFechaFin.Size = new System.Drawing.Size(112, 21);
			this.dtFechaFin.SpinButtonsVisible = true;
			this.dtFechaFin.TabIndex = 4;
			this.dtFechaFin.Value = new System.DateTime(2013, 6, 13, 0, 0, 0, 0);
			this.dtFechaFin.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaFin_BeforeDropDown);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(412, 88);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(115, 16);
			this.label8.TabIndex = 235;
			this.label8.Text = "INGRESO DE BODEGA";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtIngresoBodega
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtIngresoBodega.Appearance = appearance31;
			this.dtIngresoBodega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtIngresoBodega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.FechaIngresoBodega"));
			dateButton5.Caption = "Today";
			this.dtIngresoBodega.DateButtons.Add(dateButton5);
			this.dtIngresoBodega.Format = "dd/MM/yyyy";
			this.dtIngresoBodega.Location = new System.Drawing.Point(540, 88);
			this.dtIngresoBodega.Name = "dtIngresoBodega";
			this.dtIngresoBodega.NonAutoSizeHeight = 23;
			this.dtIngresoBodega.Size = new System.Drawing.Size(112, 21);
			this.dtIngresoBodega.SpinButtonsVisible = true;
			this.dtIngresoBodega.TabIndex = 9;
			this.dtIngresoBodega.Value = new System.DateTime(2013, 6, 13, 0, 0, 0, 0);
			this.dtIngresoBodega.Validating += new System.ComponentModel.CancelEventHandler(this.dtIngresoBodega_Validating);
			this.dtIngresoBodega.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtIngresoBodega_BeforeDropDown);
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Location = new System.Drawing.Point(116, 40);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(39, 16);
			this.label20.TabIndex = 233;
			this.label20.Text = "INICIO";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance32;
			this.dtFecha.AutoSelect = false;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.FechaIni"));
			dateButton6.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton6);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(236, 40);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 2;
			this.dtFecha.Value = new System.DateTime(2013, 6, 18, 0, 0, 0, 0);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// ultraNumericEditor11
			// 
			this.ultraNumericEditor11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraNumericEditor11.Appearance = appearance33;
			this.ultraNumericEditor11.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.CantidadTotal"));
			this.ultraNumericEditor11.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraNumericEditor11.FormatString = "#,##0.00";
			this.ultraNumericEditor11.Location = new System.Drawing.Point(912, 488);
			this.ultraNumericEditor11.Name = "ultraNumericEditor11";
			this.ultraNumericEditor11.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor11.PromptChar = ' ';
			this.ultraNumericEditor11.Size = new System.Drawing.Size(96, 22);
			this.ultraNumericEditor11.TabIndex = 249;
			this.ultraNumericEditor11.Visible = false;
			// 
			// label21
			// 
			this.label21.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(776, 488);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(116, 16);
			this.label21.TabIndex = 248;
			this.label21.Text = "TOTAL IMPORTACIÓN";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label21.Visible = false;
			// 
			// txtSalidaCapitales
			// 
			this.txtSalidaCapitales.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSalidaCapitales.Appearance = appearance34;
			this.txtSalidaCapitales.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsImportacion, "_Importacion.ImpSalCap"));
			this.txtSalidaCapitales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSalidaCapitales.FormatString = "#,##0.00";
			this.txtSalidaCapitales.Location = new System.Drawing.Point(912, 520);
			this.txtSalidaCapitales.Name = "txtSalidaCapitales";
			this.txtSalidaCapitales.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSalidaCapitales.PromptChar = ' ';
			this.txtSalidaCapitales.Size = new System.Drawing.Size(96, 22);
			this.txtSalidaCapitales.TabIndex = 15;
			// 
			// label22
			// 
			this.label22.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(776, 520);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(122, 16);
			this.label22.TabIndex = 250;
			this.label22.Text = "5% SALIDA CAPITALES";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.Enabled = false;
			this.btGrabar.Location = new System.Drawing.Point(192, 600);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(72, 24);
			this.btGrabar.TabIndex = 15;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btEditar
			// 
			this.btEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btEditar.Location = new System.Drawing.Point(272, 600);
			this.btEditar.Name = "btEditar";
			this.btEditar.Size = new System.Drawing.Size(72, 24);
			this.btEditar.TabIndex = 1;
			this.btEditar.Text = "&Editar";
			this.btEditar.Click += new System.EventHandler(this.btEditar_Click);
			// 
			// btDeshacer
			// 
			this.btDeshacer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btDeshacer.CausesValidation = false;
			this.btDeshacer.Enabled = false;
			this.btDeshacer.Location = new System.Drawing.Point(432, 600);
			this.btDeshacer.Name = "btDeshacer";
			this.btDeshacer.Size = new System.Drawing.Size(72, 24);
			this.btDeshacer.TabIndex = 16;
			this.btDeshacer.Text = "&Deshacer";
			this.btDeshacer.Click += new System.EventHandler(this.btDeshacer_Click);
			// 
			// btSalir
			// 
			this.btSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSalir.CausesValidation = false;
			this.btSalir.Location = new System.Drawing.Point(512, 600);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(72, 24);
			this.btSalir.TabIndex = 17;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// btResumenImp
			// 
			this.btResumenImp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btResumenImp.Location = new System.Drawing.Point(352, 600);
			this.btResumenImp.Name = "btResumenImp";
			this.btResumenImp.Size = new System.Drawing.Size(72, 24);
			this.btResumenImp.TabIndex = 251;
			this.btResumenImp.Text = "&Resumén";
			this.btResumenImp.Click += new System.EventHandler(this.btResumenImp_Click);
			// 
			// cmbTipoCalculo
			// 
			this.cmbTipoCalculo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoCalculo.DataSource = this.ultraDataSource1;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Width = 300;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			this.cmbTipoCalculo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoCalculo.DisplayMember = "Calculo";
			this.cmbTipoCalculo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoCalculo.Location = new System.Drawing.Point(728, 272);
			this.cmbTipoCalculo.Name = "cmbTipoCalculo";
			this.cmbTipoCalculo.Size = new System.Drawing.Size(312, 104);
			this.cmbTipoCalculo.TabIndex = 252;
			this.cmbTipoCalculo.ValueMember = "idTipoCalculo";
			this.cmbTipoCalculo.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
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
			// Importacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1072, 630);
			this.Controls.Add(this.cmbTipoCalculo);
			this.Controls.Add(this.btResumenImp);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btDeshacer);
			this.Controls.Add(this.btEditar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.txtSalidaCapitales);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.lbl2);
			this.Controls.Add(this.lbl1);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.ultraNumericEditor11);
			this.Controls.Add(this.txtDiasViaje);
			this.Controls.Add(this.txtDiasDeProduccion);
			this.Controls.Add(this.dtDesAduanizacion);
			this.Controls.Add(this.dtLlegadaPuerto);
			this.Controls.Add(this.dtFechaEmbarque);
			this.Controls.Add(this.dtFechaFin);
			this.Controls.Add(this.dtIngresoBodega);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.lblIvaImp);
			this.Controls.Add(this.lblIce);
			this.Controls.Add(this.lblFodinfa);
			this.Controls.Add(this.lblAdvalorem);
			this.Controls.Add(this.txtIce);
			this.Controls.Add(this.btAprobacion);
			this.Controls.Add(this.btProgramacion);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.btCrearAsientos);
			this.Controls.Add(this.btCalcular);
			this.Controls.Add(this.ultraNumericEditor8);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtFodInfa);
			this.Controls.Add(this.txtAdValorem);
			this.Controls.Add(this.ultraNumericEditor4);
			this.Controls.Add(this.ultraNumericEditor3);
			this.Controls.Add(this.ultraNumericEditor2);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.txtPrecio1);
			this.Controls.Add(this.btImportacionTipo);
			this.Controls.Add(this.cmbTipoImportacion);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btImportacion);
			this.Controls.Add(this.barraDato1);
			this.KeyPreview = true;
			this.Name = "Importacion";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Importación";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Importacion_KeyDown);
			this.Load += new System.EventHandler(this.Importacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsImportacionNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdValorem)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFodInfa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIce)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasViaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasDeProduccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesAduanizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLlegadaPuerto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEmbarque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIngresoBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalidaCapitales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCalculo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private C1DataRow drSeteoF;
		CultureInfo us = new CultureInfo("en-US");

		private void Importacion_Load(object sender, System.EventArgs e)
		{
			Funcion.Resolucion(this, this.Height, this.Width);

			this.barraDato1.HabilitaControles(false);
			this.barraDato1.PosUltima();
			this.cmbProyecto.Enabled=true;  
			this.ultraGrid1.Enabled=true;  
			//ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;	
			this.cmbTipoCalculo.DataSource = Funcion.dvProcedimiento(cdsImportacion, "Select idTipoCalculo, Calculo From TiposDeCalculo");
		}

		public void Texto()
		{
			if (drImportacion == null) return;
            //bsImportacionDet.DataSource = ef.ImportacionDet.Where(it => it.idImportacion == drImportacion.idImportacion).OrderBy(it => it.idImportacionDet);
      int can = Funcion.iEscalarSQL(cdsImportacion , "Select Count(*) From ImportacionDet Where NoObligatorio = 1 And idCompra is null And idImportacion = " + drImportacion["idImportacion"].ToString());
			
			if (can == 0)
			{
				lbl1 .Text = "IMPORTACIÓN COMPLETA";
				drImportacion["Estado"] = 1;
				this.lbl1.ForeColor = Color.Blue;
			}
			else 
			{
				lbl1.Text = "IMPORTACIÓN CON VALORES ESTIMADOS";
				this.lbl1.ForeColor = Color.Red;
				drImportacion["Estado"] = 0;
			}
				
      string stSelect = "select isnull(sum(saldo),0) from compra where borrar=0 and idcompra in (select idcompra from importaciondet where idcompra is not null and idimportacion=" + drImportacion["idImportacion"].ToString() + ") ";
      double dSaldo = Funcion.dEscalarSQL (cdsImportacion, stSelect.ToString() );

			lbl2.Text = "SALDOS DE IMPORTACIÓN = " + dSaldo.ToString();
		}


		private void label8_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label50_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSubProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTipoImportacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void cdsImportacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsImportacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btImportacion_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string  stSelect;
   	//--	stSelect="exec IMPORTACIONCOSTO2 " + cmbProyecto.Value.ToString(); 
			stSelect="exec IMPORTACIONCOSTO2Datos " + cmbProyecto.Value.ToString(); 
		  Funcion.EjecutaSQL(cdsImportacion, stSelect);  

			barraDato1.IdRegistro = (int) cmbProyecto.Value;
			barraDato1.ProximoId(5);
			Cursor = Cursors.Default;
			MessageBox.Show("Datos Importados");  
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			if (this.cmbProyecto.Rows.Count == 0)//.Text.ToString().Length == 0)
			{
				this.btProgramacion.Enabled = false;
				this.btImportacionTipo.Enabled = false;
				this.btFactura.Enabled = false;
				this.btAsiento.Enabled = false;
				this.btEditar.Enabled = false;
				this.barraDato1.Enabled = false;

				MessageBox.Show("No existen importaciones.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			drImportacion = cdsImportacion.TableViews["Importacion"].Rows[0];
			Calculos();
			Texto(); 	
		}

		private void label22_Click(object sender, System.EventArgs e)
		{
		
    }

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
			{
				string CodAsiento =  (string) dr.Cells["CodAsiento"].Value;
				int IdTipoF = Funcion.iEscalarSQL(cdsImportacion ,"select idasiento from asiento where codasiento='" + CodAsiento.ToString()+"'" );  	     
				Asiento miAsiento = new Asiento(IdTipoF);
				miAsiento.MdiParent = MdiParent;
				miAsiento.Show();
				return;
			}
		}

		private void btNombre_Click(object sender, System.EventArgs e)
		{

		}

		private void cdsImportacionNombre_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsImportacionNombre.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsTipoImportacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTipoImportacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					#region Flechas
					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
						#endregion Flechas
					#region Enter
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Tecla");
			}
		}

		private void btImportacionTipo_Click(object sender, System.EventArgs e)
		{
			using (ImportacionTipo miImportacionTipo = new ImportacionTipo())
			{
				miImportacionTipo.ShowDialog();

				if (this.cmbProyecto.Text.ToString().Length == 0) return;

				cdsTipoImportacion.Clear();
				cdsTipoImportacion.Fill();
				cmbTipoImportacion.Refresh();
				if ((int)drImportacion["Estado"] == 0)			
				{
					string sSelect = "Exec ImportacionCosto2Datos " + cmbProyecto.Value.ToString(); 
					Funcion.EjecutaSQL(cdsImportacion, sSelect);  
				}
				barraDato1.IdRegistro = (int) cmbProyecto.Value;
				barraDato1.ProximoId(5);

				Calculos();
			}
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{

		}

		private void label7_Click(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
//			string stFiltro = "{Importacion.idImportacion} = " + cmbProyecto.Value.ToString();
//			Reporte miRep = new Reporte("Importacion.rpt", stFiltro);
//			miRep.MdiParent = MdiParent;
//			miRep.Show();

			frmReporteImportacion miRepImp = new frmReporteImportacion((int)this.cmbProyecto.Value);
			miRepImp.ShowDialog();
			Cursor = Cursors.Default;	
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void btCalcular_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string  stSelect;
			//--	stSelect="exec IMPORTACIONCOSTO2 " + cmbProyecto.Value.ToString(); 
			stSelect="exec IMPORTACIONcosto2 " + cmbProyecto.Value.ToString(); 
			Funcion.EjecutaSQL(cdsImportacion, stSelect);  

			barraDato1.IdRegistro = (int) cmbProyecto.Value;
			barraDato1.ProximoId(5);
			Cursor = Cursors.Default;
			MessageBox.Show("Importacion Calculada");  
		}

		private void btCrearAsientos_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string  stSelect;
			stSelect="exec IMPORTACIONAsiento2 " + cmbProyecto.Value.ToString(); 
			Funcion.EjecutaSQL(cdsImportacion, stSelect);  

			barraDato1.IdRegistro = (int) cmbProyecto.Value;
			barraDato1.ProximoId(5);
			Cursor = Cursors.Default;
			MessageBox.Show("Asientos de Importacion Generados");  
		
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
			{
				int IdCompra =  (int) dr.Cells["idCompra"].Value;
				int IdTipoF = Funcion.iEscalarSQL(cdsImportacion ,"select idtipofactura from compra where idcompra=" + IdCompra.ToString());  	     
//				Compra miCompra = new Compra(IdTipoF, IdCompra);
//				miCompra.MdiParent = MdiParent;
//				miCompra.Show();
				return;
			}
		}

		private void btProgramacion_Click(object sender, System.EventArgs e)
		{	
			if (drImportacion == null)
			{
				MessageBox.Show("No existen Importaciones.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			iTipoCliente = Funcion.iEscalarSQL(cdsImportacion, string.Format("Select dbo.GrupoClienteImportacion ({0})", (int)this.cmbProyecto.Value));

			if (iTipoCliente != 3) 
			{
				MessageBox.Show("Programación de pagos solo para Proveedores del Asia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (drImportacion["FechaIni"] == System.DBNull.Value)	
			{
				MessageBox.Show("Ingrese las fechas de la Importación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (this.barraDato1.bEditar)
			{
				MessageBox.Show("No puede programar pagos en edición.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			using (ImportacionProgramacion miImportacionProgramacion = new ImportacionProgramacion())
			{
				miImportacionProgramacion.ShowDialog();


			}
		}

		private void btAprobacion_Click(object sender, System.EventArgs e)
		{
			using (ImportacionAprobacion miImportacionAprobacion = new ImportacionAprobacion())
			{
				miImportacionAprobacion.ShowDialog();
			}
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();  
			
			ModificaGrilla(true);

			this.cmbProyecto.Enabled = false;

			if (iTipoCliente == 2) //INTERNACIONAL
			{
				this.txtDiasDeProduccion.Enabled = false;
			}
			
			this.dtFechaFin.Enabled = false;
			this.dtLlegadaPuerto.Enabled = false;
			this.txtPrecio1.Enabled = false;
			this.ultraNumericEditor1.Enabled = false;
			this.ultraNumericEditor2.Enabled = false;
			this.ultraNumericEditor3.Enabled = false;
			this.ultraNumericEditor4.Enabled = false;
			this.txtSalidaCapitales.Enabled = false;
			this.ultraNumericEditor8.Enabled = false;

			this.dtFecha.Select();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (drImportacion["FechaIni"] == DBNull.Value)
			{
				MessageBox.Show("Escriba la fecha de inicio de la importación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.dtFecha.Select();
				return;
			}

			Calculos();

			Texto(); 

			this.barraDato1.GrabaRegistro();  			
	
			barraDato1.IdRegistro = (int) cmbProyecto.Value;
			barraDato1.ProximoId(5); 

			string sSQL = string.Format("Exec ReporteImportacion {0}", (int)this.cmbProyecto.Value);
			Funcion.EjecutaSQL(cdsImportacion, sSQL);  

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((int)dr.Cells["idTipoImportacion"].Value == 1)
				{
					string stExec = string.Format("Exec ArticuloActualizaPrecio {0}, 1", (int)dr.Cells["idCompra"].Value); //this.txtIdCompra.Value, (int)dr.Cells["idDetCompra"].Value, , stVerifica
					Funcion.EjecutaSQL(cdsImportacion, stExec, true);

					DateTime dtLlegada = (DateTime) drImportacion["FechaIngresoBodega"];

					Funcion.EjecutaSQL(cdsImportacion, string.Format("Update Compra Set FechaTentativaLlegada = '{0}' Where idTipoFactura = 2 And idCompra = {1}", dtLlegada.ToString("yyyyMMdd"), (int)dr.Cells["idCompra"].Value));
				}
			}

			ModificaGrilla(false);

			string sSelect = "Exec IMPORTACIONAsiento2 " + cmbProyecto.Value.ToString(); 
			Funcion.EjecutaSQL(cdsImportacion, sSelect);  

			barraDato1.IdRegistro = (int) cmbProyecto.Value;
			barraDato1.ProximoId(5);

			this.cmbProyecto.Enabled = true;
			Cursor = Cursors.Default;

		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();  
			ModificaGrilla(false);
			this.cmbProyecto.Enabled = true;
		}

		int idFlete = 0;
		int iTipoCliente = 0;

		private void cmbProyecto_RowSelected(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
		{
			Cursor = Cursors.WaitCursor;

			//Datos();

			barraDato1.IdRegistro = (int) cmbProyecto.Value;
			barraDato1.ProximoId(5);

			string sSQL = string.Format("Select idFlete From ImportacionNombre Where idImportacionNombre = {0}", (int) cmbProyecto.Value);
			idFlete = Funcion.iEscalarSQL(cdsImportacion, sSQL);

			string sSQLGrupo = string.Format("Select dbo.GrupoClienteImportacion ({0})", (int)this.cmbProyecto.Value);
			iTipoCliente = Funcion.iEscalarSQL(cdsImportacion, sSQLGrupo);

			if (iTipoCliente == 2) this.btProgramacion.Enabled = false;
			else if (iTipoCliente == 3) this.btProgramacion.Enabled = true;
//					
			//			if ((int)drImportacion["Estado"] == 0)			
			//			{
			string sSelect = "Exec ImportacionCosto2Datos " + cmbProyecto.Value.ToString(); 
			Funcion.EjecutaSQL(cdsImportacion, sSelect);  
			//			}
			//			stSelect="exec IMPORTACIONcosto2 " + cmbProyecto.Value.ToString(); 
			//			Funcion.EjecutaSQL(cdsImportacion, stSelect);  

			ModificaGrilla(false);

			barraDato1.IdRegistro = (int) cmbProyecto.Value;
			barraDato1.ProximoId(5);
			
			Calculos();

			//this.dtFecha.Select();
			Cursor = Cursors.Default;
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idTipoCalculo"].CellActivation = Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Adicional"].CellActivation = Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["NoObligatorio"].CellActivation = Activation.AllowEdit;
			}
			else
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idTipoCalculo"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Adicional"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["NoObligatorio"].CellActivation = Activation.Disabled;
			}
		}

		private void label9_Click(object sender, System.EventArgs e)
		{
		
		}

		private void lblFodinfa_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbProyecto_Validated(object sender, System.EventArgs e)
		{

		
		}

		private void cmbProyecto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void Calculos()
		{
			drImportacion["Fob"] = 0;
			drImportacion["FobExw"] = 0; 
			drImportacion["FleteExt"] = 0;
			drImportacion["Seguro"] = 0; 
			drImportacion["ImpSalCap"] = 0; 
			drImportacion["TotalG"] = 0; 
				
			string sSQL = "";
			double dImpSalCap = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((int)dr.Cells["idTipoImportacion"].Value == 1)
				{
					drImportacion["Fob"] = (double) drImportacion["Fob"] + (double)dr.Cells["TotalS"].Value;		 
				}

				if ((int)dr.Cells["idTipoImportacion"].Value == 2)
				{
					drImportacion["FobExw"] = (double) drImportacion["FobExw"] + (double)dr.Cells["TotalS"].Value;		 
				}
					
				if ((int)dr.Cells["idTipoImportacion"].Value == 3)
				{
					drImportacion["FleteExt"] = (double) drImportacion["FleteExt"] + (double)dr.Cells["TotalS"].Value;		 
				}

				if ((int)dr.Cells["idTipoImportacion"].Value == 4)
				{
					drImportacion["Seguro"] = (double) drImportacion["Seguro"] + (double)dr.Cells["TotalS"].Value;		 
				}
				
				sSQL = string.Format("Select Exterior From ImportacionTipo Where idImportacionTipo = {0}", (int)dr.Cells["idTipoImportacion"].Value);
				bool bExterior = Funcion.bEscalarSQL(cdsImportacion, sSQL, true);
				if (bExterior)
				{
					dImpSalCap = 0;
					dImpSalCap = (double)dr.Cells["TotalS"].Value * 0.05;
					drImportacion["ImpSalCap"] = (double)drImportacion["ImpSalCap"] + dImpSalCap;
				}
			}

			drImportacion["Cif"] = (double) drImportacion["Fob"]
				+ (double)drImportacion["FobExw"]
				+ (double)drImportacion["FleteExt"]
				+ (double)drImportacion["Seguro"]; 

			//drImportacion["ImpSalCap"] = (double) drImportacion["Fob"] * 0.05;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{					
				drImportacion["TotalG"] = (double) drImportacion["TotalG"] + (double)dr.Cells["TotalS"].Value;																	
			}

			drImportacion["TotalG"] = (double) drImportacion["TotalG"] 
				+ (double) drImportacion["Advalorem"] 
				+ (double) drImportacion["FodInfa"] 
				+ (double) drImportacion["Corpei"] 
				+ (double) drImportacion["IVAIMPOR"] 
				+ (double) drImportacion["ImpSalCap"];
		}

		private void Calculo(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			e.Cell.Row.Cells["TotalS"].Value = (double) e.Cell.Row.Cells["Adicional"].Value + (double)e.Cell.Row.Cells["Total"].Value;
				
			drImportacion["Fob"] = 0;
			drImportacion["FobExw"] = 0; 
			drImportacion["FleteExt"] = 0;
			drImportacion["Seguro"] = 0; 
			drImportacion["ImpSalCap"] = 0; 
			drImportacion["TotalG"] = 0; 
				
			string sSQL = "";
			double dImpSalCap = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((int)dr.Cells["idTipoImportacion"].Value == 1)
				{
					drImportacion["Fob"] = (double) drImportacion["Fob"] + (double)dr.Cells["TotalS"].Value;		 
				}

				if ((int)dr.Cells["idTipoImportacion"].Value == 2)
				{
					drImportacion["FobExw"] = (double) drImportacion["FobExw"] + (double)dr.Cells["TotalS"].Value;		 
				}
					
				if ((int)dr.Cells["idTipoImportacion"].Value == 3)
				{
					drImportacion["FleteExt"] = (double) drImportacion["FleteExt"] + (double)dr.Cells["TotalS"].Value;		 
				}

				if ((int)dr.Cells["idTipoImportacion"].Value == 4)
				{
					drImportacion["Seguro"] = (double) drImportacion["Seguro"] + (double)dr.Cells["TotalS"].Value;		 
				}	
				
				
				sSQL = string.Format("Select Exterior From ImportacionTipo Where idImportacionTipo = {0}", (int)dr.Cells["idTipoImportacion"].Value);
				bool bExterior = Funcion.bEscalarSQL(cdsImportacion, sSQL, true);
				if (bExterior)
				{
					dImpSalCap = 0;
					dImpSalCap = (double)dr.Cells["TotalS"].Value * 0.05;
					drImportacion["ImpSalCap"] = (double)drImportacion["ImpSalCap"] + dImpSalCap;
				}
			}

			drImportacion["Cif"] = (double) drImportacion["Fob"] 
													 + (double)drImportacion["FobExw"] 
													 + (double)drImportacion["FleteExt"]
													 + (double)drImportacion["Seguro"]; 

			//drImportacion["ImpSalCap"] = (double) drImportacion["Fob"] * 0.05;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{					
				drImportacion["TotalG"] = (double) drImportacion["TotalG"] 
																+ (double)dr.Cells["TotalS"].Value;																	
			}

			drImportacion["TotalG"] = (double) drImportacion["TotalG"] 
															+ (double) drImportacion["Advalorem"] 
															+ (double) drImportacion["FodInfa"] 
															+ (double) drImportacion["Corpei"] 
															+ (double) drImportacion["IVAIMPOR"] 
															+ (double) drImportacion["ImpSalCap"];
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Adicional")
			{
				Calculo(e);
//				if ((double)e.Cell.Row.Cells["Adicional"].Value > 0) e.Cell.Row.Cells["NoObligatorio"].Value = true;
//				else e.Cell.Row.Cells["NoObligatorio"].Value = false;
			}
		}

		private void txtDiasDeProduccion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				//if (drImportacion["FechaIni"] == DBNull.Value) return;				
				this.dtFechaEmbarque.Select();
			}
		}

		private void txtDiasDeProduccion_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void dtFechaFin_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFechaEmbarque_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtDiasViaje_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void dtDesAduanizacion_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtDesAduanizacion_ValueChanged(object sender, System.EventArgs e)
		{
						
		}

		private void dtFechaEmbarque_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.txtDiasViaje.Select();
			}
		}

		private void txtDiasViaje_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.dtDesAduanizacion.Select();
			}
		}

		private void dtDesAduanizacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.dtIngresoBodega.Select();
			}
		}

		private void dtLlegadaPuerto_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (iTipoCliente == 2)this.dtFechaEmbarque.Select();	
				else if (iTipoCliente == 3)this.txtDiasDeProduccion.Select();			
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			#region Decimales para el total
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decTotal"]);
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decTotal"]);
			#endregion Formato numerico

			#region Decimales
			e.Layout.Bands[0].Columns["Adicional"].Format = stFormato;
			e.Layout.Bands[0].Columns["Adicional"].MaskInput = stInput;
			#endregion Decimales
		}

		private void txtDiasDeProduccion_Click(object sender, System.EventArgs e)
		{
			this.txtDiasDeProduccion.SelectAll();
		}

		private void txtDiasDeProduccion_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasDeProduccion.SelectAll();
		}

		private void txtDiasViaje_Click(object sender, System.EventArgs e)
		{
			this.txtDiasViaje.SelectAll();
		}

		private void txtDiasViaje_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasViaje.SelectAll();
		}

		private void txtAdValorem_Click(object sender, System.EventArgs e)
		{
			this.txtAdValorem.SelectAll();
		}

		private void txtAdValorem_Enter(object sender, System.EventArgs e)
		{
			this.txtAdValorem.SelectAll();
		}

		private void txtFodInfa_Click(object sender, System.EventArgs e)
		{
			this.txtFodInfa.SelectAll();
		}

		private void txtFodInfa_Enter(object sender, System.EventArgs e)
		{
			this.txtFodInfa.SelectAll();
		}

		private void txtFodInfa_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtAdValorem_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtAdValorem_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.txtFodInfa.Select();
			}
		}

		private void txtFodInfa_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.txtIce.Select();
			}
		}

		private void txtIva_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void txtIva_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.btGrabar.Select();
			}
		}

		private void txtIce_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtIce_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				this.txtIva.Select();
			}
		}

		private void txtAdValorem_Validated(object sender, System.EventArgs e)
		{
			Calculos();
		}

		private void txtFodInfa_Validated(object sender, System.EventArgs e)
		{
			Calculos();
		}

		private void txtIce_Validated(object sender, System.EventArgs e)
		{
			Calculos();
		}

		private void txtIva_Validated(object sender, System.EventArgs e)
		{
			Calculos();
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && this.dtFecha.Text == "")
			{
				MessageBox.Show("Seleccione la fecha de inicio.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.dtFecha.Select();
				e.Cancel = true;
			}
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (!this.barraDato1.bEditar) return;

			if (drImportacion["FechaIni"] == System.DBNull.Value) return;

			drImportacion["FechaIni"] = (DateTime)this.dtFecha.Value;
			if (iTipoCliente == 2)
			{
				drImportacion["DiasProd"] = 0;
				drImportacion["FechaFin"] = (DateTime)this.dtFecha.Value;

				DateTime FInicio = (DateTime)this.dtFecha.Value;
				int DProd = (int)this.txtDiasDeProduccion.Value;
				drImportacion["DiasProd"] = DProd;
				drImportacion["FechaFin"] = FInicio.AddDays(DProd);

				DateTime FFin = (DateTime)this.dtFechaFin.Value;
				this.dtFechaEmbarque.CalendarInfo.MinDate = FFin;
				drImportacion["FechaEmb"] = FFin.AddDays(2);

				if (idFlete == 1) drImportacion["DiasViaje"] = 30;
				else if (idFlete == 2) drImportacion["DiasViaje"] = 8;

				DateTime FEmb = (DateTime)this.dtFechaEmbarque.Value;
				int DiasViaje = (int)drImportacion["DiasViaje"];
				drImportacion["FechaLlegPuerto"] = FEmb.AddDays(DiasViaje);

				DateTime LlegPto = (DateTime)this.dtLlegadaPuerto.Value;
				this.dtDesAduanizacion.CalendarInfo.MinDate = LlegPto;
				drImportacion["FechaDesAduaniz"] = LlegPto.AddDays(5);

				DateTime DesAd = (DateTime)this.dtDesAduanizacion.Value;
				this.dtIngresoBodega.CalendarInfo.MinDate = DesAd;
				drImportacion["FechaIngresoBodega"] = DesAd.AddDays(2);
			}
			else if (iTipoCliente == 3)
			{
				DateTime FInicio = (DateTime)this.dtFecha.Value;
				int DProd = (int)this.txtDiasDeProduccion.Value;
				drImportacion["DiasProd"] = DProd;
				drImportacion["FechaFin"] = FInicio.AddDays(DProd);

				DateTime FFin = (DateTime)this.dtFechaFin.Value;
				this.dtFechaEmbarque.CalendarInfo.MinDate = FFin;
				drImportacion["FechaEmb"] = FFin.AddDays(2);

				if (idFlete == 1) drImportacion["DiasViaje"] = 30;
				else if (idFlete == 2) drImportacion["DiasViaje"] = 8;

				DateTime FEmb = (DateTime)this.dtFechaEmbarque.Value;
				int DiasViaje = (int)drImportacion["DiasViaje"];
				drImportacion["FechaLlegPuerto"] = FEmb.AddDays(DiasViaje);

				DateTime LlegPto = (DateTime)this.dtLlegadaPuerto.Value;
				this.dtDesAduanizacion.CalendarInfo.MinDate = LlegPto;
				drImportacion["FechaDesAduaniz"] = LlegPto.AddDays(5);

				DateTime DesAd = (DateTime)this.dtDesAduanizacion.Value;
				this.dtIngresoBodega.CalendarInfo.MinDate = DesAd;
				drImportacion["FechaIngresoBodega"] = DesAd.AddDays(2);
			}			
		}

		private void txtDiasDeProduccion_Validated(object sender, System.EventArgs e)
		{
			this.dtFecha_Validated(sender, e);
		}

		private void txtDiasDeProduccion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && drImportacion["DiasProd"] == DBNull.Value)
			{
				MessageBox.Show("Ingrese los dias de producción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDiasDeProduccion.Select();
				e.Cancel = true;
			}
		}

		private void dtFechaEmbarque_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && this.dtFechaEmbarque.Text == "")
			{
				MessageBox.Show("Seleccione la fecha de embarque.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.dtFechaEmbarque.Select();
				e.Cancel = true;
			}
		}

		private void dtDesAduanizacion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && this.dtDesAduanizacion.Text == "")
			{
				MessageBox.Show("Seleccione la fecha de desaduanización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.dtDesAduanizacion.Select();
				e.Cancel = true;
			}
		}

		private void dtIngresoBodega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && this.dtIngresoBodega.Text == "")
			{
				MessageBox.Show("Seleccione la fecha de ingreso a bodega de la mercaderia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.dtIngresoBodega.Select();
				e.Cancel = true;
			}
		}

		private void txtDiasViaje_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.barraDato1.bEditar && drImportacion["DiasViaje"] == DBNull.Value)
			{
				MessageBox.Show("Ingrese los dias de viaje", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDiasViaje.Select();
				e.Cancel = true;
			}
		}

		private void dtFechaEmbarque_Validated(object sender, System.EventArgs e)
		{
			if (!this.barraDato1.bEditar) return;

			if (drImportacion["FechaEmb"] == System.DBNull.Value) return;

      if (idFlete == 1) drImportacion["DiasViaje"] = 30;
			else if (idFlete == 2) drImportacion["DiasViaje"] = 8;
			
			drImportacion["FechaEmb"] = (DateTime)this.dtFechaEmbarque.Value;
			
			DateTime FEmb = (DateTime)this.dtFechaEmbarque.Value;
			int DiasViaje = (int)drImportacion["DiasViaje"];
			
			drImportacion["FechaLlegPuerto"] = FEmb.AddDays(DiasViaje);
			DateTime LlegPto = (DateTime)this.dtLlegadaPuerto.Value;
			this.dtDesAduanizacion.CalendarInfo.MinDate = LlegPto;
			drImportacion["FechaDesAduaniz"] = LlegPto.AddDays(5);
			
			DateTime DesAd = (DateTime)this.dtDesAduanizacion.Value;
			this.dtIngresoBodega.CalendarInfo.MinDate = DesAd;
			drImportacion["FechaIngresoBodega"] = DesAd.AddDays(2);
		}

		private void txtDiasViaje_Validated(object sender, System.EventArgs e)
		{
			if (!this.barraDato1.bEditar) return;

			DateTime FEmb = (DateTime)this.dtFechaEmbarque.Value;
			int DiasViaje = (int)this.txtDiasViaje.Value; // (int)drImportacion["DiasViaje"];
			drImportacion["DiasViaje"] = DiasViaje;
			drImportacion["FechaLlegPuerto"] = FEmb.AddDays(DiasViaje);

			DateTime LlegPto = (DateTime)this.dtLlegadaPuerto.Value;
			this.dtDesAduanizacion.CalendarInfo.MinDate = LlegPto;	
			drImportacion["FechaDesAduaniz"] = LlegPto.AddDays(5);

			DateTime DesAd = (DateTime)this.dtDesAduanizacion.Value;
			this.dtIngresoBodega.CalendarInfo.MinDate = DesAd;
			drImportacion["FechaIngresoBodega"] = DesAd.AddDays(2);
		}

		private void dtDesAduanizacion_Validated(object sender, System.EventArgs e)
		{
			if (!this.barraDato1.bEditar) return;

			if (drImportacion["FechaDesAduaniz"] == System.DBNull.Value) return;

			DateTime DesAd = (DateTime)this.dtDesAduanizacion.Value;
			drImportacion["FechaDesAduaniz"] = DesAd;
			this.dtIngresoBodega.CalendarInfo.MinDate = DesAd;
			drImportacion["FechaIngresoBodega"] = DesAd.AddDays(2);		
		}

		private void txtIce_Click(object sender, System.EventArgs e)
		{
			this.txtIce.SelectAll();
		}

		private void txtIce_Enter(object sender, System.EventArgs e)
		{
			this.txtIce.SelectAll();
		}

		private void txtIva_Enter(object sender, System.EventArgs e)
		{
			this.txtIva.SelectAll();
		}

		private void txtIva_Click(object sender, System.EventArgs e)
		{
			this.txtIva.SelectAll();
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			this.barraDato1_Grabar(sender, e);
			this.btGrabar.Enabled = false;
			this.btDeshacer.Enabled = false;
			this.btEditar.Enabled = true;
			this.btResumenImp.Enabled = true;
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btEditar_Click(object sender, System.EventArgs e)
		{
			if (this.cmbProyecto.Text.ToString().Length == 0) return;
      this.barraDato1_Editar(sender, e);
			string sSQLGrupo = string.Format("Select dbo.GrupoClienteImportacion ({0})", (int)this.cmbProyecto.Value);
			iTipoCliente = Funcion.iEscalarSQL(cdsImportacion, sSQLGrupo);
			this.btGrabar.Enabled = true;
			this.btDeshacer.Enabled = true;
			this.btEditar.Enabled = false;
			this.btResumenImp.Enabled = false;
		}

		private void btDeshacer_Click(object sender, System.EventArgs e)
		{
			this.barraDato1_Deshacer(sender, e);
			this.btGrabar.Enabled = false;
			this.btDeshacer.Enabled = false;
			this.btEditar.Enabled = true;
			this.btResumenImp.Enabled = true;
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void Importacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) 
			{
				if (this.barraDato1.bEditar) this.btDeshacer_Click(sender, e); // this.barraDato1_Deshacer(sender, e);
				else Close();
			}			
			else if (e.KeyCode == Keys.E)
			{
				if (!this.barraDato1.bEditar) this.btEditar_Click(sender, e); // this.barraDato1_Editar(sender, e);
			}			
			else if (e.KeyCode == Keys.F12)
			{
				if (this.barraDato1.bEditar) this.btGrabar_Click(sender, e); // this.barraDato1_Grabar(sender, e);
			}
			else if (e.KeyCode == Keys.I)
			{
				if (!this.barraDato1.bEditar) this.btResumenImp_Click(sender, e); // this.barraDato1_Grabar(sender, e);
			}
			else if (e.KeyCode == Keys.F5)
			{
				if (!this.barraDato1.bEditar)
				{
					#region Actualizar 
					Cursor = Cursors.WaitCursor;

					if (this.cmbProyecto.Text.ToString().Length == 0) return;

					barraDato1.IdRegistro = (int) cmbProyecto.Value;
					barraDato1.ProximoId(5);

					string sSQL = string.Format("Select idFlete From ImportacionNombre Where idImportacionNombre = {0}", (int) cmbProyecto.Value);
					idFlete = Funcion.iEscalarSQL(cdsImportacion, sSQL);

					string sSQLGrupo = string.Format("Select dbo.GrupoClienteImportacion ({0})", (int)this.cmbProyecto.Value);
					iTipoCliente = Funcion.iEscalarSQL(cdsImportacion, sSQLGrupo);

					if (iTipoCliente == 2) this.btProgramacion.Enabled = false;
					else if (iTipoCliente == 3) this.btProgramacion.Enabled = true;
					
					string sSelect = "Exec ImportacionCosto2Datos " + cmbProyecto.Value.ToString(); 
					Funcion.EjecutaSQL(cdsImportacion, sSelect);  
					
					ModificaGrilla(false);

					barraDato1.IdRegistro = (int) cmbProyecto.Value;
					barraDato1.ProximoId(5);
			
					Calculos();

					Cursor = Cursors.Default;
					#endregion Actualizar

					this.cdsImportacionNombre.Clear();
					this.cdsImportacionNombre.Fill();	
				}
			}			
		}

		private void btResumenImp_Click(object sender, System.EventArgs e)
		{
			if (this.barraDato1.bEditar) return;
			this.barraDato1_Imprimir(sender, e);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void dtIngresoBodega_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}
	}
}
