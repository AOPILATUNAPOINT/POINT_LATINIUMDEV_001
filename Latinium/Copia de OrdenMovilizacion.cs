using System;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de OrdenMovilizacion.
	/// </summary>
	public class OrdenMovilizacion : System.Windows.Forms.Form
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCabezal;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorGalon;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private C1.Data.C1DataSet cdsOrdenMov;
		private C1.Data.C1DataSet cdsArticulo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.Misc.UltraLabel ultraLabel9;
		private Infragistics.Win.Misc.UltraLabel ultraLabel10;
		private Infragistics.Win.Misc.UltraLabel ultraLabel11;
		private Infragistics.Win.Misc.UltraLabel ultraLabel12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor9;
		private Infragistics.Win.Misc.UltraLabel ultraLabel13;
		private Infragistics.Win.Misc.UltraLabel ultraLabel14;
		private Infragistics.Win.Misc.UltraLabel ultraLabel15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel16;
		private Infragistics.Win.Misc.UltraLabel ultraLabel17;
		private Infragistics.Win.Misc.UltraLabel ultraLabel18;
		private Infragistics.Win.Misc.UltraLabel ultraLabel19;
		private Infragistics.Win.Misc.UltraLabel ultraLabel20;
		private Infragistics.Win.Misc.UltraLabel ultraLabel21;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel22;
		private Infragistics.Win.Misc.UltraLabel ultraLabel23;
		private Infragistics.Win.Misc.UltraLabel ultraLabel24;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel25;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor18;
		private Infragistics.Win.Misc.UltraLabel ultraLabel26;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ultraCheckEditor3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel27;
		private Infragistics.Win.Misc.UltraLabel ultraLabel32;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArticulo;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCarro;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor17;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor19;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor8;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor14;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor16;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor17;
		private System.Windows.Forms.Button btPlaca;
		private Infragistics.Win.Misc.UltraLabel ultraLabel28;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdMovilizacion;
		private Infragistics.Win.Misc.UltraLabel ultraLabel29;
		private Infragistics.Win.Misc.UltraLabel ultraLabel30;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor19;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor ultraNumericEditor20;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFurgon;
		private Infragistics.Win.Misc.UltraLabel ultraLabel31;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCodigo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEquipo;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel33;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo5;
		private Infragistics.Win.Misc.UltraLabel ultraLabel34;
		private Infragistics.Win.Misc.UltraLabel ultraLabel35;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public OrdenMovilizacion()
		{
			InitializeComponent();
		}

		int IdArticulo = 0;
		public OrdenMovilizacion(int idArticulo)
		{
			InitializeComponent();
			IdArticulo = idArticulo;
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Carros", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCarro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Placa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tamanio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Licencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Millas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Equipo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Carros", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCarro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Placa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tamanio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Licencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Millas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Equipo");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton6 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsOrdenMov = new C1.Data.C1DataSet();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtCabezal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.txtValorGalon = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraCheckEditor1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel9 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel10 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel11 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel12 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor9 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel13 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel14 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel15 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel16 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel17 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel18 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel19 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel20 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel21 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor3 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel22 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel23 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel24 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor4 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel25 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor18 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel26 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraCheckEditor2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraCheckEditor3 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraLabel27 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel32 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbArticulo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCalendarCombo1 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraCalendarCombo2 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbCarro = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtFurgon = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor17 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraTextEditor19 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraNumericEditor5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor6 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor7 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraNumericEditor2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor8 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor9 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor10 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor11 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor12 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor13 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor14 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor15 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor16 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor17 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btPlaca = new System.Windows.Forms.Button();
			this.ultraLabel28 = new Infragistics.Win.Misc.UltraLabel();
			this.txtIdMovilizacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel29 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel30 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraNumericEditor18 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor19 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraNumericEditor20 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbCodigo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraLabel31 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbEquipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.ultraCalendarCombo3 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel33 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraCalendarCombo4 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraCalendarCombo5 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel34 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel35 = new Infragistics.Win.Misc.UltraLabel();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsOrdenMov)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCabezal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorGalon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor18)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCarro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFurgon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor17)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor19)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor15)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor16)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor17)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdMovilizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor18)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor19)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEquipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo5)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(24, 8);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(36, 14);
			this.ultraLabel1.TabIndex = 0;
			this.ultraLabel1.Text = "Fecha";
			// 
			// cmbFecha
			// 
			this.cmbFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Fecha"));
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Location = new System.Drawing.Point(104, 8);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 24;
			this.cmbFecha.Size = new System.Drawing.Size(112, 21);
			this.cmbFecha.TabIndex = 0;
			this.cmbFecha.Value = new System.DateTime(2005, 12, 17, 0, 0, 0, 0);
			// 
			// cdsOrdenMov
			// 
			this.cdsOrdenMov.BindingContextCtrl = this;
			this.cdsOrdenMov.DataLibrary = "LibFacturacion";
			this.cdsOrdenMov.DataLibraryUrl = "";
			this.cdsOrdenMov.DataSetDef = "dsOrdenMov";
			this.cdsOrdenMov.FillOnRequest = false;
			this.cdsOrdenMov.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsOrdenMov.Name = "cdsOrdenMov";
			this.cdsOrdenMov.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsOrdenMov.SchemaDef = null;
			this.cdsOrdenMov.BeforeFill += new C1.Data.FillEventHandler(this.cdsOrdenMov_BeforeFill);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(240, 8);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(46, 14);
			this.ultraLabel2.TabIndex = 2;
			this.ultraLabel2.Text = "Cabezal";
			// 
			// txtCabezal
			// 
			this.txtCabezal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Termoking"));
			this.txtCabezal.Location = new System.Drawing.Point(104, 56);
			this.txtCabezal.Name = "txtCabezal";
			this.txtCabezal.Size = new System.Drawing.Size(112, 21);
			this.txtCabezal.TabIndex = 2;
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(24, 80);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(70, 14);
			this.ultraLabel3.TabIndex = 4;
			this.ultraLabel3.Text = "Valor Galon1";
			// 
			// txtValorGalon
			// 
			this.txtValorGalon.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Galones"));
			this.txtValorGalon.FormatString = "#,##0.000";
			this.txtValorGalon.Location = new System.Drawing.Point(104, 32);
			this.txtValorGalon.MaskInput = "{LOC}nnnnnnn.nnn";
			this.txtValorGalon.Name = "txtValorGalon";
			this.txtValorGalon.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorGalon.PromptChar = ' ';
			this.txtValorGalon.Size = new System.Drawing.Size(112, 21);
			this.txtValorGalon.TabIndex = 1;
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.BarraMovimiento = true;
			this.barraDatoSQL1.DataMember = "OrdenMovilizacion";
			this.barraDatoSQL1.DataNombreId = "idMovilizacion";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsOrdenMov;
			this.barraDatoSQL1.DataTabla = "OrdenMovilizacion";
			this.barraDatoSQL1.DataTablaHija = "";
			this.barraDatoSQL1.Location = new System.Drawing.Point(24, 384);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(304, 24);
			this.barraDatoSQL1.TabIndex = 32;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = false;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(24, 32);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(46, 14);
			this.ultraLabel4.TabIndex = 8;
			this.ultraLabel4.Text = "Galones";
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.Location = new System.Drawing.Point(24, 56);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(60, 14);
			this.ultraLabel5.TabIndex = 9;
			this.ultraLabel5.Text = "TermoKing";
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.Location = new System.Drawing.Point(24, 152);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(36, 14);
			this.ultraLabel6.TabIndex = 10;
			this.ultraLabel6.Text = "Varios";
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.AutoSize = true;
			this.ultraLabel7.Location = new System.Drawing.Point(24, 128);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(40, 14);
			this.ultraLabel7.TabIndex = 11;
			this.ultraLabel7.Text = "Sueldo";
			// 
			// ultraCheckEditor1
			// 
			this.ultraCheckEditor1.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsOrdenMov, "OrdenMovilizacion.Descontar"));
			this.ultraCheckEditor1.Location = new System.Drawing.Point(104, 104);
			this.ultraCheckEditor1.Name = "ultraCheckEditor1";
			this.ultraCheckEditor1.Size = new System.Drawing.Size(80, 20);
			this.ultraCheckEditor1.TabIndex = 4;
			this.ultraCheckEditor1.Text = "Descontar";
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.AutoSize = true;
			this.ultraLabel8.Location = new System.Drawing.Point(24, 176);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(38, 14);
			this.ultraLabel8.TabIndex = 14;
			this.ultraLabel8.Text = "Policia";
			// 
			// ultraLabel9
			// 
			this.ultraLabel9.AutoSize = true;
			this.ultraLabel9.Location = new System.Drawing.Point(464, 200);
			this.ultraLabel9.Name = "ultraLabel9";
			this.ultraLabel9.Size = new System.Drawing.Size(58, 14);
			this.ultraLabel9.TabIndex = 18;
			this.ultraLabel9.Text = "Hora Inicio";
			// 
			// ultraLabel10
			// 
			this.ultraLabel10.AutoSize = true;
			this.ultraLabel10.Location = new System.Drawing.Point(240, 200);
			this.ultraLabel10.Name = "ultraLabel10";
			this.ultraLabel10.Size = new System.Drawing.Size(58, 14);
			this.ultraLabel10.TabIndex = 20;
			this.ultraLabel10.Text = "Horometro";
			// 
			// ultraLabel11
			// 
			this.ultraLabel11.AutoSize = true;
			this.ultraLabel11.Location = new System.Drawing.Point(24, 224);
			this.ultraLabel11.Name = "ultraLabel11";
			this.ultraLabel11.Size = new System.Drawing.Size(66, 14);
			this.ultraLabel11.TabIndex = 22;
			this.ultraLabel11.Text = "Kl Transport";
			// 
			// ultraLabel12
			// 
			this.ultraLabel12.AutoSize = true;
			this.ultraLabel12.Location = new System.Drawing.Point(24, 200);
			this.ultraLabel12.Name = "ultraLabel12";
			this.ultraLabel12.Size = new System.Drawing.Size(61, 14);
			this.ultraLabel12.TabIndex = 24;
			this.ultraLabel12.Text = "Kilometraje";
			// 
			// ultraTextEditor9
			// 
			this.ultraTextEditor9.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Observacion"));
			this.ultraTextEditor9.Location = new System.Drawing.Point(104, 248);
			this.ultraTextEditor9.Multiline = true;
			this.ultraTextEditor9.Name = "ultraTextEditor9";
			this.ultraTextEditor9.Size = new System.Drawing.Size(112, 88);
			this.ultraTextEditor9.TabIndex = 10;
			// 
			// ultraLabel13
			// 
			this.ultraLabel13.AutoSize = true;
			this.ultraLabel13.Location = new System.Drawing.Point(24, 248);
			this.ultraLabel13.Name = "ultraLabel13";
			this.ultraLabel13.Size = new System.Drawing.Size(68, 14);
			this.ultraLabel13.TabIndex = 26;
			this.ultraLabel13.Text = "Observacion";
			// 
			// ultraLabel14
			// 
			this.ultraLabel14.AutoSize = true;
			this.ultraLabel14.Location = new System.Drawing.Point(240, 56);
			this.ultraLabel14.Name = "ultraLabel14";
			this.ultraLabel14.Size = new System.Drawing.Size(41, 14);
			this.ultraLabel14.TabIndex = 32;
			this.ultraLabel14.Text = "Galon1";
			// 
			// ultraLabel15
			// 
			this.ultraLabel15.AutoSize = true;
			this.ultraLabel15.Location = new System.Drawing.Point(240, 32);
			this.ultraLabel15.Name = "ultraLabel15";
			this.ultraLabel15.Size = new System.Drawing.Size(64, 14);
			this.ultraLabel15.TabIndex = 31;
			this.ultraLabel15.Text = "Valor Galon";
			// 
			// ultraNumericEditor1
			// 
			this.ultraNumericEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.ValorGalon"));
			this.ultraNumericEditor1.FormatString = "#,##0.000";
			this.ultraNumericEditor1.Location = new System.Drawing.Point(320, 32);
			this.ultraNumericEditor1.MaskInput = "{LOC}nnnnnnn.nnn";
			this.ultraNumericEditor1.Name = "ultraNumericEditor1";
			this.ultraNumericEditor1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor1.PromptChar = ' ';
			this.ultraNumericEditor1.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor1.TabIndex = 12;
			// 
			// ultraLabel16
			// 
			this.ultraLabel16.AutoSize = true;
			this.ultraLabel16.Location = new System.Drawing.Point(240, 80);
			this.ultraLabel16.Name = "ultraLabel16";
			this.ultraLabel16.Size = new System.Drawing.Size(45, 14);
			this.ultraLabel16.TabIndex = 29;
			this.ultraLabel16.Text = "Efectivo";
			// 
			// ultraLabel17
			// 
			this.ultraLabel17.AutoSize = true;
			this.ultraLabel17.Location = new System.Drawing.Point(240, 128);
			this.ultraLabel17.Name = "ultraLabel17";
			this.ultraLabel17.Size = new System.Drawing.Size(43, 14);
			this.ultraLabel17.TabIndex = 38;
			this.ultraLabel17.Text = "Comida";
			// 
			// ultraLabel18
			// 
			this.ultraLabel18.AutoSize = true;
			this.ultraLabel18.Location = new System.Drawing.Point(240, 104);
			this.ultraLabel18.Name = "ultraLabel18";
			this.ultraLabel18.Size = new System.Drawing.Size(58, 14);
			this.ultraLabel18.TabIndex = 37;
			this.ultraLabel18.Text = "Ref. Gasto";
			// 
			// ultraLabel19
			// 
			this.ultraLabel19.AutoSize = true;
			this.ultraLabel19.Location = new System.Drawing.Point(240, 152);
			this.ultraLabel19.Name = "ultraLabel19";
			this.ultraLabel19.Size = new System.Drawing.Size(33, 14);
			this.ultraLabel19.TabIndex = 35;
			this.ultraLabel19.Text = "Peaje";
			// 
			// ultraLabel20
			// 
			this.ultraLabel20.AutoSize = true;
			this.ultraLabel20.Location = new System.Drawing.Point(240, 272);
			this.ultraLabel20.Name = "ultraLabel20";
			this.ultraLabel20.Size = new System.Drawing.Size(34, 14);
			this.ultraLabel20.TabIndex = 50;
			this.ultraLabel20.Text = "Jabas";
			// 
			// ultraLabel21
			// 
			this.ultraLabel21.AutoSize = true;
			this.ultraLabel21.Location = new System.Drawing.Point(240, 248);
			this.ultraLabel21.Name = "ultraLabel21";
			this.ultraLabel21.Size = new System.Drawing.Size(74, 14);
			this.ultraLabel21.TabIndex = 49;
			this.ultraLabel21.Text = "Horometro TL";
			// 
			// ultraNumericEditor3
			// 
			this.ultraNumericEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.KmVacio"));
			this.ultraNumericEditor3.FormatString = "#,##0.000";
			this.ultraNumericEditor3.Location = new System.Drawing.Point(320, 248);
			this.ultraNumericEditor3.MaskInput = "{LOC}nnnnnnn.nnn";
			this.ultraNumericEditor3.Name = "ultraNumericEditor3";
			this.ultraNumericEditor3.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor3.PromptChar = ' ';
			this.ultraNumericEditor3.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor3.TabIndex = 21;
			// 
			// ultraLabel22
			// 
			this.ultraLabel22.AutoSize = true;
			this.ultraLabel22.Location = new System.Drawing.Point(240, 296);
			this.ultraLabel22.Name = "ultraLabel22";
			this.ultraLabel22.Size = new System.Drawing.Size(60, 14);
			this.ultraLabel22.TabIndex = 47;
			this.ultraLabel22.Text = "Reposicion";
			// 
			// ultraLabel23
			// 
			this.ultraLabel23.AutoSize = true;
			this.ultraLabel23.Location = new System.Drawing.Point(464, 224);
			this.ultraLabel23.Name = "ultraLabel23";
			this.ultraLabel23.Size = new System.Drawing.Size(48, 14);
			this.ultraLabel23.TabIndex = 44;
			this.ultraLabel23.Text = "Hora Fin";
			// 
			// ultraLabel24
			// 
			this.ultraLabel24.AutoSize = true;
			this.ultraLabel24.Location = new System.Drawing.Point(240, 176);
			this.ultraLabel24.Name = "ultraLabel24";
			this.ultraLabel24.Size = new System.Drawing.Size(52, 14);
			this.ultraLabel24.TabIndex = 43;
			this.ultraLabel24.Text = "Comision";
			// 
			// ultraNumericEditor4
			// 
			this.ultraNumericEditor4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Comision"));
			this.ultraNumericEditor4.FormatString = "#,##0.00";
			this.ultraNumericEditor4.Location = new System.Drawing.Point(320, 176);
			this.ultraNumericEditor4.Name = "ultraNumericEditor4";
			this.ultraNumericEditor4.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor4.PromptChar = ' ';
			this.ultraNumericEditor4.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor4.TabIndex = 18;
			// 
			// ultraLabel25
			// 
			this.ultraLabel25.AutoSize = true;
			this.ultraLabel25.Location = new System.Drawing.Point(240, 224);
			this.ultraLabel25.Name = "ultraLabel25";
			this.ultraLabel25.Size = new System.Drawing.Size(64, 14);
			this.ultraLabel25.TabIndex = 41;
			this.ultraLabel25.Text = "HorometroL";
			// 
			// ultraTextEditor18
			// 
			this.ultraTextEditor18.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Referencia"));
			this.ultraTextEditor18.Location = new System.Drawing.Point(320, 320);
			this.ultraTextEditor18.Name = "ultraTextEditor18";
			this.ultraTextEditor18.Size = new System.Drawing.Size(112, 21);
			this.ultraTextEditor18.TabIndex = 24;
			// 
			// ultraLabel26
			// 
			this.ultraLabel26.AutoSize = true;
			this.ultraLabel26.Location = new System.Drawing.Point(240, 320);
			this.ultraLabel26.Name = "ultraLabel26";
			this.ultraLabel26.Size = new System.Drawing.Size(59, 14);
			this.ultraLabel26.TabIndex = 52;
			this.ultraLabel26.Text = "Referencia";
			// 
			// ultraCheckEditor2
			// 
			this.ultraCheckEditor2.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsOrdenMov, "OrdenMovilizacion.DescCabezal"));
			this.ultraCheckEditor2.Location = new System.Drawing.Point(464, 8);
			this.ultraCheckEditor2.Name = "ultraCheckEditor2";
			this.ultraCheckEditor2.Size = new System.Drawing.Size(80, 20);
			this.ultraCheckEditor2.TabIndex = 25;
			this.ultraCheckEditor2.Text = "Cabezal";
			// 
			// ultraCheckEditor3
			// 
			this.ultraCheckEditor3.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsOrdenMov, "OrdenMovilizacion.DescEfectivo"));
			this.ultraCheckEditor3.Location = new System.Drawing.Point(464, 80);
			this.ultraCheckEditor3.Name = "ultraCheckEditor3";
			this.ultraCheckEditor3.Size = new System.Drawing.Size(80, 20);
			this.ultraCheckEditor3.TabIndex = 27;
			this.ultraCheckEditor3.Text = "Efectivo";
			// 
			// ultraLabel27
			// 
			this.ultraLabel27.AutoSize = true;
			this.ultraLabel27.Location = new System.Drawing.Point(464, 32);
			this.ultraLabel27.Name = "ultraLabel27";
			this.ultraLabel27.Size = new System.Drawing.Size(59, 14);
			this.ultraLabel27.TabIndex = 58;
			this.ultraLabel27.Text = "Gasolinera";
			// 
			// ultraLabel32
			// 
			this.ultraLabel32.AutoSize = true;
			this.ultraLabel32.Location = new System.Drawing.Point(464, 176);
			this.ultraLabel32.Name = "ultraLabel32";
			this.ultraLabel32.Size = new System.Drawing.Size(40, 14);
			this.ultraLabel32.TabIndex = 68;
			this.ultraLabel32.Text = "Furgon";
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.idArticulo"));
			this.cmbArticulo.DataMember = "Articulo";
			this.cmbArticulo.DataSource = this.cdsArticulo;
			this.cmbArticulo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 38;
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 34;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Width = 309;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbArticulo.DisplayMember = "Articulo";
			this.cmbArticulo.Enabled = false;
			this.cmbArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo.Location = new System.Drawing.Point(336, 384);
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.ReadOnly = true;
			this.cmbArticulo.Size = new System.Drawing.Size(328, 24);
			this.cmbArticulo.TabIndex = 33;
			this.cmbArticulo.ValueMember = "idArticulo";
			// 
			// ultraCalendarCombo1
			// 
			this.ultraCalendarCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.HoraInicio"));
			dateButton2.Caption = "Today";
			this.ultraCalendarCombo1.DateButtons.Add(dateButton2);
			this.ultraCalendarCombo1.Format = "dd/MMM/yyyy HH:mm";
			this.ultraCalendarCombo1.Location = new System.Drawing.Point(536, 200);
			this.ultraCalendarCombo1.Name = "ultraCalendarCombo1";
			this.ultraCalendarCombo1.NonAutoSizeHeight = 24;
			this.ultraCalendarCombo1.Size = new System.Drawing.Size(128, 21);
			this.ultraCalendarCombo1.TabIndex = 30;
			this.ultraCalendarCombo1.Value = new System.DateTime(2008, 3, 18, 0, 0, 0, 0);
			// 
			// ultraCalendarCombo2
			// 
			this.ultraCalendarCombo2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.HoraFinal"));
			dateButton3.Caption = "Today";
			this.ultraCalendarCombo2.DateButtons.Add(dateButton3);
			this.ultraCalendarCombo2.Format = "dd/MMM/yyyy HH:mm";
			this.ultraCalendarCombo2.Location = new System.Drawing.Point(536, 224);
			this.ultraCalendarCombo2.Name = "ultraCalendarCombo2";
			this.ultraCalendarCombo2.NonAutoSizeHeight = 24;
			this.ultraCalendarCombo2.Size = new System.Drawing.Size(128, 21);
			this.ultraCalendarCombo2.TabIndex = 31;
			this.ultraCalendarCombo2.Value = new System.DateTime(2008, 3, 18, 0, 0, 0, 0);
			// 
			// cmbCarro
			// 
			this.cmbCarro.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCarro.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCarro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.idCarro"));
			this.cmbCarro.DataMember = "Carros";
			this.cmbCarro.DataSource = this.cdsOrdenMov;
			ultraGridColumn8.Header.VisiblePosition = 11;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Width = 123;
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 4;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 5;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 7;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 8;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 12;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 9;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 10;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 149;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn20});
			this.cmbCarro.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCarro.DisplayMember = "Placa";
			this.cmbCarro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCarro.Location = new System.Drawing.Point(536, 112);
			this.cmbCarro.Name = "cmbCarro";
			this.cmbCarro.Size = new System.Drawing.Size(128, 21);
			this.cmbCarro.TabIndex = 28;
			this.cmbCarro.ValueMember = "idCarro";
			this.cmbCarro.Validated += new System.EventHandler(this.cmbCarro_Validated);
			// 
			// txtFurgon
			// 
			this.txtFurgon.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Clase"));
			this.txtFurgon.Location = new System.Drawing.Point(536, 168);
			this.txtFurgon.Name = "txtFurgon";
			this.txtFurgon.Size = new System.Drawing.Size(128, 21);
			this.txtFurgon.TabIndex = 29;
			// 
			// ultraTextEditor17
			// 
			this.ultraTextEditor17.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Gasolinera"));
			this.ultraTextEditor17.Location = new System.Drawing.Point(536, 32);
			this.ultraTextEditor17.Name = "ultraTextEditor17";
			this.ultraTextEditor17.Size = new System.Drawing.Size(128, 21);
			this.ultraTextEditor17.TabIndex = 26;
			// 
			// ultraTextEditor19
			// 
			this.ultraTextEditor19.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Cabezal"));
			this.ultraTextEditor19.Location = new System.Drawing.Point(320, 8);
			this.ultraTextEditor19.Name = "ultraTextEditor19";
			this.ultraTextEditor19.Size = new System.Drawing.Size(112, 21);
			this.ultraTextEditor19.TabIndex = 11;
			// 
			// ultraNumericEditor5
			// 
			this.ultraNumericEditor5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Galon1"));
			this.ultraNumericEditor5.FormatString = "#,##0.000";
			this.ultraNumericEditor5.Location = new System.Drawing.Point(320, 56);
			this.ultraNumericEditor5.MaskInput = "{LOC}nnnnnnn.nnn";
			this.ultraNumericEditor5.Name = "ultraNumericEditor5";
			this.ultraNumericEditor5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor5.PromptChar = ' ';
			this.ultraNumericEditor5.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor5.TabIndex = 13;
			// 
			// ultraNumericEditor6
			// 
			this.ultraNumericEditor6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.ValorGalon1"));
			this.ultraNumericEditor6.FormatString = "#,##0.000";
			this.ultraNumericEditor6.Location = new System.Drawing.Point(104, 80);
			this.ultraNumericEditor6.MaskInput = "{LOC}nnnnnnn.nnn";
			this.ultraNumericEditor6.Name = "ultraNumericEditor6";
			this.ultraNumericEditor6.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor6.PromptChar = ' ';
			this.ultraNumericEditor6.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor6.TabIndex = 3;
			// 
			// ultraNumericEditor7
			// 
			this.ultraNumericEditor7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Efectivo"));
			this.ultraNumericEditor7.FormatString = "#,##0.00";
			this.ultraNumericEditor7.Location = new System.Drawing.Point(320, 80);
			this.ultraNumericEditor7.Name = "ultraNumericEditor7";
			this.ultraNumericEditor7.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor7.PromptChar = ' ';
			this.ultraNumericEditor7.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor7.TabIndex = 14;
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.RefGasto"));
			this.ultraTextEditor1.Location = new System.Drawing.Point(320, 104);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(112, 21);
			this.ultraTextEditor1.TabIndex = 15;
			// 
			// ultraNumericEditor2
			// 
			this.ultraNumericEditor2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Comida"));
			this.ultraNumericEditor2.FormatString = "#,##0.00";
			this.ultraNumericEditor2.Location = new System.Drawing.Point(320, 128);
			this.ultraNumericEditor2.Name = "ultraNumericEditor2";
			this.ultraNumericEditor2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor2.PromptChar = ' ';
			this.ultraNumericEditor2.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor2.TabIndex = 16;
			// 
			// ultraNumericEditor8
			// 
			this.ultraNumericEditor8.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Peaje"));
			this.ultraNumericEditor8.FormatString = "#,##0.00";
			this.ultraNumericEditor8.Location = new System.Drawing.Point(320, 152);
			this.ultraNumericEditor8.Name = "ultraNumericEditor8";
			this.ultraNumericEditor8.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor8.PromptChar = ' ';
			this.ultraNumericEditor8.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor8.TabIndex = 17;
			// 
			// ultraNumericEditor9
			// 
			this.ultraNumericEditor9.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Sueldo"));
			this.ultraNumericEditor9.FormatString = "#,##0.00";
			this.ultraNumericEditor9.Location = new System.Drawing.Point(104, 128);
			this.ultraNumericEditor9.Name = "ultraNumericEditor9";
			this.ultraNumericEditor9.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor9.PromptChar = ' ';
			this.ultraNumericEditor9.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor9.TabIndex = 5;
			// 
			// ultraNumericEditor10
			// 
			this.ultraNumericEditor10.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Varios"));
			this.ultraNumericEditor10.FormatString = "#,##0.00";
			this.ultraNumericEditor10.Location = new System.Drawing.Point(104, 152);
			this.ultraNumericEditor10.Name = "ultraNumericEditor10";
			this.ultraNumericEditor10.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor10.PromptChar = ' ';
			this.ultraNumericEditor10.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor10.TabIndex = 6;
			// 
			// ultraNumericEditor11
			// 
			this.ultraNumericEditor11.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Policia"));
			this.ultraNumericEditor11.FormatString = "#,##0.00";
			this.ultraNumericEditor11.Location = new System.Drawing.Point(104, 176);
			this.ultraNumericEditor11.Name = "ultraNumericEditor11";
			this.ultraNumericEditor11.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor11.PromptChar = ' ';
			this.ultraNumericEditor11.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor11.TabIndex = 7;
			// 
			// ultraNumericEditor12
			// 
			this.ultraNumericEditor12.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Horometros"));
			this.ultraNumericEditor12.FormatString = "#,##0.000";
			this.ultraNumericEditor12.Location = new System.Drawing.Point(320, 200);
			this.ultraNumericEditor12.MaskInput = "{LOC}nnnnnnn.nnn";
			this.ultraNumericEditor12.Name = "ultraNumericEditor12";
			this.ultraNumericEditor12.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor12.PromptChar = ' ';
			this.ultraNumericEditor12.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor12.TabIndex = 19;
			// 
			// ultraNumericEditor13
			// 
			this.ultraNumericEditor13.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.HorometroL"));
			this.ultraNumericEditor13.FormatString = "#,##0.000";
			this.ultraNumericEditor13.Location = new System.Drawing.Point(320, 224);
			this.ultraNumericEditor13.MaskInput = "{LOC}nnnnnnn.nnn";
			this.ultraNumericEditor13.Name = "ultraNumericEditor13";
			this.ultraNumericEditor13.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor13.PromptChar = ' ';
			this.ultraNumericEditor13.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor13.TabIndex = 20;
			// 
			// ultraNumericEditor14
			// 
			this.ultraNumericEditor14.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.KmCargado"));
			this.ultraNumericEditor14.FormatString = "#,##0.000";
			this.ultraNumericEditor14.Location = new System.Drawing.Point(104, 200);
			this.ultraNumericEditor14.MaskInput = "{LOC}nnnnnnn.nnn";
			this.ultraNumericEditor14.Name = "ultraNumericEditor14";
			this.ultraNumericEditor14.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor14.PromptChar = ' ';
			this.ultraNumericEditor14.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor14.TabIndex = 8;
			// 
			// ultraNumericEditor15
			// 
			this.ultraNumericEditor15.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.KmTransporte"));
			this.ultraNumericEditor15.FormatString = "#,##0.000";
			this.ultraNumericEditor15.Location = new System.Drawing.Point(104, 224);
			this.ultraNumericEditor15.MaskInput = "{LOC}nnnnnnn.nnn";
			this.ultraNumericEditor15.Name = "ultraNumericEditor15";
			this.ultraNumericEditor15.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor15.PromptChar = ' ';
			this.ultraNumericEditor15.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor15.TabIndex = 9;
			this.ultraNumericEditor15.ValueChanged += new System.EventHandler(this.ultraNumericEditor15_ValueChanged);
			// 
			// ultraNumericEditor16
			// 
			this.ultraNumericEditor16.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Jabas"));
			this.ultraNumericEditor16.FormatString = "#,##0.00";
			this.ultraNumericEditor16.Location = new System.Drawing.Point(320, 272);
			this.ultraNumericEditor16.Name = "ultraNumericEditor16";
			this.ultraNumericEditor16.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor16.PromptChar = ' ';
			this.ultraNumericEditor16.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor16.TabIndex = 22;
			// 
			// ultraNumericEditor17
			// 
			this.ultraNumericEditor17.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Reposicion"));
			this.ultraNumericEditor17.FormatString = "#,##0.00";
			this.ultraNumericEditor17.Location = new System.Drawing.Point(320, 296);
			this.ultraNumericEditor17.Name = "ultraNumericEditor17";
			this.ultraNumericEditor17.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor17.PromptChar = ' ';
			this.ultraNumericEditor17.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor17.TabIndex = 23;
			// 
			// btPlaca
			// 
			this.btPlaca.Location = new System.Drawing.Point(464, 112);
			this.btPlaca.Name = "btPlaca";
			this.btPlaca.Size = new System.Drawing.Size(56, 23);
			this.btPlaca.TabIndex = 69;
			this.btPlaca.Text = "Carro";
			this.btPlaca.Click += new System.EventHandler(this.btPlaca_Click);
			// 
			// ultraLabel28
			// 
			this.ultraLabel28.AutoSize = true;
			this.ultraLabel28.Location = new System.Drawing.Point(464, 248);
			this.ultraLabel28.Name = "ultraLabel28";
			this.ultraLabel28.Size = new System.Drawing.Size(70, 14);
			this.ultraLabel28.TabIndex = 71;
			this.ultraLabel28.Text = "Horomtre EN";
			// 
			// txtIdMovilizacion
			// 
			this.txtIdMovilizacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.idMovilizacion"));
			this.txtIdMovilizacion.Location = new System.Drawing.Point(600, 72);
			this.txtIdMovilizacion.Name = "txtIdMovilizacion";
			this.txtIdMovilizacion.Size = new System.Drawing.Size(40, 21);
			this.txtIdMovilizacion.TabIndex = 72;
			// 
			// ultraLabel29
			// 
			this.ultraLabel29.AutoSize = true;
			this.ultraLabel29.Location = new System.Drawing.Point(464, 272);
			this.ultraLabel29.Name = "ultraLabel29";
			this.ultraLabel29.Size = new System.Drawing.Size(60, 14);
			this.ultraLabel29.TabIndex = 74;
			this.ultraLabel29.Text = "Km Furgon";
			// 
			// ultraLabel30
			// 
			this.ultraLabel30.AutoSize = true;
			this.ultraLabel30.Location = new System.Drawing.Point(464, 296);
			this.ultraLabel30.Name = "ultraLabel30";
			this.ultraLabel30.Size = new System.Drawing.Size(26, 14);
			this.ultraLabel30.TabIndex = 76;
			this.ultraLabel30.Text = "Otro";
			// 
			// ultraNumericEditor18
			// 
			this.ultraNumericEditor18.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.HorometroEN"));
			this.ultraNumericEditor18.FormatString = "#,##0.000";
			this.ultraNumericEditor18.Location = new System.Drawing.Point(552, 248);
			this.ultraNumericEditor18.MaskInput = "{LOC}nnnnnnn.nnn";
			this.ultraNumericEditor18.Name = "ultraNumericEditor18";
			this.ultraNumericEditor18.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor18.PromptChar = ' ';
			this.ultraNumericEditor18.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor18.TabIndex = 77;
			// 
			// ultraNumericEditor19
			// 
			this.ultraNumericEditor19.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.KmFurgon"));
			this.ultraNumericEditor19.FormatString = "#,##0.000";
			this.ultraNumericEditor19.Location = new System.Drawing.Point(552, 272);
			this.ultraNumericEditor19.MaskInput = "{LOC}nnnnnnn.nnn";
			this.ultraNumericEditor19.Name = "ultraNumericEditor19";
			this.ultraNumericEditor19.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor19.PromptChar = ' ';
			this.ultraNumericEditor19.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor19.TabIndex = 78;
			// 
			// ultraNumericEditor20
			// 
			this.ultraNumericEditor20.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.Otro"));
			this.ultraNumericEditor20.FormatString = "#,##0.000";
			this.ultraNumericEditor20.Location = new System.Drawing.Point(552, 296);
			this.ultraNumericEditor20.MaskInput = "{LOC}nnnnnnn.nnn";
			this.ultraNumericEditor20.Name = "ultraNumericEditor20";
			this.ultraNumericEditor20.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.ultraNumericEditor20.PromptChar = ' ';
			this.ultraNumericEditor20.Size = new System.Drawing.Size(112, 21);
			this.ultraNumericEditor20.TabIndex = 79;
			// 
			// cmbCodigo
			// 
			this.cmbCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCodigo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.idArticulo"));
			this.cmbCodigo.DataMember = "Articulo";
			this.cmbCodigo.DataSource = this.cdsArticulo;
			this.cmbCodigo.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 5;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 17;
			ultraGridColumn24.Header.VisiblePosition = 6;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 17;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 93;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Width = 93;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbCodigo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCodigo.DisplayMember = "Codigo";
			this.cmbCodigo.Enabled = false;
			this.cmbCodigo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodigo.Location = new System.Drawing.Point(552, 320);
			this.cmbCodigo.Name = "cmbCodigo";
			this.cmbCodigo.ReadOnly = true;
			this.cmbCodigo.Size = new System.Drawing.Size(112, 24);
			this.cmbCodigo.TabIndex = 80;
			this.cmbCodigo.ValueMember = "idArticulo";
			// 
			// ultraLabel31
			// 
			this.ultraLabel31.AutoSize = true;
			this.ultraLabel31.Location = new System.Drawing.Point(464, 320);
			this.ultraLabel31.Name = "ultraLabel31";
			this.ultraLabel31.Size = new System.Drawing.Size(40, 14);
			this.ultraLabel31.TabIndex = 81;
			this.ultraLabel31.Text = "Codigo";
			// 
			// cmbEquipo
			// 
			this.cmbEquipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEquipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEquipo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.idCarro"));
			this.cmbEquipo.DataMember = "Carros";
			this.cmbEquipo.DataSource = this.cdsOrdenMov;
			ultraGridColumn28.Header.VisiblePosition = 11;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn31.Header.VisiblePosition = 3;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 4;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 6;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 7;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 8;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 12;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 9;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 10;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Width = 131;
			ultraGridBand4.Columns.AddRange(new object[] {
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
			this.cmbEquipo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbEquipo.DisplayMember = "Equipo";
			this.cmbEquipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEquipo.Location = new System.Drawing.Point(536, 136);
			this.cmbEquipo.Name = "cmbEquipo";
			this.cmbEquipo.Size = new System.Drawing.Size(128, 21);
			this.cmbEquipo.TabIndex = 82;
			this.cmbEquipo.ValueMember = "idCarro";
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// ultraCalendarCombo3
			// 
			this.ultraCalendarCombo3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.FechaVenta"));
			dateButton4.Caption = "Today";
			this.ultraCalendarCombo3.DateButtons.Add(dateButton4);
			this.ultraCalendarCombo3.Format = "dd/MMM/yyyy";
			this.ultraCalendarCombo3.Location = new System.Drawing.Point(104, 344);
			this.ultraCalendarCombo3.Name = "ultraCalendarCombo3";
			this.ultraCalendarCombo3.NonAutoSizeHeight = 24;
			this.ultraCalendarCombo3.Size = new System.Drawing.Size(112, 21);
			this.ultraCalendarCombo3.TabIndex = 84;
			this.ultraCalendarCombo3.Value = new System.DateTime(2008, 3, 18, 0, 0, 0, 0);
			// 
			// ultraLabel33
			// 
			this.ultraLabel33.AutoSize = true;
			this.ultraLabel33.Location = new System.Drawing.Point(24, 344);
			this.ultraLabel33.Name = "ultraLabel33";
			this.ultraLabel33.Size = new System.Drawing.Size(68, 14);
			this.ultraLabel33.TabIndex = 83;
			this.ultraLabel33.Text = "Fecha Venta";
			// 
			// ultraCalendarCombo4
			// 
			this.ultraCalendarCombo4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.FinDescargue"));
			dateButton5.Caption = "Today";
			this.ultraCalendarCombo4.DateButtons.Add(dateButton5);
			this.ultraCalendarCombo4.Format = "dd/MMM/yyyy HH:mm";
			this.ultraCalendarCombo4.Location = new System.Drawing.Point(536, 352);
			this.ultraCalendarCombo4.Name = "ultraCalendarCombo4";
			this.ultraCalendarCombo4.NonAutoSizeHeight = 24;
			this.ultraCalendarCombo4.Size = new System.Drawing.Size(128, 21);
			this.ultraCalendarCombo4.TabIndex = 87;
			this.ultraCalendarCombo4.Value = new System.DateTime(2008, 3, 18, 0, 0, 0, 0);
			// 
			// ultraCalendarCombo5
			// 
			this.ultraCalendarCombo5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenMov, "OrdenMovilizacion.InicioViaje"));
			dateButton6.Caption = "Today";
			this.ultraCalendarCombo5.DateButtons.Add(dateButton6);
			this.ultraCalendarCombo5.Format = "dd/MMM/yyyy HH:mm";
			this.ultraCalendarCombo5.Location = new System.Drawing.Point(312, 352);
			this.ultraCalendarCombo5.Name = "ultraCalendarCombo5";
			this.ultraCalendarCombo5.NonAutoSizeHeight = 24;
			this.ultraCalendarCombo5.Size = new System.Drawing.Size(128, 21);
			this.ultraCalendarCombo5.TabIndex = 86;
			this.ultraCalendarCombo5.Value = new System.DateTime(2008, 3, 18, 0, 0, 0, 0);
			// 
			// ultraLabel34
			// 
			this.ultraLabel34.AutoSize = true;
			this.ultraLabel34.Location = new System.Drawing.Point(464, 352);
			this.ultraLabel34.Name = "ultraLabel34";
			this.ultraLabel34.Size = new System.Drawing.Size(68, 14);
			this.ultraLabel34.TabIndex = 88;
			this.ultraLabel34.Text = "Fin Descarg.";
			// 
			// ultraLabel35
			// 
			this.ultraLabel35.AutoSize = true;
			this.ultraLabel35.Location = new System.Drawing.Point(240, 352);
			this.ultraLabel35.Name = "ultraLabel35";
			this.ultraLabel35.Size = new System.Drawing.Size(59, 14);
			this.ultraLabel35.TabIndex = 85;
			this.ultraLabel35.Text = "Inicio Viaje";
			// 
			// OrdenMovilizacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(680, 430);
			this.Controls.Add(this.ultraCalendarCombo4);
			this.Controls.Add(this.ultraCalendarCombo5);
			this.Controls.Add(this.ultraLabel34);
			this.Controls.Add(this.ultraLabel35);
			this.Controls.Add(this.ultraCalendarCombo3);
			this.Controls.Add(this.ultraLabel33);
			this.Controls.Add(this.cmbEquipo);
			this.Controls.Add(this.ultraLabel31);
			this.Controls.Add(this.cmbCodigo);
			this.Controls.Add(this.ultraNumericEditor20);
			this.Controls.Add(this.ultraNumericEditor19);
			this.Controls.Add(this.ultraNumericEditor18);
			this.Controls.Add(this.ultraLabel30);
			this.Controls.Add(this.ultraLabel29);
			this.Controls.Add(this.txtIdMovilizacion);
			this.Controls.Add(this.ultraLabel28);
			this.Controls.Add(this.btPlaca);
			this.Controls.Add(this.ultraNumericEditor17);
			this.Controls.Add(this.ultraNumericEditor16);
			this.Controls.Add(this.ultraNumericEditor15);
			this.Controls.Add(this.ultraNumericEditor14);
			this.Controls.Add(this.ultraNumericEditor13);
			this.Controls.Add(this.ultraNumericEditor12);
			this.Controls.Add(this.ultraNumericEditor11);
			this.Controls.Add(this.ultraNumericEditor10);
			this.Controls.Add(this.ultraNumericEditor9);
			this.Controls.Add(this.ultraNumericEditor8);
			this.Controls.Add(this.ultraNumericEditor2);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.ultraNumericEditor7);
			this.Controls.Add(this.ultraNumericEditor6);
			this.Controls.Add(this.ultraNumericEditor5);
			this.Controls.Add(this.ultraTextEditor19);
			this.Controls.Add(this.ultraTextEditor17);
			this.Controls.Add(this.txtFurgon);
			this.Controls.Add(this.cmbCarro);
			this.Controls.Add(this.ultraCalendarCombo2);
			this.Controls.Add(this.ultraCalendarCombo1);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.ultraLabel32);
			this.Controls.Add(this.ultraLabel27);
			this.Controls.Add(this.ultraCheckEditor3);
			this.Controls.Add(this.ultraCheckEditor2);
			this.Controls.Add(this.ultraTextEditor18);
			this.Controls.Add(this.ultraLabel26);
			this.Controls.Add(this.ultraLabel20);
			this.Controls.Add(this.ultraLabel21);
			this.Controls.Add(this.ultraNumericEditor3);
			this.Controls.Add(this.ultraLabel22);
			this.Controls.Add(this.ultraLabel23);
			this.Controls.Add(this.ultraLabel24);
			this.Controls.Add(this.ultraNumericEditor4);
			this.Controls.Add(this.ultraLabel25);
			this.Controls.Add(this.ultraLabel17);
			this.Controls.Add(this.ultraLabel18);
			this.Controls.Add(this.ultraLabel19);
			this.Controls.Add(this.ultraLabel14);
			this.Controls.Add(this.ultraLabel15);
			this.Controls.Add(this.ultraNumericEditor1);
			this.Controls.Add(this.ultraLabel16);
			this.Controls.Add(this.ultraTextEditor9);
			this.Controls.Add(this.ultraLabel13);
			this.Controls.Add(this.ultraLabel12);
			this.Controls.Add(this.ultraLabel11);
			this.Controls.Add(this.ultraLabel10);
			this.Controls.Add(this.ultraLabel9);
			this.Controls.Add(this.ultraLabel8);
			this.Controls.Add(this.ultraCheckEditor1);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.txtValorGalon);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.txtCabezal);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.ultraLabel1);
			this.Name = "OrdenMovilizacion";
			this.Text = "Orden de Movilización";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.OrdenMovilizacion_Closing);
			this.Load += new System.EventHandler(this.OrdenMovilizacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsOrdenMov)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCabezal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorGalon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor18)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCarro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFurgon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor17)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor19)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor15)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor16)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor17)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdMovilizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor18)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor19)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraNumericEditor20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEquipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo5)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar Ordenes de Movilizacion",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show("Desea Borrar esta orden de Movilizacion",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
				MessageBoxDefaultButton.Button2) ) return;
			barraDatoSQL1.BorraRegistro();
		}

		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void ValoresIniciales()
		{
			drOrdenM["Fecha"] = DateTime.Today;
			drOrdenM["Cabezal"] = "";
			txtCabezal.Select();
		}

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a Crear Nuevas Ordenes de Movilizacion",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}			
			barraDatoSQL1.CrearRegistro();
			ValoresIniciales();
		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			barraDatoSQL1.DeshacerRegistro();
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar Ordenes de Movilzacion",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDatoSQL1.EditarRegistro();
		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir Ordenes de Movilizacion",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			OrdenMovRep miRep = new OrdenMovRep();
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
		}

		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			barraDatoSQL1.GrabaRegistro();
		}

		private C1.Data.C1DataRow drOrdenM;
		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drOrdenM = cdsOrdenMov.TableViews["OrdenMovilizacion"].Rows[0];
			}
			catch{};
			
		}

		private void cdsOrdenMov_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsOrdenMov.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		CultureInfo us = new CultureInfo("en-US");
		private Acceso miAcceso;
		private void OrdenMovilizacion_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Orden Movilizacion'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			miAcceso =new Acceso("0410M", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			txtIdMovilizacion.Width = 0;
			if (IdArticulo > 0)
			{
				string stSelect = "Select idMovilizacion From OrdenMovilizacion "
					+ "Where idArticulo = " + IdArticulo.ToString();
				int IdOrdenMov = Funcion.iEscalarSQL(cdsOrdenMov, stSelect, false);
				if (IdOrdenMov == 0)
				{
					stSelect = "Select IsNull(Caracteristica, '') From Articulo "
            + "Where idArticulo = " + IdArticulo.ToString();
					stSelect = Funcion.sEscalarSQL(cdsArticulo, stSelect).Trim();
					DateTime dtFecha;
					try
					{
						stSelect +=  " 01:01:00";
						dtFecha = DateTime.Parse(stSelect);
					}
					catch (System.Exception ex)
					{
						MessageBox.Show(ex.Message + "Fecha Erronea en caracteristica", "Información");
						dtFecha = DateTime.Today;
					}
					barraDatoSQL1.CrearRegistro();
					drOrdenM["Fecha"] = dtFecha;
					drOrdenM["HoraInicio"] = DateTime.Now;
					drOrdenM["HoraFinal"] = DateTime.Now;
					drOrdenM["idArticulo"] = IdArticulo;
				}
				else
				{
					barraDatoSQL1.IdRegistro = IdOrdenMov;
					barraDatoSQL1.ProximoId(5);
					barraDatoSQL1.HabilitaControles(false);
				}
			}
		}

		private void btPlaca_Click(object sender, System.EventArgs e)
		{
			using(Carro miCarro = new Carro())
			{
				miCarro.ShowDialog();
			}
			cdsOrdenMov.Clear();
			cdsOrdenMov.Fill();
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			string stFiltro = "{OrdenMovilizacion.idMovilizacion} = " 
				+ txtIdMovilizacion.Value.ToString();
			Reporte miRep = new Reporte("OrdenMov1.rpt", stFiltro);
			miRep.MdiParent = MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void cmbCarro_Validated(object sender, System.EventArgs e)
		{
			string stFurgon = cmbCarro.ActiveRow.Cells["Numero"].Value.ToString();
			txtFurgon.Text = stFurgon;
			txtFurgon.Refresh();
			txtFurgon.Select();
		}

		private void ultraNumericEditor15_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void OrdenMovilizacion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Orden Movilizacion'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}
	}

}
