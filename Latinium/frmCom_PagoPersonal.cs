using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_PagoPersonal.
	/// </summary>
	public class frmCom_PagoPersonal : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnExportar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpPeriodo;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECentroCosto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEDepartamento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridComisiones;
		private System.Windows.Forms.Button btnAcreditacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdComision;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.ContextMenu mnuAcreditacion;
		private System.Windows.Forms.MenuItem mnuAcreditacionTransferencias;
		private System.Windows.Forms.MenuItem mnuAcreditacionCheques;
		private System.Windows.Forms.Button btnAgregar;
		private System.Windows.Forms.Button btnSeleccionar;
		private System.Windows.Forms.Button btnCerrar;
		private System.Windows.Forms.ContextMenu mnuAgregar;
		private System.Windows.Forms.MenuItem mnuAgregarRegistro;
		private System.Windows.Forms.MenuItem mnuActualizar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAntipoUtilidades;
		private System.Windows.Forms.Button btnUtilidades;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		
		public frmCom_PagoPersonal()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int Estado = 0;
		bool bComisionQuincena = false;

		public frmCom_PagoPersonal(bool BComisionQuincena)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			
			bComisionQuincena = BComisionQuincena;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_PagoPersonal));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Revisado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Acreditado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Apellidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Salida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDePago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDePagoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Centro_De_Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Departamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Infraestructura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobranzas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RecursosHumanos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Administrativo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ventas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Financiero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Compras");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contabilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sistemas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marketing");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Auditoria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DepartamentoTecnico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comisiones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AnticipoUtilidades");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComprasCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComprasCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FaltantesCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Multas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AnticiposSueldos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Otros");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SeguroMedico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Farmacia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalEgresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesFaltCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesMultas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesAnticipos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesOtros");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraC");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Revisado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acreditado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellidos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Salida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDePago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDePagoM");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Centro_De_Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Infraestructura");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobranzas");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operaciones");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RecursosHumanos");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Administrativo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ventas");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Financiero");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Compras");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contabilidad");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sistemas");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marketing");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Auditoria");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DepartamentoTecnico");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comisiones");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AnticipoUtilidades");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComprasCredito");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComprasCrediPoint");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FaltantesCaja");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Multas");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AnticiposSueldos");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Otros");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SeguroMedico");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Farmacia");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalEgresos");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesFaltCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesMultas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesAnticipos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesOtros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnFC", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnOtros", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnticipos", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMultas", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCredito", 4);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comisiones", 34, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comisiones", 34, true);
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ComprasCredito", 36, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ComprasCredito", 36, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ComprasCrediPoint", 37, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ComprasCrediPoint", 37, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FaltantesCaja", 38, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FaltantesCaja", 38, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Multas", 39, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Multas", 39, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AnticiposSueldos", 40, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AnticiposSueldos", 40, true);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SeguroMedico", 42, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SeguroMedico", 42, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Farmacia", 43, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Farmacia", 43, true);
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Otros", 41, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Otros", 41, true);
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalEgresos", 44, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalEgresos", 44, true);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 45, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 45, true);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Infraestructura", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Infraestructura", 19, true);
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cobranzas", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cobranzas", 20, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Operaciones", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Operaciones", 21, true);
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RecursosHumanos", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RecursosHumanos", 22, true);
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bodega", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bodega", 23, true);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Administrativo", 24, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Administrativo", 24, true);
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Ventas", 25, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Ventas", 25, true);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Financiero", 26, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Financiero", 26, true);
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 27, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 27, true);
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Compras", 28, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Compras", 28, true);
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Contabilidad", 29, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Contabilidad", 29, true);
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Sistemas", 30, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Sistemas", 30, true);
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Marketing", 31, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Marketing", 31, true);
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Auditoria", 32, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Auditoria", 32, true);
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DepartamentoTecnico", 33, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DepartamentoTecnico", 33, true);
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings27 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AnticipoUtilidades", 35, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AnticipoUtilidades", 35, true);
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.btnExportar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtIdComision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.dtpPeriodo = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridComisiones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label12 = new System.Windows.Forms.Label();
			this.cmbNECentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNEDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label33 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAcreditacion = new System.Windows.Forms.Button();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.mnuAcreditacion = new System.Windows.Forms.ContextMenu();
			this.mnuAcreditacionTransferencias = new System.Windows.Forms.MenuItem();
			this.mnuAcreditacionCheques = new System.Windows.Forms.MenuItem();
			this.btnAgregar = new System.Windows.Forms.Button();
			this.btnSeleccionar = new System.Windows.Forms.Button();
			this.btnCerrar = new System.Windows.Forms.Button();
			this.mnuAgregar = new System.Windows.Forms.ContextMenu();
			this.mnuAgregarRegistro = new System.Windows.Forms.MenuItem();
			this.mnuActualizar = new System.Windows.Forms.MenuItem();
			this.txtIdAntipoUtilidades = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnUtilidades = new System.Windows.Forms.Button();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpPeriodo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridComisiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAntipoUtilidades)).BeginInit();
			this.SuspendLayout();
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(888, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.TabIndex = 733;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
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
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			ultraDataColumn4.DataType = typeof(bool);
			ultraDataColumn4.DefaultValue = false;
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(System.DateTime);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn20.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn21.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn22.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn23.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn25.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn26.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn27.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn28.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn29.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn30.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn31.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn32.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn33.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn34.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn35.DataType = typeof(System.Decimal);
			ultraDataColumn35.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn36.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn37.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn38.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn39.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn40.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn41.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn42.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn43.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn44.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn45.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn46.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn47.DefaultValue = 0;
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn48.DefaultValue = 0;
			ultraDataColumn54.DataType = typeof(int);
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
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54});
			// 
			// txtIdComision
			// 
			this.txtIdComision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdComision.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdComision.Enabled = false;
			this.txtIdComision.Location = new System.Drawing.Point(605, 9);
			this.txtIdComision.Name = "txtIdComision";
			this.txtIdComision.PromptChar = ' ';
			this.txtIdComision.Size = new System.Drawing.Size(8, 22);
			this.txtIdComision.TabIndex = 732;
			this.txtIdComision.Visible = false;
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Enabled = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(616, 8);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.TabIndex = 730;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 73);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(82, 17);
			this.label7.TabIndex = 729;
			this.label7.Text = "Cedula/Nombre";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(42, 17);
			this.label6.TabIndex = 724;
			this.label6.Text = "Periodo";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(104, 72);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(200, 20);
			this.txtBuscar.TabIndex = 728;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			this.txtBuscar.ValueChanged += new System.EventHandler(this.txtBuscar_ValueChanged);
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(328, 70);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(64, 23);
			this.btnVer.TabIndex = 727;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1152, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 726;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// dtpPeriodo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpPeriodo.Appearance = appearance1;
			this.dtpPeriodo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpPeriodo.DateButtons.Add(dateButton1);
			this.dtpPeriodo.Format = "MM/yyyy";
			this.dtpPeriodo.Location = new System.Drawing.Point(104, 9);
			this.dtpPeriodo.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpPeriodo.Name = "dtpPeriodo";
			this.dtpPeriodo.NonAutoSizeHeight = 23;
			this.dtpPeriodo.Size = new System.Drawing.Size(104, 21);
			this.dtpPeriodo.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtpPeriodo.TabIndex = 725;
			this.dtpPeriodo.Value = ((object)(resources.GetObject("dtpPeriodo.Value")));
			this.dtpPeriodo.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtpPeriodo_BeforeDropDown);
			this.dtpPeriodo.ValueChanged += new System.EventHandler(this.dtpPeriodo_ValueChanged);
			// 
			// uGridComisiones
			// 
			this.uGridComisiones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridComisiones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridComisiones.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridComisiones.DisplayLayout.Appearance = appearance2;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.Caption = "R";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Width = 20;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance3;
			ultraGridColumn4.Header.VisiblePosition = 11;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 40;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance4;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 50;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Width = 150;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 150;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Width = 70;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 9;
			ultraGridColumn11.Width = 70;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance5;
			ultraGridColumn13.Header.Caption = "Pago";
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.Width = 80;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 120;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.Caption = "Centro De Costo";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 200;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Format = "";
			ultraGridColumn18.Header.VisiblePosition = 18;
			ultraGridColumn18.Width = 150;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn19.Header.VisiblePosition = 17;
			ultraGridColumn19.Width = 150;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance6;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 120;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance7;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 120;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance8;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 120;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance9;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "RRHH";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 120;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance10;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 120;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance11;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 120;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance12;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 120;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance13;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Width = 120;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance14;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Width = 120;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance15;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Width = 120;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance16;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Width = 120;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance17;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Width = 120;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance18;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Width = 120;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance19;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Width = 120;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance20;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.Caption = "D T";
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.Width = 120;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance21;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 35;
			ultraGridColumn35.PromptChar = ' ';
			ultraGridColumn35.Width = 90;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance22;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.Caption = "A. Utilidades";
			ultraGridColumn36.Header.VisiblePosition = 34;
			ultraGridColumn36.Width = 90;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance23;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.Caption = "Compras Credito";
			ultraGridColumn37.Header.VisiblePosition = 38;
			ultraGridColumn37.PromptChar = ' ';
			ultraGridColumn37.Width = 90;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance24;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "Credi Point";
			ultraGridColumn38.Header.VisiblePosition = 36;
			ultraGridColumn38.PromptChar = ' ';
			ultraGridColumn38.Width = 90;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance25;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.Caption = "Falt. de Caja";
			ultraGridColumn39.Header.VisiblePosition = 39;
			ultraGridColumn39.PromptChar = ' ';
			ultraGridColumn39.Width = 90;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance26;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.VisiblePosition = 41;
			ultraGridColumn40.PromptChar = ' ';
			ultraGridColumn40.Width = 90;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance27;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.Caption = "Anticipos";
			ultraGridColumn41.Header.VisiblePosition = 43;
			ultraGridColumn41.PromptChar = ' ';
			ultraGridColumn41.Width = 90;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance28;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.VisiblePosition = 47;
			ultraGridColumn42.PromptChar = ' ';
			ultraGridColumn42.Width = 90;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance29;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.Caption = "Seguro";
			ultraGridColumn43.Header.VisiblePosition = 45;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.PromptChar = ' ';
			ultraGridColumn43.Width = 90;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance30;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 46;
			ultraGridColumn44.PromptChar = ' ';
			ultraGridColumn44.Width = 90;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance31;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.Caption = "Egresos";
			ultraGridColumn45.Header.VisiblePosition = 49;
			ultraGridColumn45.Width = 90;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance32;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 52;
			ultraGridColumn46.Width = 90;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn47.Header.VisiblePosition = 58;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn48.Header.VisiblePosition = 57;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 50;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 51;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 53;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 56;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 54;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 55;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn55.Header.Caption = "F C";
			ultraGridColumn55.Header.VisiblePosition = 40;
			ultraGridColumn55.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn55.Width = 30;
			ultraGridColumn56.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn56.Header.Caption = "O";
			ultraGridColumn56.Header.VisiblePosition = 48;
			ultraGridColumn56.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn56.Width = 30;
			ultraGridColumn57.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn57.Header.Caption = "A";
			ultraGridColumn57.Header.VisiblePosition = 44;
			ultraGridColumn57.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn57.Width = 30;
			ultraGridColumn58.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn58.Header.Caption = "M";
			ultraGridColumn58.Header.VisiblePosition = 42;
			ultraGridColumn58.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn58.Width = 30;
			ultraGridColumn59.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn59.Header.Caption = "C";
			ultraGridColumn59.Header.VisiblePosition = 37;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn59.Width = 30;
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
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59});
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance33;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance34;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance35;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance36;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance37;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance38;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance39;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.Hidden = true;
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance40;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance41;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance42;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance43;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance44;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance45;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance46;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance47;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance48;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance49;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance50;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance51;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance52;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance53;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance54;
			summarySettings22.DisplayFormat = "{0: #,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance55;
			summarySettings23.DisplayFormat = "{0: #,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance56;
			summarySettings24.DisplayFormat = "{0: #,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance57;
			summarySettings25.DisplayFormat = "{0: #,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance58;
			summarySettings26.DisplayFormat = "{0: #,##0.00}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings27.Appearance = appearance59;
			summarySettings27.DisplayFormat = "{0: #,##0.00}";
			summarySettings27.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13,
																																															summarySettings14,
																																															summarySettings15,
																																															summarySettings16,
																																															summarySettings17,
																																															summarySettings18,
																																															summarySettings19,
																																															summarySettings20,
																																															summarySettings21,
																																															summarySettings22,
																																															summarySettings23,
																																															summarySettings24,
																																															summarySettings25,
																																															summarySettings26,
																																															summarySettings27});
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.uGridComisiones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridComisiones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridComisiones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridComisiones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance60.BackColor = System.Drawing.Color.Transparent;
			this.uGridComisiones.DisplayLayout.Override.CardAreaAppearance = appearance60;
			appearance61.ForeColor = System.Drawing.Color.Black;
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridComisiones.DisplayLayout.Override.CellAppearance = appearance61;
			appearance62.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance62.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance62.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance62.FontData.BoldAsString = "True";
			appearance62.FontData.Name = "Arial";
			appearance62.FontData.SizeInPoints = 8.5F;
			appearance62.ForeColor = System.Drawing.Color.White;
			appearance62.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridComisiones.DisplayLayout.Override.HeaderAppearance = appearance62;
			this.uGridComisiones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance63.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance63.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance63.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridComisiones.DisplayLayout.Override.RowSelectorAppearance = appearance63;
			appearance64.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance64.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance64.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridComisiones.DisplayLayout.Override.SelectedRowAppearance = appearance64;
			this.uGridComisiones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridComisiones.Location = new System.Drawing.Point(8, 120);
			this.uGridComisiones.Name = "uGridComisiones";
			this.uGridComisiones.Size = new System.Drawing.Size(1232, 240);
			this.uGridComisiones.TabIndex = 723;
			this.uGridComisiones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridComisiones_KeyDown);
			this.uGridComisiones.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridComisiones_ClickCellButton);
			this.uGridComisiones.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridComisiones_AfterCellUpdate);
			this.uGridComisiones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridComisiones_InitializeLayout);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(328, 42);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(76, 17);
			this.label12.TabIndex = 739;
			this.label12.Text = "Departamento";
			// 
			// cmbNECentroCosto
			// 
			this.cmbNECentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECentroCosto.DataSource = this.ultraDataSource7;
			ultraGridBand2.AddButtonCaption = "RolTipo";
			ultraGridColumn60.Header.VisiblePosition = 1;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 85;
			ultraGridColumn61.Header.VisiblePosition = 0;
			ultraGridColumn61.Width = 208;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn60,
																										 ultraGridColumn61});
			this.cmbNECentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNECentroCosto.DisplayMember = "Nombre";
			this.cmbNECentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECentroCosto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECentroCosto.Location = new System.Drawing.Point(104, 40);
			this.cmbNECentroCosto.Name = "cmbNECentroCosto";
			this.cmbNECentroCosto.Size = new System.Drawing.Size(200, 22);
			this.cmbNECentroCosto.TabIndex = 738;
			this.cmbNECentroCosto.ValueMember = "idProyecto";
			// 
			// ultraDataSource7
			// 
			ultraDataColumn55.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn55,
																																 ultraDataColumn56});
			// 
			// cmbNEDepartamento
			// 
			this.cmbNEDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEDepartamento.DataSource = this.ultraDataSource8;
			ultraGridBand3.AddButtonCaption = "RolTipo";
			ultraGridColumn62.Header.VisiblePosition = 0;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 1;
			ultraGridColumn63.Width = 208;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn62,
																										 ultraGridColumn63});
			this.cmbNEDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbNEDepartamento.DisplayMember = "Descripcion";
			this.cmbNEDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEDepartamento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEDepartamento.Location = new System.Drawing.Point(416, 40);
			this.cmbNEDepartamento.Name = "cmbNEDepartamento";
			this.cmbNEDepartamento.Size = new System.Drawing.Size(200, 22);
			this.cmbNEDepartamento.TabIndex = 740;
			this.cmbNEDepartamento.ValueMember = "idDepartamento";
			// 
			// ultraDataSource8
			// 
			ultraDataColumn57.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn57,
																																 ultraDataColumn58});
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(8, 42);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(86, 17);
			this.label33.TabIndex = 737;
			this.label33.Text = "Centro De Costo";
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1248, 8);
			this.groupBox1.TabIndex = 741;
			this.groupBox1.TabStop = false;
			// 
			// btnAcreditacion
			// 
			this.btnAcreditacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAcreditacion.CausesValidation = false;
			this.btnAcreditacion.Image = ((System.Drawing.Image)(resources.GetObject("btnAcreditacion.Image")));
			this.btnAcreditacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAcreditacion.Location = new System.Drawing.Point(788, 8);
			this.btnAcreditacion.Name = "btnAcreditacion";
			this.btnAcreditacion.Size = new System.Drawing.Size(92, 23);
			this.btnAcreditacion.TabIndex = 742;
			this.btnAcreditacion.Text = "Acreditación";
			this.btnAcreditacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAcreditacion.Click += new System.EventHandler(this.btnAcreditacion_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(416, 11);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 16);
			this.lblEstado.TabIndex = 743;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(648, 72);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 16);
			this.lblContador.TabIndex = 744;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// mnuAcreditacion
			// 
			this.mnuAcreditacion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																										this.mnuAcreditacionTransferencias,
																																										this.mnuAcreditacionCheques,
																																										this.menuItem1,
																																										this.menuItem2});
			// 
			// mnuAcreditacionTransferencias
			// 
			this.mnuAcreditacionTransferencias.Index = 0;
			this.mnuAcreditacionTransferencias.Text = "Transferencias";
			this.mnuAcreditacionTransferencias.Click += new System.EventHandler(this.mnuAcreditacionTransferencias_Click);
			// 
			// mnuAcreditacionCheques
			// 
			this.mnuAcreditacionCheques.Index = 1;
			this.mnuAcreditacionCheques.Text = "Cheques";
			this.mnuAcreditacionCheques.Click += new System.EventHandler(this.mnuAcreditacionCheques_Click);
			// 
			// btnAgregar
			// 
			this.btnAgregar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAgregar.CausesValidation = false;
			this.btnAgregar.Enabled = false;
			this.btnAgregar.Image = ((System.Drawing.Image)(resources.GetObject("btnAgregar.Image")));
			this.btnAgregar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAgregar.Location = new System.Drawing.Point(702, 8);
			this.btnAgregar.Name = "btnAgregar";
			this.btnAgregar.TabIndex = 745;
			this.btnAgregar.Text = "Actualizar";
			this.btnAgregar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAgregar.Click += new System.EventHandler(this.btnAgregar_Click);
			// 
			// btnSeleccionar
			// 
			this.btnSeleccionar.CausesValidation = false;
			this.btnSeleccionar.Enabled = false;
			this.btnSeleccionar.Image = ((System.Drawing.Image)(resources.GetObject("btnSeleccionar.Image")));
			this.btnSeleccionar.Location = new System.Drawing.Point(8, 120);
			this.btnSeleccionar.Name = "btnSeleccionar";
			this.btnSeleccionar.Size = new System.Drawing.Size(30, 22);
			this.btnSeleccionar.TabIndex = 746;
			this.btnSeleccionar.Click += new System.EventHandler(this.btnSeleccionar_Click);
			// 
			// btnCerrar
			// 
			this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCerrar.CausesValidation = false;
			this.btnCerrar.Image = ((System.Drawing.Image)(resources.GetObject("btnCerrar.Image")));
			this.btnCerrar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCerrar.Location = new System.Drawing.Point(976, 8);
			this.btnCerrar.Name = "btnCerrar";
			this.btnCerrar.TabIndex = 747;
			this.btnCerrar.Text = "Cerrar";
			this.btnCerrar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
			// 
			// mnuAgregar
			// 
			this.mnuAgregar.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																							 this.mnuAgregarRegistro,
																																							 this.mnuActualizar});
			// 
			// mnuAgregarRegistro
			// 
			this.mnuAgregarRegistro.Index = 0;
			this.mnuAgregarRegistro.Text = "Agregar";
			this.mnuAgregarRegistro.Click += new System.EventHandler(this.mnuAgregarRegistro_Click);
			// 
			// mnuActualizar
			// 
			this.mnuActualizar.Index = 1;
			this.mnuActualizar.Text = "Actualizar";
			this.mnuActualizar.Click += new System.EventHandler(this.mnuActualizar_Click);
			// 
			// txtIdAntipoUtilidades
			// 
			this.txtIdAntipoUtilidades.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAntipoUtilidades.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAntipoUtilidades.Enabled = false;
			this.txtIdAntipoUtilidades.Location = new System.Drawing.Point(584, 8);
			this.txtIdAntipoUtilidades.Name = "txtIdAntipoUtilidades";
			this.txtIdAntipoUtilidades.PromptChar = ' ';
			this.txtIdAntipoUtilidades.Size = new System.Drawing.Size(8, 22);
			this.txtIdAntipoUtilidades.TabIndex = 748;
			this.txtIdAntipoUtilidades.Visible = false;
			// 
			// btnUtilidades
			// 
			this.btnUtilidades.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnUtilidades.CausesValidation = false;
			this.btnUtilidades.Enabled = false;
			this.btnUtilidades.Image = ((System.Drawing.Image)(resources.GetObject("btnUtilidades.Image")));
			this.btnUtilidades.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnUtilidades.Location = new System.Drawing.Point(1064, 8);
			this.btnUtilidades.Name = "btnUtilidades";
			this.btnUtilidades.TabIndex = 749;
			this.btnUtilidades.Text = "Utilidades";
			this.btnUtilidades.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnUtilidades.Click += new System.EventHandler(this.btnUtilidades_Click);
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 2;
			this.menuItem1.Text = "Transferencias Ant Utilidades";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 3;
			this.menuItem2.Text = "Cheque Ant Utilidades";
			this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
			// 
			// frmCom_PagoPersonal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1245, 366);
			this.Controls.Add(this.btnUtilidades);
			this.Controls.Add(this.txtIdAntipoUtilidades);
			this.Controls.Add(this.btnCerrar);
			this.Controls.Add(this.btnSeleccionar);
			this.Controls.Add(this.btnAgregar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnAcreditacion);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cmbNEDepartamento);
			this.Controls.Add(this.label33);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.txtIdComision);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.dtpPeriodo);
			this.Controls.Add(this.uGridComisiones);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.cmbNECentroCosto);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmCom_PagoPersonal";
			this.Text = "Comisiones";
			this.Load += new System.EventHandler(this.frmCom_PagoPersonal_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpPeriodo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridComisiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAntipoUtilidades)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCom_PagoPersonal_Load(object sender, System.EventArgs e)
		{
			DateTime dtFecha = DateTime.Parse("01/06/2016");			
			this.dtpPeriodo.CalendarInfo.MinDate = dtFecha.Date;
			this.dtpPeriodo.CalendarInfo.MaxDate = Convert.ToDateTime(DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month).ToString() + "/" + DateTime.Today.Month.ToString() + "/" + DateTime.Today.Year.ToString());

			this.cmbNECentroCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 45, 0");
			this.cmbNEDepartamento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 48, 0");

			this.dtpPeriodo.Value = DateTime.Today.AddDays(-31);

			if (bComisionQuincena)
			{
				//				string nombrCol = "";
							
				int iColumnas = (int)this.uGridComisiones.DisplayLayout.Bands[0].Columns.Count;
				
				for (int j = 35; j < iColumnas; j++)
				{
					//					nombrCol = this.uGridComisiones.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						
					this.uGridComisiones.DisplayLayout.Bands[0].Columns[j].Hidden = true;						
				}							
			}
			menuItem1.Visible = false;
			menuItem2.Visible = false;
			VerificaMenuAnticipoUtilidades();
		}

		string sComisionesQuincena = "COMISIONES";

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			int iMes = Convert.ToDateTime(this.dtpPeriodo.Value).Month;
			int iAño = Convert.ToDateTime(this.dtpPeriodo.Value).Year;
			string sPeriodo = iMes.ToString() + "/" + iAño.ToString();

			string sSQLValida = string.Format("Exec Com_ValidaPeriodoCerrado '{0}'", sPeriodo);
			if (!Funcion.bEscalarSQL(cdsSeteoF, sSQLValida, true) && !MenuLatinium.bComisionQuincena)
			{
				MessageBox.Show(string.Format("No se han generado las comisiones del mes de {0}\n\nConsulte con Gerencia Comercial", Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();
				return;
			}
			
			if (MenuLatinium.bComisionQuincena) sComisionesQuincena = "QUINCENA";

			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de GENERAR EL PAGO DE {1} DEL PERIODO DE '{0}'", 
				Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper(), sComisionesQuincena), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				#region Comisones 
				string sSQL = string.Format("Exec Com_GuardaEncabezado {0}, '{1}'", (int)this.txtIdComision.Value, Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"));
				this.txtIdComision.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
	
				string sSQLDetalle = string.Format("Exec Com_GuardaDetalleComisiones '{0}', {1}, {2}", 
					Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"), (int)this.txtIdComision.Value, MenuLatinium.bComisionQuincena);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
				#endregion Comisones

				MessageBox.Show(string.Format("{1} del Periodo de '{0}' Generado Correctamente", 
					Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper(), sComisionesQuincena), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.dtpPeriodo_ValueChanged(sender, e);

				this.btnVer_Click(sender, e);				
			}			
		}

		private void dtpPeriodo_ValueChanged(object sender, System.EventArgs e)
		{						
			this.btnGenerar.Enabled = false;			
			//this.btnAcreditacion.Enabled = false;
			this.btnExportar.Enabled = false;			
			this.btnAgregar.Enabled = false;
			this.btnSeleccionar.Enabled = false;
			this.btnCerrar.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(false, this.uGridComisiones);
			
			string sSQL = string.Format("Exec Com_ValidaComisionesExiste '{0}'", Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"));			
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdComision.Value = dr.GetInt32(0);
				this.lblEstado.Text = dr.GetString(1);
				Estado = dr.GetInt32(2);	
				this.txtIdAntipoUtilidades.Value = dr.GetInt32(3);
				dr.Close();

				if (Estado == 1)
				{
					this.btnAcreditacion.Enabled = true;
					this.btnExportar.Enabled = true;
					this.btnAgregar.Enabled = true;
					this.btnSeleccionar.Enabled = true;
				
					FuncionesProcedimientos.EstadoGrids(true, this.uGridComisiones);
				}
				else if (Estado == 2) 
				{
					this.btnCerrar.Enabled = true;
					this.btnExportar.Enabled = true;
				}
				else if (Estado == 3) 
				{
					this.btnExportar.Enabled = true;
//					if (MenuLatinium.stIdDB == 1 || MenuLatinium.stIdDB == 33|| MenuLatinium.stIdDB == 63)	
//					{
//						this.btnUtilidades.Enabled = true;
//					}
				}
				else this.btnExportar.Enabled = true;								
			}
			else
			{
				dr.Close();				
				this.txtIdComision.Value = 0;
				this.lblEstado.Text = "PENDIENTE DE GENERAR";

				FuncionesProcedimientos.EstadoGrids(true, this.uGridComisiones);

				this.btnGenerar.Enabled = true;				
			}
			dr.Close();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int idProyecto = 0;
			int idDepartamento = 0;

			if (this.cmbNECentroCosto.ActiveRow != null) idProyecto = (int)this.cmbNECentroCosto.Value;
			if (this.cmbNEDepartamento.ActiveRow != null) idDepartamento = (int)this.cmbNEDepartamento.Value;

			string sSQL = string.Format("Exec Com_ConsultaDetalleComisiones {0}, {1}, {2}, '{3}'", (int)this.txtIdComision.Value, idProyecto, idDepartamento, this.txtBuscar.Text.ToString());
			this.uGridComisiones.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

			this.lblContador.Text = this.uGridComisiones.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void uGridComisiones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Infraestructura", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Cobranzas", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Operaciones", 2);
			FuncionesProcedimientos.FormatoGrid(e, "RecursosHumanos", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Bodega", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Administrativo", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Ventas", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Financiero", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Credito", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Cobranzas", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Compras", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Contabilidad", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Sistemas", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Marketing", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Auditoria", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DepartamentoTecnico", 2);

			FuncionesProcedimientos.FormatoGrid(e, "Comisiones", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AnticipoUtilidades", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ComprasCredito", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ComprasCrediPoint", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FaltantesCaja", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Multas", 2);			
			FuncionesProcedimientos.FormatoGrid(e, "AnticiposSueldos", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SeguroMedico", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Farmacia", 2);			
			FuncionesProcedimientos.FormatoGrid(e, "Otros", 2);
		}

		private void uGridComisiones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridComisiones);
		}

		private void dSumatoria()
		{								
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridComisiones.Rows.All)
			{
				dr.Cells["Comisiones"].Value = 				
					Convert.ToDecimal(dr.Cells["Infraestructura"].Value) +
					Convert.ToDecimal(dr.Cells["Cobranzas"].Value) +
					Convert.ToDecimal(dr.Cells["Operaciones"].Value) +
					Convert.ToDecimal(dr.Cells["RecursosHumanos"].Value) +
					Convert.ToDecimal(dr.Cells["Bodega"].Value) +
					Convert.ToDecimal(dr.Cells["Administrativo"].Value) +
					Convert.ToDecimal(dr.Cells["Ventas"].Value) +
					Convert.ToDecimal(dr.Cells["Financiero"].Value) +
					Convert.ToDecimal(dr.Cells["Credito"].Value) +
					Convert.ToDecimal(dr.Cells["Compras"].Value) +
					Convert.ToDecimal(dr.Cells["Contabilidad"].Value) +
					Convert.ToDecimal(dr.Cells["Sistemas"].Value) +
					Convert.ToDecimal(dr.Cells["Marketing"].Value) +
					Convert.ToDecimal(dr.Cells["Auditoria"].Value) +
					Convert.ToDecimal(dr.Cells["DepartamentoTecnico"].Value);

				dr.Cells["TotalEgresos"].Value = 				
					Convert.ToDecimal(dr.Cells["ComprasCrediPoint"].Value) +
					Convert.ToDecimal(dr.Cells["ComprasCredito"].Value) +
					Convert.ToDecimal(dr.Cells["FaltantesCaja"].Value) +
					Convert.ToDecimal(dr.Cells["Multas"].Value) +
					Convert.ToDecimal(dr.Cells["AnticiposSueldos"].Value) +
					Convert.ToDecimal(dr.Cells["SeguroMedico"].Value) +
					Convert.ToDecimal(dr.Cells["Farmacia"].Value) +
					Convert.ToDecimal(dr.Cells["Otros"].Value);

				ActualizaValores("Comisiones", Convert.ToDecimal(dr.Cells["Comisiones"].Value), (int)dr.Cells["idDetalleComision"].Value);
				dr.Cells["Total"].Value = Convert.ToDecimal(dr.Cells["Comisiones"].Value) - Convert.ToDecimal(dr.Cells["TotalEgresos"].Value);
				
				ActualizaValores("TotalEgresos", Convert.ToDecimal(dr.Cells["TotalEgresos"].Value), (int)dr.Cells["idDetalleComision"].Value);
				ActualizaValores("Total", Convert.ToDecimal(dr.Cells["Total"].Value), (int)dr.Cells["idDetalleComision"].Value);
			}
		}

		private void Sumatorias(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			decimal totalBase = 
				Convert.ToDecimal(e.Cell.Row.Cells["Infraestructura"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Cobranzas"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Operaciones"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["RecursosHumanos"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Bodega"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Administrativo"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Ventas"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Financiero"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Credito"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Compras"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Contabilidad"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Sistemas"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Marketing"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Auditoria"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["DepartamentoTecnico"].Value);

			e.Cell.Row.Cells["AnticipoUtilidades"].Value = Decimal.Round(totalBase * 0.2m, 2);
			e.Cell.Row.Cells["Comisiones"].Value = Decimal.Round(totalBase * 0.8m, 2);

			e.Cell.Row.Cells["TotalEgresos"].Value = 				
				Convert.ToDecimal(e.Cell.Row.Cells["ComprasCrediPoint"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["ComprasCredito"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["FaltantesCaja"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Multas"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["AnticiposSueldos"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["SeguroMedico"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Farmacia"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Otros"].Value);

			e.Cell.Row.Cells["Total"].Value = 
				Convert.ToDecimal(e.Cell.Row.Cells["Comisiones"].Value) - 
				Convert.ToDecimal(e.Cell.Row.Cells["TotalEgresos"].Value);

			ActualizaValores("AnticipoUtilidades", Convert.ToDecimal(e.Cell.Row.Cells["AnticipoUtilidades"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
			ActualizaValores("Comisiones", Convert.ToDecimal(e.Cell.Row.Cells["Comisiones"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
			ActualizaValores("TotalEgresos", Convert.ToDecimal(e.Cell.Row.Cells["TotalEgresos"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
			ActualizaValores("Total", Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
		}

		private void ActualizaValores(string sCampo, decimal dValor, int idDetalleComision)
		{
			string sSQL = string.Format("Exec Com_ActualizaValores '{0}', {1}, {2}", sCampo, dValor, idDetalleComision);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void uGridComisiones_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string nombreColumna = e.Cell.Column.ToString();

			if (e.Cell.Column.ToString() == "Revisado") 
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaActualizaRevisado {0}, {1}, 1", (bool)e.Cell.Row.Cells["Revisado"].Value, (int)e.Cell.Row.Cells["idDetalleComision"].Value));				
			}

			if (
				nombreColumna == "Infraestructura" || nombreColumna == "Cobranzas" ||
				nombreColumna == "Operaciones" || nombreColumna == "RecursosHumanos" ||
				nombreColumna == "Bodega" || nombreColumna == "Administrativo" ||
				nombreColumna == "Ventas" || nombreColumna == "Financiero" ||
				nombreColumna == "Credito" || nombreColumna == "Compras" ||
				nombreColumna == "Contabilidad" || nombreColumna == "Sistemas" ||
				nombreColumna == "Marketing" || nombreColumna == "Auditoria" ||
				nombreColumna == "DepartamentoTecnico"
				)
			{
				decimal valor = Convert.ToDecimal(e.Cell.Row.Cells[nombreColumna].Value);
				decimal anticipo = Decimal.Round(valor * 0.2m, 2);
				decimal comision = Decimal.Round(valor * 0.8m, 2);

				if (e.Cell.Row.Cells["AnticipoUtilidades"].Value == null || 
					(decimal)e.Cell.Row.Cells["AnticipoUtilidades"].Value != anticipo)
				{
					e.Cell.Row.Cells["AnticipoUtilidades"].Value = anticipo;
				}

				if (e.Cell.Row.Cells["Comisiones"].Value == null || 
					(decimal)e.Cell.Row.Cells["Comisiones"].Value != comision)
				{
					e.Cell.Row.Cells["Comisiones"].Value = comision;
				}

				Sumatorias(e);
				ActualizaValores(nombreColumna, valor, (int)e.Cell.Row.Cells["idDetalleComision"].Value);
				return;
			}

			if (e.Cell.Column.ToString() == "ComprasCredito") 
			{
				Sumatorias(e);
				ActualizaValores("ComprasCredito", Convert.ToDecimal(e.Cell.Row.Cells["ComprasCredito"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
			}

			if (e.Cell.Column.ToString() == "ComprasCrediPoint") 
			{
				Sumatorias(e);
				ActualizaValores("ComprasCrediPoint", Convert.ToDecimal(e.Cell.Row.Cells["ComprasCrediPoint"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
			}

			if (e.Cell.Column.ToString() == "FaltantesCaja") 
			{
				Sumatorias(e);
				ActualizaValores("FaltantesCaja", Convert.ToDecimal(e.Cell.Row.Cells["FaltantesCaja"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
			}

			if (e.Cell.Column.ToString() == "Multas") 
			{
				Sumatorias(e);
				ActualizaValores("Multas", Convert.ToDecimal(e.Cell.Row.Cells["Multas"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
			}

			if (e.Cell.Column.ToString() == "AnticiposSueldos") 
			{
				Sumatorias(e);
				ActualizaValores("AnticiposSueldos", Convert.ToDecimal(e.Cell.Row.Cells["AnticiposSueldos"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
			}

			if (e.Cell.Column.ToString() == "SeguroMedico") 
			{
				Sumatorias(e);
				ActualizaValores("SeguroMedico", Convert.ToDecimal(e.Cell.Row.Cells["SeguroMedico"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
			}

			if (e.Cell.Column.ToString() == "Farmacia") 
			{
				Sumatorias(e);
				ActualizaValores("Farmacia", Convert.ToDecimal(e.Cell.Row.Cells["Farmacia"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
			}

			if (e.Cell.Column.ToString() == "Otros") 
			{
				Sumatorias(e);
				ActualizaValores("Otros", Convert.ToDecimal(e.Cell.Row.Cells["Otros"].Value), (int)e.Cell.Row.Cells["idDetalleComision"].Value);
			}
		}

		private void mnuAcreditacionTransferencias_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());

			using (frmRolAcreditacion RAC = new frmRolAcreditacion((int)this.txtIdComision.Value, 0, dtFecha, 1))
			{
				if (DialogResult.OK == RAC.ShowDialog())
				{
					this.btnAgregar.Enabled = false;
				}

				this.dtpPeriodo_ValueChanged(sender, e);
			}			
		}

		private void mnuAcreditacionCheques_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());

			using (frmRolAutorizacionPagoCheques RAPC = new frmRolAutorizacionPagoCheques((int)this.txtIdComision.Value, dtFecha, 1))
			{
				if (DialogResult.OK == RAPC.ShowDialog())
				{
					this.btnAgregar.Enabled = false;
				}

				this.dtpPeriodo_ValueChanged(sender, e);
			}
		}

		private void dtpPeriodo_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnAcreditacion_Click(object sender, System.EventArgs e)
		{
			this.mnuAcreditacion.Show(this.btnAcreditacion, new Point(1, 25));
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridComisiones);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{				
			this.Close();
		}

		private void txtBuscar_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnAgregar_Click(object sender, System.EventArgs e)
		{
			this.mnuAgregar.Show(this.btnAgregar, new Point(1, 25));	
		}

		private void btnSeleccionar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridComisiones.Rows.All)
			{	
				dr.Cells["Revisado"].Value = !(bool)dr.Cells["Revisado"].Value;
			}
		}

		private void uGridComisiones_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnCredito")
			{	
				if (Convert.ToDecimal(e.Cell.Row.Cells["ComprasCredito"].Value) == 0)
				{
					MessageBox.Show("El valor de Credito es CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}

				using (frmRolFactura Busqueda = new frmRolFactura((int)e.Cell.Row.Cells["idCompraC"].Value, e.Cell.Row.Cells["ObservacionesCredito"].Value.ToString()))
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleComisiones Set ObservacionesCredito = '{0}', idCompraC = {1} Where idDetalleComision = {2}", 
							Busqueda.txtObservaciones.Text.ToString().Trim(), (int)Busqueda.txtNumIdCompra.Value, (int)e.Cell.Row.Cells["idDetalleComision"].Value));

						e.Cell.Row.Cells["ObservacionesCredito"].Value = Busqueda.txtObservaciones.Text.ToString().Trim();
					}
				}
			}

			#region Faltantes de Caja
			if (e.Cell.Column.ToString() == "btnFC")
			{	
				if (Convert.ToDecimal(e.Cell.Row.Cells["FaltantesCaja"].Value) == 0)
				{
					MessageBox.Show("El valor de Faltantes de Caja es CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}

				using (frmNominaObservaciones Observaciones = new frmNominaObservaciones(1, e.Cell.Row.Cells["ObservacionesFaltCaja"].Value.ToString().Trim()))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleComisiones Set ObservacionesFaltCaja = '{0}' Where idDetalleComision = {1}", 
							Observaciones.txtObservaciones.Text.ToString().Trim(), (int)e.Cell.Row.Cells["idDetalleComision"].Value));
					}
				}
			}
			#endregion Faltantes de Caja
			
			#region Multas
			if (e.Cell.Column.ToString() == "btnMultas")
			{	
				if (Convert.ToDecimal(e.Cell.Row.Cells["Multas"].Value) == 0)
				{
					MessageBox.Show("El valor de Multas es CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}

				using (frmNominaObservaciones Observaciones = new frmNominaObservaciones(2, e.Cell.Row.Cells["ObservacionesMultas"].Value.ToString().Trim()))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleComisiones Set ObservacionesMultas = '{0}' Where idDetalleComision = {1}", 
							Observaciones.txtObservaciones.Text.ToString().Trim(), (int)e.Cell.Row.Cells["idDetalleComision"].Value));
					}
				}
			}
			#endregion Multas

			#region Anticipos
			if (e.Cell.Column.ToString() == "btnAnticipos")
			{	
				if (Convert.ToDecimal(e.Cell.Row.Cells["AnticiposSueldos"].Value) == 0)
				{
					MessageBox.Show("El valor de Anticipos es CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}

				using (frmNominaObservaciones Observaciones = new frmNominaObservaciones(4, e.Cell.Row.Cells["ObservacionesAnticipos"].Value.ToString().Trim()))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleComisiones Set ObservacionesAnticipos = '{0}' Where idDetalleComision = {1}", 
							Observaciones.txtObservaciones.Text.ToString().Trim(), (int)e.Cell.Row.Cells["idDetalleComision"].Value));
					}
				}
			}
			#endregion Anticipos

			#region Otros Descuentos
			if (e.Cell.Column.ToString() == "btnOtros")
			{	
				if (Convert.ToDecimal(e.Cell.Row.Cells["Otros"].Value) == 0)
				{
					MessageBox.Show("El valor de Otros descuentos es CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}

				using (frmNominaObservaciones Observaciones = new frmNominaObservaciones(3, e.Cell.Row.Cells["ObservacionesOtros"].Value.ToString().Trim()))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleComisiones Set ObservacionesOtros = '{0}' Where idDetalleComision = {1}", 
							Observaciones.txtObservaciones.Text.ToString().Trim(), (int)e.Cell.Row.Cells["idDetalleComision"].Value));
					}
				}
			}
			#endregion Otros Descuentos
		}

		private void btnCerrar_Click(object sender, System.EventArgs e)
		{
			if (Estado == 2)
			{
				if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de CERRAR EL PAGO DE {1} DEL PERIODO DE '{0}'", 
					Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper(), sComisionesQuincena), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				{
					int iContadorCrediPoint = 0;
					int idCompra = 0;
					int idCliente = 0;
					decimal dAbono = 0.00m;
					string sObservaciones = "";
					string sNumero = "";
					int idAnticipo = 0;
					DataSet oDSCP = new DataSet();
					cdsSeteoF.Schema.Connections[0].Open();
					C1.Data.SchemaObjects.Connection oConexion = cdsSeteoF.Schema.Connections[0];
					oConexion = cdsSeteoF.Schema.Connections[0];

					#region Cobros
					string sSQLCP = string.Format("Exec Com_ListaCobrosCrediPoint {0}", (int)this.txtIdComision.Value);
					SqlDataAdapter oDACP = new SqlDataAdapter(sSQLCP, (SqlConnection) oConexion.DbConnection);
					oDACP.Fill(oDSCP);

					if (oDSCP.Tables[0].Rows.Count >= 1)
					{					
						foreach(DataRow oDRCP in oDSCP.Tables[0].Rows)
						{	
							idAnticipo = 0;
							sNumero = "";

							idCompra = (int)oDRCP["idCompra"];
							idCliente = (int)oDRCP["idCliente"];
							dAbono = (decimal)oDRCP["Valor"];						
							sObservaciones = oDRCP["Observaciones"].ToString();
															
							sNumero = Funcion.sEscalarSQL(cdsSeteoF, "Exec NumeracionLocales 24, 57, 0");

							string sSQLGraba = string.Format("Exec GuardaAnticipos {0}, {1}, '{2}', {3}, '{4}', {5}, {6}, {7}, '{8}', {9}, {10}, {11}, '{12}', {13}, 0", 
								idAnticipo, 0, DateTime.Today.ToString("yyyyMMdd"), 57, sNumero, idCliente, 18, 0, sObservaciones, 1, dAbono, 12, "", idCompra);
							idAnticipo = Funcion.iEscalarSQL(cdsSeteoF, sSQLGraba);

							string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, 0", 
								0, idAnticipo, 2, 18, dAbono, 0, 0, "", "", DateTime.Today.ToString("yyyyMMdd"),	"", "", 0, "", "", 0, 0, 0);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro);
						
							string sSQLAsiento = string.Format("Exec CreaAsientoAnticipo 0, {0}, 12, 2", 0, idAnticipo, 12, 2);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);

							iContadorCrediPoint++;
						}					
					}
					#endregion Cobros

					#region Utilidades
					string sSQLUtilidades = string.Format("Exec Com_GuardaEncabezadoAnticipoUtilidades {0}, '{1}'", (int)this.txtIdAntipoUtilidades.Value, Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"));
					this.txtIdAntipoUtilidades.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLUtilidades);
	
					string sSQLDetalleUtilidades = string.Format("Exec Com_GuardaDetalleAnticipoUtilidades {0},{1}",(int)this.txtIdComision.Value, (int)this.txtIdAntipoUtilidades.Value, MenuLatinium.bComisionQuincena);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalleUtilidades);

					#endregion Utilidades

					Estado = 3;

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaActualizaEstado {0}",  (int)this.txtIdComision.Value));

					MessageBox.Show(string.Format("Se generaron {0} pagos CrediPoint Correctamente.", iContadorCrediPoint));
	
					dtpPeriodo_ValueChanged(sender, e);
				}

			}			
		}

		private void mnuAgregarRegistro_Click(object sender, System.EventArgs e)
		{
		int iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Com_CuentaFaltantesPagoNomina {0}, '{1}', {2}", 
				(int)this.txtIdComision.Value, Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"), MenuLatinium.bComisionQuincena));

			if (iCont == 0)
			{
				MessageBox.Show("No existe personal nuevo para agregar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			string sSQLDetalle = string.Format("Exec Com_GuardaDetalleComisionesFaltantes '{0}', {1}, {2}", 
				Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"), (int)this.txtIdComision.Value, MenuLatinium.bComisionQuincena);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);

			MessageBox.Show(string.Format("Se agregaron {0} nuevos registros correctamente", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.btnVer_Click(sender, e);				
		}

		private void mnuActualizar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbValidaSelGrids(this.uGridComisiones, "Revisado", "Seleccione al menos un registro para actualizar")) return;

			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de ACTUALIZAR LA INFORMACIÓN DEL PERSONAL SELECCIONADO DEL PERIODO DE '{0}'", 
				Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				this.Cursor = Cursors.WaitCursor;
						
				string sSQLActualiza = string.Format("Exec Com_ActualizaParametros {0}", (int)this.txtIdComision.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLActualiza);
				MessageBox.Show(string.Format("Comisiones del Periodo de '{0}' Actualizado Correctamente", Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnVer_Click(sender, e);

				this.Cursor = Cursors.Default;
			}	
		}

		private void btnUtilidades_Click(object sender, System.EventArgs e)
		{
			return;

			using (frmRolAnticipoUtilidades AntUti = new frmRolAnticipoUtilidades((int)this.txtIdAntipoUtilidades.Value,5))
			{
				if (DialogResult.OK == AntUti.ShowDialog())
				{
		
				}

				this.dtpPeriodo_ValueChanged(sender, e);
			}
	
		}
		private void VerificaMenuAnticipoUtilidades()
		{
			string sSQL = string.Format("select AnticipoUtilidades from Seteo");
			bool iEstadoUtl = Funcion.bEscalarSQL(cdsSeteoF, sSQL,true);

			menuItem1.Visible = iEstadoUtl;
			menuItem2.Visible = iEstadoUtl;
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());

			using (frmRolAcreditacion RAC = new frmRolAcreditacion((int)this.txtIdComision.Value, 0, dtFecha,5))
			{
				if (DialogResult.OK == RAC.ShowDialog())
				{
		
				}

				this.dtpPeriodo_ValueChanged(sender, e);
			}	
		}

		private void menuItem2_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());

			using (frmRolAutorizacionPagoCheques RAPC = new frmRolAutorizacionPagoCheques((int)this.txtIdComision.Value, dtFecha, 5))
			{
				if (DialogResult.OK == RAPC.ShowDialog())
				{
		
				}
				
				this.dtpPeriodo_ValueChanged(sender, e);
			}
		}
	}
}
