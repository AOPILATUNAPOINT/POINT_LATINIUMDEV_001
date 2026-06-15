using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRevisiondeCedulas.
	/// </summary>
	public class frmCred_RevisiondeCedulas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnver;
		private System.Windows.Forms.Label lblEntidad;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDocumento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidoP;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidoM;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPnombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSnombre;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkWeb;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtScore;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdCalificacion1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPichicncha;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.Label label16;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid gdHistorial;
		private Infragistics.Win.UltraWinGrid.UltraGrid gdObservaciones;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCre_DatosGenerales;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCre_solicitud;
				
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaValor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdEntidad;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdCalificacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;



		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;


		#region Variables
		int IdDocumento=0;
		int IdTipoRuc= 0;		
		int iBodega = 0;  
		int iCalificacion=0;
		int Calificacion1=0;
		int Bodegas = 0;
		DateTime dtFechas;
		int iEstadoConteo = 0;
		bool bEliminaAlValidar = false;
		bool bNuevo = false;
		bool bEdicion = false;	
		bool bActivar=false;
		string scedula="";
		#endregion Variables
			
		string sCliente = "";
		bool bCarga = false;
		bool bVista;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbidTipoCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbdetallecliente;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaVivienda;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIngreso;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource13;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoVivienda;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optDependiente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFuenteIngresos;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label23;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaNacimiento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource14;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSituacionLaboral;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdEstadoSolicitud;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCupo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNuovo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMarca;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label25;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource15;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource16;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbModelo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ChkOmite;
		bool bVistaReiterativo = false;

		

		public frmCred_RevisiondeCedulas(string sCedula, bool BVista)
		{
			InitializeComponent();
				scedula = sCedula;
			bVista = BVista;
		}

		public frmCred_RevisiondeCedulas(string sCedula, bool BVista, bool BVistaReiterativo)
		{
			InitializeComponent();
				scedula = sCedula;
			bVista = BVista;
			bVistaReiterativo = BVistaReiterativo;
		}

		public frmCred_RevisiondeCedulas(string sCedula)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			scedula = sCedula;
						
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmCred_RevisiondeCedulas()
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
		/// 

		protected override void Dispose( bool disposing )
		{
			if ( disposing )
			{
				if (components != null)
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCred_RevisiondeCedulas));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("E_Financiera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Pend");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Pag");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuotas");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPrimerPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasPendientes");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasPagadas");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCapital");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromedioRetraso");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ultimo_Pago");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPrimerPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasPendientes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasPagadas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoCapital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromedioRetraso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ultimo_Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ent_Fin");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPrimerPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasPendiente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasPagos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoCapital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromedioRetro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UltimoPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Score");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupo");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pantalla");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Score");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pantalla");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pantalla");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleTipoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleTipoCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoVivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNegadoPendiente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNegadoPendiente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idKnox");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idKnox");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idKNOXModelo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Modelo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idKNOXModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDocumento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnver = new System.Windows.Forms.Button();
			this.lblEntidad = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtApellidoP = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtApellidoM = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPnombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSnombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.chkWeb = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtScore = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmdCalificacion1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtPichicncha = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.gdHistorial = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.gdObservaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtidCre_DatosGenerales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidCre_solicitud = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.txtCuotaValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmdEntidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmdCalificacion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbidTipoCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbdetallecliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label11 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtFechaVivienda = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.dtFechaIngreso = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource13 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipoVivienda = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmdEstadoSolicitud = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource14 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label21 = new System.Windows.Forms.Label();
			this.optDependiente = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtFuenteIngresos = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label22 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.dtFechaNacimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbSituacionLaboral = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtCupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkNuovo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbMarca = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource15 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label24 = new System.Windows.Forms.Label();
			this.label25 = new System.Windows.Forms.Label();
			this.ultraDataSource16 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbModelo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ChkOmite = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoM)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPnombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSnombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtScore)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdCalificacion1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPichicncha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gdHistorial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gdObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCre_DatosGenerales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCre_solicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdEntidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdCalificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidTipoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbdetallecliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdEstadoSolicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optDependiente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFuenteIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSituacionLaboral)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource15)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource16)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 310;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm";
			this.dtFecha.Location = new System.Drawing.Point(64, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 309;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(184, 11);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(86, 17);
			this.lblBodega.TabIndex = 345;
			this.lblBodega.Text = "Tipo Documento";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(456, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(61, 17);
			this.label1.TabIndex = 348;
			this.label1.Text = "Documento";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDocumento
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDocumento.Appearance = appearance2;
			this.txtDocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDocumento.Location = new System.Drawing.Point(536, 8);
			this.txtDocumento.MaxLength = 10;
			this.txtDocumento.Name = "txtDocumento";
			this.txtDocumento.Size = new System.Drawing.Size(120, 22);
			this.txtDocumento.TabIndex = 349;
			this.txtDocumento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDocumento_KeyDown);
			this.txtDocumento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDocumento_KeyPress);
			this.txtDocumento.Validating += new System.ComponentModel.CancelEventHandler(this.txtDocumento_Validating);
			this.txtDocumento.Validated += new System.EventHandler(this.txtDocumento_Validated);
			this.txtDocumento.ValueChanged += new System.EventHandler(this.txtDocumento_ValueChanged);
			// 
			// btnver
			// 
			this.btnver.Image = ((System.Drawing.Image)(resources.GetObject("btnver.Image")));
			this.btnver.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnver.Location = new System.Drawing.Point(672, 8);
			this.btnver.Name = "btnver";
			this.btnver.Size = new System.Drawing.Size(104, 23);
			this.btnver.TabIndex = 350;
			this.btnver.Text = "Ver Documento";
			this.btnver.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnver.Click += new System.EventHandler(this.btnver_Click);
			// 
			// lblEntidad
			// 
			this.lblEntidad.AutoSize = true;
			this.lblEntidad.Location = new System.Drawing.Point(8, 200);
			this.lblEntidad.Name = "lblEntidad";
			this.lblEntidad.Size = new System.Drawing.Size(96, 17);
			this.lblEntidad.TabIndex = 351;
			this.lblEntidad.Text = "Entidad Financiera";
			this.lblEntidad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 80);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(85, 17);
			this.label2.TabIndex = 353;
			this.label2.Text = "Apellido Paterno";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(184, 80);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 17);
			this.label3.TabIndex = 354;
			this.label3.Text = "Apellido Materno";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(360, 80);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 17);
			this.label4.TabIndex = 355;
			this.label4.Text = "Primer Nombre";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(536, 80);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(91, 17);
			this.label5.TabIndex = 356;
			this.label5.Text = "Segundo Nombre";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 40);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(30, 17);
			this.label7.TabIndex = 357;
			this.label7.Text = "Local";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtApellidoP
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoP.Appearance = appearance3;
			this.txtApellidoP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoP.Enabled = false;
			this.txtApellidoP.Location = new System.Drawing.Point(8, 104);
			this.txtApellidoP.Name = "txtApellidoP";
			this.txtApellidoP.Size = new System.Drawing.Size(170, 22);
			this.txtApellidoP.TabIndex = 358;
			this.txtApellidoP.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApellidoP_KeyPress);
			// 
			// txtApellidoM
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoM.Appearance = appearance4;
			this.txtApellidoM.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoM.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoM.Enabled = false;
			this.txtApellidoM.Location = new System.Drawing.Point(184, 104);
			this.txtApellidoM.Name = "txtApellidoM";
			this.txtApellidoM.Size = new System.Drawing.Size(170, 22);
			this.txtApellidoM.TabIndex = 359;
			this.txtApellidoM.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApellidoM_KeyPress);
			// 
			// txtPnombre
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPnombre.Appearance = appearance5;
			this.txtPnombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPnombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPnombre.Enabled = false;
			this.txtPnombre.Location = new System.Drawing.Point(360, 104);
			this.txtPnombre.Name = "txtPnombre";
			this.txtPnombre.Size = new System.Drawing.Size(170, 22);
			this.txtPnombre.TabIndex = 360;
			this.txtPnombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPnombre_KeyPress);
			// 
			// txtSnombre
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSnombre.Appearance = appearance6;
			this.txtSnombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSnombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSnombre.Enabled = false;
			this.txtSnombre.Location = new System.Drawing.Point(536, 104);
			this.txtSnombre.Name = "txtSnombre";
			this.txtSnombre.Size = new System.Drawing.Size(170, 22);
			this.txtSnombre.TabIndex = 361;
			this.txtSnombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSnombre_KeyPress);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 240);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(60, 17);
			this.label8.TabIndex = 363;
			this.label8.Text = "Calificación";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(192, 240);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(32, 17);
			this.label10.TabIndex = 365;
			this.label10.Text = "Score";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(360, 240);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(30, 17);
			this.label12.TabIndex = 367;
			this.label12.Text = "Cupo";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkWeb
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkWeb.Appearance = appearance7;
			this.chkWeb.BackColor = System.Drawing.Color.Transparent;
			this.chkWeb.CausesValidation = false;
			this.chkWeb.Enabled = false;
			this.chkWeb.Location = new System.Drawing.Point(528, 240);
			this.chkWeb.Name = "chkWeb";
			this.chkWeb.Size = new System.Drawing.Size(56, 22);
			this.chkWeb.TabIndex = 637;
			this.chkWeb.Text = "WEB";
			this.chkWeb.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkWeb_KeyDown);
			// 
			// txtScore
			// 
			this.txtScore.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtScore.Enabled = false;
			this.txtScore.FormatString = "#,##0";
			this.txtScore.Location = new System.Drawing.Point(240, 240);
			this.txtScore.MaskInput = "nnn";
			this.txtScore.MaxValue = 999;
			this.txtScore.MinValue = 0;
			this.txtScore.Name = "txtScore";
			this.txtScore.PromptChar = ' ';
			this.txtScore.Size = new System.Drawing.Size(100, 22);
			this.txtScore.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtScore.TabIndex = 638;
			this.txtScore.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtScore_KeyDown);
			this.txtScore.Click += new System.EventHandler(this.txtScore_Click);
			this.txtScore.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtScore_KeyPress);
			this.txtScore.Enter += new System.EventHandler(this.txtScore_Enter);
			// 
			// cmdCalificacion1
			// 
			this.cmdCalificacion1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmdCalificacion1.CausesValidation = false;
			this.cmdCalificacion1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmdCalificacion1.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmdCalificacion1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmdCalificacion1.DisplayMember = "";
			this.cmdCalificacion1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdCalificacion1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdCalificacion1.Enabled = false;
			this.cmdCalificacion1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdCalificacion1.Location = new System.Drawing.Point(492, 605);
			this.cmdCalificacion1.MaxDropDownItems = 30;
			this.cmdCalificacion1.Name = "cmdCalificacion1";
			this.cmdCalificacion1.Size = new System.Drawing.Size(96, 21);
			this.cmdCalificacion1.TabIndex = 640;
			this.cmdCalificacion1.ValueMember = "";
			this.cmdCalificacion1.Visible = false;
			this.cmdCalificacion1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmdCalificacion1_KeyDown);
			// 
			// txtPichicncha
			// 
			this.txtPichicncha.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtPichicncha.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPichicncha.Enabled = false;
			this.txtPichicncha.FormatString = "#,##0.00";
			this.txtPichicncha.Location = new System.Drawing.Point(416, 605);
			this.txtPichicncha.MaskInput = "nn";
			this.txtPichicncha.MaxValue = 1000;
			this.txtPichicncha.MinValue = 0;
			this.txtPichicncha.Name = "txtPichicncha";
			this.txtPichicncha.PromptChar = ' ';
			this.txtPichicncha.Size = new System.Drawing.Size(56, 22);
			this.txtPichicncha.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtPichicncha.TabIndex = 641;
			this.txtPichicncha.Visible = false;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(280, 200);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(38, 17);
			this.label13.TabIndex = 643;
			this.label13.Text = "Estado";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCelular
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance8;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Location = new System.Drawing.Point(720, 104);
			this.txtCelular.MaxLength = 10;
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(120, 22);
			this.txtCelular.TabIndex = 647;
			this.txtCelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCelular_KeyDown);
			this.txtCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCelular_KeyPress);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Enabled = false;
			this.label15.Location = new System.Drawing.Point(720, 80);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(38, 17);
			this.label15.TabIndex = 648;
			this.label15.Text = "Celular";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservaciones
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance9;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(104, 272);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(400, 80);
			this.txtObservaciones.TabIndex = 649;
			this.txtObservaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtObservaciones_KeyDown);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(8, 304);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(76, 17);
			this.label16.TabIndex = 650;
			this.label16.Text = "Observaciones";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource1
			// 
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
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
			// gdHistorial
			// 
			this.gdHistorial.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gdHistorial.Cursor = System.Windows.Forms.Cursors.Default;
			this.gdHistorial.DataSource = this.ultraDataSource4;
			appearance10.BackColor = System.Drawing.Color.White;
			this.gdHistorial.DisplayLayout.Appearance = appearance10;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance11;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellButtonAppearance = appearance12;
			ultraGridColumn5.Header.Caption = "Cedula";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 90;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.Caption = "Cliente";
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 220;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance13;
			ultraGridColumn7.Header.Caption = "Almacen";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 60;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridColumn8.CellAppearance = appearance14;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Width = 140;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 140;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellAppearance = appearance15;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellButtonAppearance = appearance16;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Width = 50;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance17;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellButtonAppearance = appearance18;
			ultraGridColumn12.Header.Caption = "Cuotas";
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Width = 50;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance19;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellButtonAppearance = appearance20;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Width = 75;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Format = "";
			ultraGridColumn14.Header.Caption = "Fecha Primer Pago";
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Width = 120;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance21;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellButtonAppearance = appearance22;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Valor Cuotas";
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Width = 100;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance23;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellButtonAppearance = appearance24;
			ultraGridColumn16.Header.Caption = "C Pend";
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Width = 70;
			ultraGridColumn17.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellAppearance = appearance25;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellButtonAppearance = appearance26;
			ultraGridColumn17.Header.Caption = "C Pag";
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 70;
			ultraGridColumn18.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance27;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellButtonAppearance = appearance28;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Saldo";
			ultraGridColumn18.Header.VisiblePosition = 15;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 100;
			ultraGridColumn19.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn19.CellAppearance = appearance29;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellButtonAppearance = appearance30;
			ultraGridColumn19.Header.Caption = "Promedio Pago";
			ultraGridColumn19.Header.VisiblePosition = 13;
			ultraGridColumn19.Width = 110;
			ultraGridColumn20.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellButtonAppearance = appearance31;
			ultraGridColumn20.Header.Caption = "Ultimo Pago";
			ultraGridColumn20.Header.VisiblePosition = 14;
			ultraGridColumn20.Width = 100;
			ultraGridColumn21.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn21.CellAppearance = appearance32;
			ultraGridColumn21.Header.VisiblePosition = 16;
			ultraGridColumn21.Width = 120;
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
																										 ultraGridColumn21});
			this.gdHistorial.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance33.ForeColor = System.Drawing.Color.Black;
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.gdHistorial.DisplayLayout.Override.ActiveRowAppearance = appearance33;
			this.gdHistorial.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance34.BackColor = System.Drawing.Color.Transparent;
			this.gdHistorial.DisplayLayout.Override.CardAreaAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance35.FontData.BoldAsString = "True";
			appearance35.FontData.Name = "Arial";
			appearance35.FontData.SizeInPoints = 8F;
			appearance35.ForeColor = System.Drawing.Color.White;
			appearance35.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.gdHistorial.DisplayLayout.Override.HeaderAppearance = appearance35;
			this.gdHistorial.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.gdHistorial.DisplayLayout.Override.RowAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gdHistorial.DisplayLayout.Override.RowSelectorAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gdHistorial.DisplayLayout.Override.SelectedRowAppearance = appearance38;
			this.gdHistorial.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.gdHistorial.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.gdHistorial.Location = new System.Drawing.Point(8, 360);
			this.gdHistorial.Name = "gdHistorial";
			this.gdHistorial.Size = new System.Drawing.Size(882, 112);
			this.gdHistorial.TabIndex = 671;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn14.DataType = typeof(System.DateTime);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.DateTime);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(System.DateTime);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn27});
			// 
			// ultraDataSource2
			// 
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn44});
			// 
			// gdObservaciones
			// 
			this.gdObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gdObservaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.gdObservaciones.DataSource = this.ultraDataSource9;
			appearance39.BackColor = System.Drawing.Color.White;
			this.gdObservaciones.DisplayLayout.Appearance = appearance39;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Width = 160;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 100;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn24.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Width = 100;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn25.Header.Caption = "Almacen";
			ultraGridColumn25.Header.VisiblePosition = 3;
			ultraGridColumn25.Width = 140;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn26.Header.VisiblePosition = 4;
			ultraGridColumn26.Width = 120;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn27.Header.VisiblePosition = 5;
			ultraGridColumn27.Width = 90;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn28.CellAppearance = appearance40;
			ultraGridColumn28.Header.Caption = "Calificación";
			ultraGridColumn28.Header.VisiblePosition = 6;
			ultraGridColumn28.Width = 80;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn29.CellAppearance = appearance41;
			ultraGridColumn29.Header.VisiblePosition = 7;
			ultraGridColumn29.Width = 60;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance42;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 8;
			ultraGridColumn30.Width = 60;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn31.Header.VisiblePosition = 9;
			ultraGridColumn31.Width = 160;
			ultraGridBand3.Columns.AddRange(new object[] {
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
			this.gdObservaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance43.ForeColor = System.Drawing.Color.Black;
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.gdObservaciones.DisplayLayout.Override.ActiveRowAppearance = appearance43;
			this.gdObservaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance44.BackColor = System.Drawing.Color.Transparent;
			this.gdObservaciones.DisplayLayout.Override.CardAreaAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance45.FontData.BoldAsString = "True";
			appearance45.FontData.Name = "Arial";
			appearance45.FontData.SizeInPoints = 8F;
			appearance45.ForeColor = System.Drawing.Color.White;
			appearance45.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.gdObservaciones.DisplayLayout.Override.HeaderAppearance = appearance45;
			this.gdObservaciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.gdObservaciones.DisplayLayout.Override.RowAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gdObservaciones.DisplayLayout.Override.RowSelectorAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gdObservaciones.DisplayLayout.Override.SelectedRowAppearance = appearance48;
			this.gdObservaciones.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.gdObservaciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.gdObservaciones.Location = new System.Drawing.Point(8, 480);
			this.gdObservaciones.Name = "gdObservaciones";
			this.gdObservaciones.Size = new System.Drawing.Size(882, 120);
			this.gdObservaciones.TabIndex = 672;
			// 
			// ultraDataSource9
			// 
			ultraDataColumn47.DataType = typeof(System.DateTime);
			ultraDataColumn52.DataType = typeof(int);
			ultraDataColumn53.DataType = typeof(System.Decimal);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
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
			// ultraDataSource3
			// 
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 605);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 673;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(96, 605);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 674;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtidCre_DatosGenerales
			// 
			this.txtidCre_DatosGenerales.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtidCre_DatosGenerales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCre_DatosGenerales.Enabled = false;
			this.txtidCre_DatosGenerales.Location = new System.Drawing.Point(700, 605);
			this.txtidCre_DatosGenerales.Name = "txtidCre_DatosGenerales";
			this.txtidCre_DatosGenerales.PromptChar = ' ';
			this.txtidCre_DatosGenerales.Size = new System.Drawing.Size(64, 22);
			this.txtidCre_DatosGenerales.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidCre_DatosGenerales.TabIndex = 677;
			this.txtidCre_DatosGenerales.Visible = false;
			// 
			// txtidCre_solicitud
			// 
			this.txtidCre_solicitud.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtidCre_solicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCre_solicitud.Enabled = false;
			this.txtidCre_solicitud.Location = new System.Drawing.Point(784, 605);
			this.txtidCre_solicitud.Name = "txtidCre_solicitud";
			this.txtidCre_solicitud.PromptChar = ' ';
			this.txtidCre_solicitud.Size = new System.Drawing.Size(64, 22);
			this.txtidCre_solicitud.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidCre_solicitud.TabIndex = 678;
			this.txtidCre_solicitud.Visible = false;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(672, 200);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(62, 17);
			this.label17.TabIndex = 679;
			this.label17.Text = "Cuota Valor";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label17.Visible = false;
			// 
			// txtCuotaValor
			// 
			this.txtCuotaValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaValor.FormatString = "#,##0";
			this.txtCuotaValor.Location = new System.Drawing.Point(752, 200);
			this.txtCuotaValor.MaskInput = "nnn";
			this.txtCuotaValor.MaxValue = 999;
			this.txtCuotaValor.MinValue = 0;
			this.txtCuotaValor.Name = "txtCuotaValor";
			this.txtCuotaValor.PromptChar = ' ';
			this.txtCuotaValor.Size = new System.Drawing.Size(100, 22);
			this.txtCuotaValor.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtCuotaValor.TabIndex = 680;
			this.txtCuotaValor.Visible = false;
			this.txtCuotaValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuotaValor_KeyDown);
			this.txtCuotaValor.Click += new System.EventHandler(this.txtCuotaValor_Click);
			this.txtCuotaValor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCuotaValor_KeyPress);
			this.txtCuotaValor.Enter += new System.EventHandler(this.txtCuotaValor_Enter);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource5;
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 60;
			ultraGridColumn33.Header.Caption = "Local";
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridColumn33.Width = 180;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn32,
																										 ultraGridColumn33});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBodega.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(48, 40);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 776;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown_1);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn59.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn59,
																																 ultraDataColumn60});
			// 
			// cmdEntidad
			// 
			this.cmdEntidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmdEntidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdEntidad.DataSource = this.ultraDataSource10;
			ultraGridColumn34.Header.VisiblePosition = 0;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 1;
			ultraGridColumn35.Width = 160;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			this.cmdEntidad.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmdEntidad.DisplayMember = "Nombre";
			this.cmdEntidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdEntidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdEntidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdEntidad.Location = new System.Drawing.Point(112, 200);
			this.cmdEntidad.MaxDropDownItems = 30;
			this.cmdEntidad.Name = "cmdEntidad";
			this.cmdEntidad.Size = new System.Drawing.Size(160, 21);
			this.cmdEntidad.TabIndex = 777;
			this.cmdEntidad.ValueMember = "idEntidadFinanciera";
			this.cmdEntidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmdEntidad_KeyDown);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn61.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn61,
																																	ultraDataColumn62});
			// 
			// cmbEstado
			// 
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.ultraDataSource7;
			ultraGridColumn36.Header.VisiblePosition = 0;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Width = 160;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbEstado.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbEstado.DisplayMember = "Estado";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Enabled = false;
			this.cmbEstado.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(328, 200);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(176, 22);
			this.cmbEstado.TabIndex = 779;
			this.cmbEstado.ValueMember = "idEstado";
			this.cmbEstado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstado_KeyDown_1);
			this.cmbEstado.ValueChanged += new System.EventHandler(this.cmbEstado_ValueChanged_1);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn63.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn63,
																																 ultraDataColumn64});
			// 
			// cmdCalificacion
			// 
			this.cmdCalificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmdCalificacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdCalificacion.DataSource = this.ultraDataSource6;
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.Caption = "Calificacion";
			ultraGridColumn39.Header.VisiblePosition = 1;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn38,
																										 ultraGridColumn39});
			this.cmdCalificacion.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmdCalificacion.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmdCalificacion.DisplayMember = "Codigo";
			this.cmdCalificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdCalificacion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdCalificacion.Enabled = false;
			this.cmdCalificacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdCalificacion.Location = new System.Drawing.Point(80, 240);
			this.cmdCalificacion.Name = "cmdCalificacion";
			this.cmdCalificacion.Size = new System.Drawing.Size(90, 22);
			this.cmdCalificacion.TabIndex = 782;
			this.cmdCalificacion.ValueMember = "idTipoCalificacion";
			this.cmdCalificacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmdCalificacion_KeyDown_1);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn65.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn65,
																																 ultraDataColumn66});
			// 
			// cmbTipoRuc
			// 
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRuc.DisplayMember = "Nombre";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(288, 8);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 22);
			this.cmbTipoRuc.TabIndex = 783;
			this.cmbTipoRuc.ValueMember = "idTipoDoc";
			// 
			// cmbidTipoCliente
			// 
			this.cmbidTipoCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbidTipoCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidTipoCliente.DataSource = this.ultraDataSource8;
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.Caption = "Local";
			ultraGridColumn41.Header.VisiblePosition = 1;
			ultraGridColumn41.Width = 180;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn40,
																										 ultraGridColumn41});
			this.cmbidTipoCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbidTipoCliente.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbidTipoCliente.DisplayMember = "Nombre";
			this.cmbidTipoCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbidTipoCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbidTipoCliente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidTipoCliente.Location = new System.Drawing.Point(288, 40);
			this.cmbidTipoCliente.Name = "cmbidTipoCliente";
			this.cmbidTipoCliente.Size = new System.Drawing.Size(180, 22);
			this.cmbidTipoCliente.TabIndex = 785;
			this.cmbidTipoCliente.ValueMember = "idTipoCliente";
			// 
			// ultraDataSource8
			// 
			ultraDataColumn67.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn67,
																																 ultraDataColumn68});
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(248, 40);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(26, 17);
			this.label9.TabIndex = 784;
			this.label9.Text = "Tipo";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbdetallecliente
			// 
			this.cmbdetallecliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbdetallecliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbdetallecliente.DataSource = this.ultraDataSource11;
			ultraGridColumn42.Header.VisiblePosition = 0;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.Caption = "Local";
			ultraGridColumn43.Header.VisiblePosition = 1;
			ultraGridColumn43.Width = 180;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn42,
																										 ultraGridColumn43});
			this.cmbdetallecliente.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbdetallecliente.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbdetallecliente.DisplayMember = "Nombre";
			this.cmbdetallecliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbdetallecliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbdetallecliente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbdetallecliente.Location = new System.Drawing.Point(552, 40);
			this.cmbdetallecliente.Name = "cmbdetallecliente";
			this.cmbdetallecliente.Size = new System.Drawing.Size(168, 22);
			this.cmbdetallecliente.TabIndex = 787;
			this.cmbdetallecliente.ValueMember = "idDetalleTipoCliente";
			// 
			// ultraDataSource11
			// 
			ultraDataColumn69.DataType = typeof(int);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn69,
																																	ultraDataColumn70});
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(480, 40);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(63, 17);
			this.label11.TabIndex = 786;
			this.label11.Text = "Tipo Cliente";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(8, 136);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(71, 17);
			this.label14.TabIndex = 788;
			this.label14.Text = "Tipo Vivienda";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(200, 136);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(79, 17);
			this.label18.TabIndex = 791;
			this.label18.Text = "Fecha Vivienda";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaVivienda
			// 
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaVivienda.Appearance = appearance49;
			this.dtFechaVivienda.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaVivienda.DateButtons.Add(dateButton2);
			this.dtFechaVivienda.Format = "dd/MM/yyyy";
			this.dtFechaVivienda.Location = new System.Drawing.Point(200, 160);
			this.dtFechaVivienda.Name = "dtFechaVivienda";
			this.dtFechaVivienda.NonAutoSizeHeight = 23;
			this.dtFechaVivienda.Size = new System.Drawing.Size(112, 21);
			this.dtFechaVivienda.SpinButtonsVisible = true;
			this.dtFechaVivienda.TabIndex = 790;
			this.dtFechaVivienda.Value = ((object)(resources.GetObject("dtFechaVivienda.Value")));
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(328, 136);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(68, 17);
			this.label19.TabIndex = 792;
			this.label19.Text = "Tipo Trabajo";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Location = new System.Drawing.Point(520, 136);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(122, 17);
			this.label20.TabIndex = 795;
			this.label20.Text = "Fecha de Ingreso/Inicio";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaIngreso
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaIngreso.Appearance = appearance50;
			this.dtFechaIngreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaIngreso.DateButtons.Add(dateButton3);
			this.dtFechaIngreso.Format = "dd/MM/yyyy";
			this.dtFechaIngreso.Location = new System.Drawing.Point(520, 160);
			this.dtFechaIngreso.Name = "dtFechaIngreso";
			this.dtFechaIngreso.NonAutoSizeHeight = 23;
			this.dtFechaIngreso.Size = new System.Drawing.Size(112, 21);
			this.dtFechaIngreso.SpinButtonsVisible = true;
			this.dtFechaIngreso.TabIndex = 794;
			this.dtFechaIngreso.Value = ((object)(resources.GetObject("dtFechaIngreso.Value")));
			this.dtFechaIngreso.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.ultraCalendarCombo2_BeforeDropDown);
			// 
			// ultraDataSource12
			// 
			ultraDataColumn71.DataType = typeof(int);
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn71,
																																	ultraDataColumn72});
			// 
			// ultraDataSource13
			// 
			ultraDataColumn73.DataType = typeof(int);
			this.ultraDataSource13.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn73,
																																	ultraDataColumn74});
			// 
			// cmbTipoVivienda
			// 
			this.cmbTipoVivienda.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoVivienda.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoVivienda.DataSource = this.ultraDataSource12;
			ultraGridColumn44.Header.VisiblePosition = 0;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 1;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn44,
																											ultraGridColumn45});
			this.cmbTipoVivienda.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbTipoVivienda.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbTipoVivienda.DisplayMember = "Nombre";
			this.cmbTipoVivienda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoVivienda.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoVivienda.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoVivienda.Location = new System.Drawing.Point(8, 160);
			this.cmbTipoVivienda.Name = "cmbTipoVivienda";
			this.cmbTipoVivienda.Size = new System.Drawing.Size(180, 22);
			this.cmbTipoVivienda.TabIndex = 796;
			this.cmbTipoVivienda.ValueMember = "idTipoVivienda";
			// 
			// cmdEstadoSolicitud
			// 
			this.cmdEstadoSolicitud.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmdEstadoSolicitud.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdEstadoSolicitud.DataSource = this.ultraDataSource14;
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Width = 175;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn46,
																											ultraGridColumn47});
			this.cmdEstadoSolicitud.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmdEstadoSolicitud.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmdEstadoSolicitud.DisplayMember = "Nombre";
			this.cmdEstadoSolicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdEstadoSolicitud.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdEstadoSolicitud.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdEstadoSolicitud.Location = new System.Drawing.Point(616, 200);
			this.cmdEstadoSolicitud.Name = "cmdEstadoSolicitud";
			this.cmdEstadoSolicitud.Size = new System.Drawing.Size(176, 22);
			this.cmdEstadoSolicitud.TabIndex = 798;
			this.cmdEstadoSolicitud.ValueMember = "idNegadoPendiente";
			// 
			// ultraDataSource14
			// 
			ultraDataColumn75.DataType = typeof(int);
			this.ultraDataSource14.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn75,
																																	ultraDataColumn76});
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.Location = new System.Drawing.Point(520, 200);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(83, 17);
			this.label21.TabIndex = 797;
			this.label21.Text = "Estado Solicitud";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optDependiente
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.optDependiente.Appearance = appearance51;
			this.optDependiente.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optDependiente.CheckedIndex = 0;
			this.optDependiente.ItemAppearance = appearance52;
			this.optDependiente.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = false;
			valueListItem1.DisplayText = "Dependiente";
			valueListItem2.DataValue = true;
			valueListItem2.DisplayText = "Independiente";
			this.optDependiente.Items.Add(valueListItem1);
			this.optDependiente.Items.Add(valueListItem2);
			this.optDependiente.ItemSpacingVertical = 10;
			this.optDependiente.Location = new System.Drawing.Point(616, 240);
			this.optDependiente.Name = "optDependiente";
			this.optDependiente.Size = new System.Drawing.Size(208, 26);
			this.optDependiente.TabIndex = 1031;
			this.optDependiente.Text = "Dependiente";
			// 
			// txtFuenteIngresos
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFuenteIngresos.Appearance = appearance53;
			this.txtFuenteIngresos.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtFuenteIngresos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFuenteIngresos.Location = new System.Drawing.Point(624, 280);
			this.txtFuenteIngresos.Name = "txtFuenteIngresos";
			this.txtFuenteIngresos.Size = new System.Drawing.Size(224, 22);
			this.txtFuenteIngresos.TabIndex = 1032;
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.Location = new System.Drawing.Point(512, 280);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(101, 17);
			this.label22.TabIndex = 1033;
			this.label22.Text = "Fuente de Ingresos";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.Location = new System.Drawing.Point(512, 322);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(108, 17);
			this.label23.TabIndex = 1035;
			this.label23.Text = "Fecha de Nacimiento";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaNacimiento
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaNacimiento.Appearance = appearance54;
			this.dtFechaNacimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtFechaNacimiento.DateButtons.Add(dateButton4);
			this.dtFechaNacimiento.Format = "dd/MM/yyyy";
			this.dtFechaNacimiento.Location = new System.Drawing.Point(632, 320);
			this.dtFechaNacimiento.Name = "dtFechaNacimiento";
			this.dtFechaNacimiento.NonAutoSizeHeight = 23;
			this.dtFechaNacimiento.Size = new System.Drawing.Size(112, 21);
			this.dtFechaNacimiento.SpinButtonsVisible = true;
			this.dtFechaNacimiento.TabIndex = 1034;
			this.dtFechaNacimiento.Value = ((object)(resources.GetObject("dtFechaNacimiento.Value")));
			// 
			// cmbSituacionLaboral
			// 
			this.cmbSituacionLaboral.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSituacionLaboral.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSituacionLaboral.DataSource = this.ultraDataSource13;
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.Caption = "Local";
			ultraGridColumn49.Header.VisiblePosition = 1;
			ultraGridColumn49.Width = 180;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn48,
																											ultraGridColumn49});
			this.cmbSituacionLaboral.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbSituacionLaboral.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbSituacionLaboral.DisplayMember = "Descripcion";
			this.cmbSituacionLaboral.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSituacionLaboral.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSituacionLaboral.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSituacionLaboral.Location = new System.Drawing.Point(328, 160);
			this.cmbSituacionLaboral.Name = "cmbSituacionLaboral";
			this.cmbSituacionLaboral.Size = new System.Drawing.Size(180, 22);
			this.cmbSituacionLaboral.TabIndex = 1036;
			this.cmbSituacionLaboral.ValueMember = "idSituacionLaboral";
			// 
			// txtCupo
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCupo.Appearance = appearance55;
			this.txtCupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCupo.Enabled = false;
			this.txtCupo.Location = new System.Drawing.Point(400, 240);
			this.txtCupo.MinValue = 0;
			this.txtCupo.Name = "txtCupo";
			this.txtCupo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCupo.PromptChar = ' ';
			this.txtCupo.Size = new System.Drawing.Size(104, 22);
			this.txtCupo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtCupo.TabIndex = 1037;
			// 
			// chkNuovo
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkNuovo.Appearance = appearance56;
			this.chkNuovo.BackColor = System.Drawing.Color.Transparent;
			this.chkNuovo.Enabled = false;
			this.chkNuovo.Location = new System.Drawing.Point(752, 320);
			this.chkNuovo.Name = "chkNuovo";
			this.chkNuovo.Size = new System.Drawing.Size(56, 22);
			this.chkNuovo.TabIndex = 1038;
			this.chkNuovo.Text = "Nuovo";
			// 
			// cmbMarca
			// 
			this.cmbMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMarca.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMarca.DataSource = this.ultraDataSource15;
			ultraGridColumn50.Header.VisiblePosition = 0;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 1;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn50,
																											ultraGridColumn51});
			this.cmbMarca.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbMarca.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbMarca.DisplayMember = "Marca";
			this.cmbMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMarca.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMarca.Location = new System.Drawing.Point(640, 160);
			this.cmbMarca.Name = "cmbMarca";
			this.cmbMarca.Size = new System.Drawing.Size(120, 22);
			this.cmbMarca.TabIndex = 1039;
			this.cmbMarca.ValueMember = "idKnox";
			this.cmbMarca.Visible = false;
			this.cmbMarca.ValueChanged += new System.EventHandler(this.cmbMarca_ValueChanged);
			// 
			// ultraDataSource15
			// 
			ultraDataColumn77.DataType = typeof(int);
			this.ultraDataSource15.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn77,
																																	ultraDataColumn78});
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.Location = new System.Drawing.Point(648, 136);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(86, 17);
			this.label24.TabIndex = 1041;
			this.label24.Text = "Colateral: Marca";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label24.Visible = false;
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.Location = new System.Drawing.Point(776, 136);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(40, 17);
			this.label25.TabIndex = 1042;
			this.label25.Text = "Modelo";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label25.Visible = false;
			// 
			// ultraDataSource16
			// 
			ultraDataColumn79.DataType = typeof(int);
			this.ultraDataSource16.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn79,
																																	ultraDataColumn80});
			// 
			// cmbModelo
			// 
			this.cmbModelo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbModelo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbModelo.DataSource = this.ultraDataSource16;
			ultraGridColumn52.Header.VisiblePosition = 0;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 1;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn52,
																											ultraGridColumn53});
			this.cmbModelo.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbModelo.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbModelo.DisplayMember = "Modelo";
			this.cmbModelo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbModelo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbModelo.Location = new System.Drawing.Point(768, 160);
			this.cmbModelo.Name = "cmbModelo";
			this.cmbModelo.Size = new System.Drawing.Size(112, 22);
			this.cmbModelo.TabIndex = 1044;
			this.cmbModelo.ValueMember = "idKNOXModelo";
			this.cmbModelo.Visible = false;
			this.cmbModelo.ValueChanged += new System.EventHandler(this.cmbModelo_ValueChanged);
			this.cmbModelo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbModelo_InitializeLayout);
			// 
			// ChkOmite
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.ChkOmite.Appearance = appearance57;
			this.ChkOmite.BackColor = System.Drawing.Color.Transparent;
			this.ChkOmite.Location = new System.Drawing.Point(736, 40);
			this.ChkOmite.Name = "ChkOmite";
			this.ChkOmite.Size = new System.Drawing.Size(104, 22);
			this.ChkOmite.TabIndex = 1045;
			this.ChkOmite.Text = "No aplica Nuovo";
			// 
			// frmCred_RevisiondeCedulas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(898, 635);
			this.Controls.Add(this.ChkOmite);
			this.Controls.Add(this.cmbModelo);
			this.Controls.Add(this.label25);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblEntidad);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbMarca);
			this.Controls.Add(this.chkNuovo);
			this.Controls.Add(this.txtCupo);
			this.Controls.Add(this.cmbSituacionLaboral);
			this.Controls.Add(this.dtFechaNacimiento);
			this.Controls.Add(this.txtFuenteIngresos);
			this.Controls.Add(this.optDependiente);
			this.Controls.Add(this.cmdEstadoSolicitud);
			this.Controls.Add(this.cmbTipoVivienda);
			this.Controls.Add(this.dtFechaIngreso);
			this.Controls.Add(this.dtFechaVivienda);
			this.Controls.Add(this.cmbdetallecliente);
			this.Controls.Add(this.cmbidTipoCliente);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.cmdCalificacion);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.cmdEntidad);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtCuotaValor);
			this.Controls.Add(this.txtidCre_solicitud);
			this.Controls.Add(this.txtidCre_DatosGenerales);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.gdObservaciones);
			this.Controls.Add(this.gdHistorial);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.txtCelular);
			this.Controls.Add(this.txtPichicncha);
			this.Controls.Add(this.cmdCalificacion1);
			this.Controls.Add(this.txtScore);
			this.Controls.Add(this.chkWeb);
			this.Controls.Add(this.txtSnombre);
			this.Controls.Add(this.txtPnombre);
			this.Controls.Add(this.txtApellidoM);
			this.Controls.Add(this.txtApellidoP);
			this.Controls.Add(this.btnver);
			this.Controls.Add(this.txtDocumento);
			this.Controls.Add(this.dtFecha);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCred_RevisiondeCedulas";
			this.Text = "Revisión de Cedulas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCred_RevisiondeCedulas_KeyDown);
			this.Load += new System.EventHandler(this.frmRevisiondeCedulas_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoM)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPnombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSnombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtScore)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdCalificacion1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPichicncha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gdHistorial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gdObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCre_DatosGenerales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCre_solicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdEntidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdCalificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidTipoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbdetallecliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdEstadoSolicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optDependiente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFuenteIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSituacionLaboral)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource15)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource16)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbModelo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		//bool bCarga = false;
		private Acceso miAcceso;

		//	bool bVistaReiterativo = false;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmRevisiondeCedulas_Load(object sender, System.EventArgs e)	
		{	
			miAcceso = new Acceso(cdsSeteoF, "083402");

			if (!Funcion.Permiso("982", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Revisión de Cedulas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			bCarga = true;
	

			this.cmbBodega.DataSource = FuncionesProcedimientos.dtGeneral("Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 Order By Nombre");
			this.dtFecha.Value = DateTime.Today;	

			this.dtFechaIngreso.CalendarInfo.MaxDate  = DateTime.Today.AddYears(-1);
			this.dtFechaIngreso.CalendarInfo.MinDate = DateTime.Today.AddYears(-52);

			this.dtFechaNacimiento.CalendarInfo.MaxDate  = DateTime.Today.AddYears(-1);
			this.dtFechaNacimiento.CalendarInfo.MinDate = DateTime.Today.AddYears(-52);

			//this.dtFechaVivienda.CalendarInfo.MaxDate  = DateTime.Today.AddYears(-1);
			//this.dtFechaVivienda.CalendarInfo.MinDate = DateTime.Today.AddYears(-1);
		
			#region Carga combo

			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 4, 0, 0");
			this.cmdEntidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 95, 0, 0");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 7, 0, 0");
			this.cmdCalificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 8, 0, 0");
			this.cmbEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 9, 0, 0");

			this.cmbidTipoCliente.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 78, 0, 0");
			this.cmbdetallecliente.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 79, 0, 0");

			this.cmbTipoVivienda.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 88, 0, 0");
			this.cmbSituacionLaboral.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 87, 0, 0");

			this.cmbMarca.DataSource = Funcion.dvProcedimiento(cdsSeteoF, " EXEC [ConsultaKnoxMarcaModelo] 1, 0 ");
			this.cmbModelo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaKnoxMarcaModelo  {0}, {1}, {2}", 4, 0, 0));
				

			this.cmbTipoRuc.Value = 1;
			
			#endregion Carga combo

			CargaControles(true);

			if (scedula.Length > 0)
			{
				this.txtDocumento.Text = scedula;
				CargarDatos (this.txtDocumento.Text);
				this.btnSalir.Text = "Cancelar";
				this.ControlBox = false;
			}
			else
			{
				this.txtDocumento.Enabled = true;
				this.txtDocumento.Focus ();
				
			}
			bCarga = false;
		}

		

		private void txtDocumento_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{ 
			Funcion.CamposNumericos(sender, e);
		}

		private void txtDocumento_Validated(object sender, System.EventArgs e)
		{	
			CargarDatos(this.txtDocumento.Text.ToString().Trim());	
		}

		private void txtDocumento_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.cmbTipoRuc.ActiveRow != null)
			{
				int iTipoRuc = 0;

				if ((int) this.cmbTipoRuc.Value == 1) iTipoRuc = 5;// cedula
				if ((int) this.cmbTipoRuc.Value == 2) iTipoRuc = 6;// pasaporte
				if ((int) this.cmbTipoRuc.Value == 3) iTipoRuc = 4;// ruc

				if (!Validacion.vbIdentificacion(this.txtDocumento, iTipoRuc, cdsSeteoF)) e.Cancel = true;
			}			
		}

		private void txtDocumento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)		
		{
			if (e.KeyCode == Keys.Enter) this.txtApellidoP.Focus();
		}

		private void btnver_Click(object sender, System.EventArgs e)
		{
			
		}

		private void CargarDatos(string sDocumento)
		{
			CargaControles(true);

			this.Limpia();

			int cmbModelo = 0;
			int iScore = 0;
//			bool bRes = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDatosCognoScoring '{0}'",sDocumento), true);
//			if (bRes)
//			{
//				string sSQLDatosNovaScoring = string.Format("Exec ActualizaDatosNovaScoring '{0}'",sDocumento);
//				Funcion.EjecutaSQL(cdsSeteoF, sSQLDatosNovaScoring);
//				MessageBox.Show("Datos de Sincronizados con CognoWare", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
//			}
			String ssql= string.Format ("Exec Cre_ConsultaRevisionCedula '{0}'",sDocumento); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();			
				
			if (dr.HasRows)

			{		     
				if (dr.GetValue(2) != System.DBNull.Value)  this.txtidCre_solicitud.Value = dr.GetInt32 (2);															  
				if (dr.GetValue(0) != System.DBNull.Value)  this.txtidCre_DatosGenerales.Value = dr.GetInt32 (0);				
				if (dr.GetValue(3) != System.DBNull.Value) this.cmbTipoRuc.Value = dr.GetInt32 (3);			
				if (dr.GetValue(4) != System.DBNull.Value) this.txtDocumento.Text = dr.GetValue(4).ToString();
				if (dr.GetValue(5) != System.DBNull.Value) this.txtApellidoP.Text = dr.GetValue(5).ToString();
				if (dr.GetValue(6) != System.DBNull.Value) this.txtApellidoM.Text = dr.GetValue(6).ToString();
				if (dr.GetValue(7) != System.DBNull.Value) this.txtPnombre.Text = dr.GetValue(7).ToString();
				if (dr.GetValue(8) != System.DBNull.Value) this.txtSnombre.Text = dr.GetValue(8).ToString();
				if (dr.GetValue(9) != System.DBNull.Value) this.txtCelular.Text = dr.GetValue(9).ToString();
				//this.cmbBodega.Value = dr.GetInt32(11);
				//this.cmdEntidad.Value = dr.GetInt32(12);
				//this.cmbEstado.Value = dr.GetInt32(13);
				if (dr.GetValue(14) != System.DBNull.Value) this.cmdEstadoSolicitud.Value = dr.GetInt32(14);
				if (dr.GetValue(11) != System.DBNull.Value) this.cmdCalificacion.Value = dr.GetInt32(11);
				//this.cmdCalificacion1.Value = dr.GetInt32(16);
				if (dr.GetValue(13) != System.DBNull.Value) iScore = dr.GetInt32(13);
				if (dr.GetValue(16) != System.DBNull.Value) this.cmbEstado.Value = dr.GetInt32(16);
				if (dr.GetValue(17) != System.DBNull.Value) cmbModelo = dr.GetInt32(17);	
				if (dr.GetValue(18) != System.DBNull.Value) this.chkNuovo.Checked = dr.GetBoolean(18);	
				if (dr.GetValue(19) != System.DBNull.Value) this.ChkOmite.Checked = dr.GetBoolean(19);												
			}
			dr.Close();
			this.cmbModelo.Value = cmbModelo;

			

			String ssqlsol= string.Format ("Exec Cre_ConsultaSolicitudRevisionCedulas {0}",(int) this.txtidCre_solicitud.Value); 
			SqlDataReader dra = Funcion.rEscalarSQL(cdsSeteoF,ssqlsol,true);
			dra.Read();	

			if (dra.HasRows)
			{


				if (dra.GetValue(0) != System.DBNull.Value) this.cmdEntidad.Value = dra.GetInt32 (0);	
				if (dra.GetValue(1) != System.DBNull.Value) this.cmbBodega.Value = dra.GetInt32(1);
				//if (dra.GetValue(6) != System.DBNull.Value) this.cmbEstado.Value = dra.GetInt32(6);
				if (dra.GetValue(10) != System.DBNull.Value) this.txtObservaciones.Text = dra.GetValue(10).ToString();
				if (dra.GetValue(11) != System.DBNull.Value) this.cmdEstadoSolicitud.Value = dra.GetInt32(11);
				if (dra.GetValue(12) != System.DBNull.Value) this.txtCupo.Value = dra.GetDecimal(12);
				if (dra.GetValue(13) != System.DBNull.Value) this.chkWeb.Checked = dra.GetBoolean(13);
				//this.optBancoProveedor.Value = dr.GetInt32(1);
				if (dra.GetValue(14) != System.DBNull.Value) this.optDependiente.Value = dra.GetBoolean(14);
				if (dra.GetValue(15) != System.DBNull.Value) this.txtFuenteIngresos.Text = dra.GetValue(15).ToString();
				if (dra.GetValue(16) != System.DBNull.Value) 	this.dtFechaIngreso.Value = dra.GetDateTime(16);
				if (dra.GetValue(17) != System.DBNull.Value) this.cmbidTipoCliente.Value = dra.GetInt32 (17);	
				if (dra.GetValue(18) != System.DBNull.Value) this.cmbdetallecliente.Value = dra.GetInt32 (18);	
				if (dra.GetValue(19) != System.DBNull.Value) this.cmbTipoVivienda.Value = dra.GetInt32 (19);	
				if (dra.GetValue(20) != System.DBNull.Value) this.dtFechaVivienda.Value = dra.GetDateTime(20);

				if (dra.GetValue(21) != System.DBNull.Value) this.cmbSituacionLaboral.Value = dra.GetInt32 (21);	
				if (dra.GetValue(22) != System.DBNull.Value)  this.dtFechaNacimiento.Value = dra.GetDateTime(22);


				 

				
			}
			dra.Close();
				

			this.txtScore.Value = iScore;


			if ((int)this.txtidCre_solicitud.Value > 0)
			{			
				this.gdHistorial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaHistorialCreditos {0}", this.txtDocumento.Text));
				this.gdObservaciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaObservacionesSolicitud '{0}'",txtidCre_solicitud.Value));
			
				if ((int) this.cmdEntidad.Value == 3)
				{					
					this.cmdEntidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera Where idEntidadFinanciera not in (6) Order By idEntidadFinanciera");
					this.cmdEntidad.Value = 3;
					CargaControles(false);
					this.btnGuardar.Enabled = false;
				}

				if (!Validacion.vbValidaFacturasCliente(this.txtDocumento.Text, cdsSeteoF)) 
				{
					this.txtPnombre.Enabled = false;
					this.txtSnombre.Enabled = false;
					this.txtApellidoP.Enabled = false;
					this.txtApellidoM.Enabled = false;

					this.txtCelular.Focus();
				}
				else this.cmbBodega.Focus();
			}

		
		}
		
		private void CargaControles(bool bEstado)
		{
			this.txtDocumento.Enabled = bEstado;
			this.txtApellidoP.Enabled = bEstado;	
			this.txtApellidoM.Enabled = bEstado;	
			this.txtPnombre.Enabled = bEstado;	
			this.txtSnombre.Enabled = bEstado;	
			this.txtCelular.Enabled = bEstado;	
			this.cmbBodega.Enabled = bEstado;	
			this.cmdEntidad.Enabled = bEstado;	
			this.cmbEstado.Enabled = bEstado;	
			this.cmdCalificacion.Enabled = bEstado;	
			this.cmdCalificacion1.Enabled = bEstado;	
			this.txtScore.Enabled = bEstado;	
			this.txtPichicncha.Enabled = bEstado;	
			this.txtCupo.Enabled = bEstado;	
			this.chkNuovo.Enabled = bEstado;
			this.cmdEstadoSolicitud.Enabled = bEstado;	
			this.txtObservaciones.Enabled = bEstado;		
			this.chkWeb.Enabled = bEstado;
			this.txtCuotaValor.Enabled = bEstado;
			this.btnGuardar.Enabled = bEstado;
			if (!miAcceso.Editar) 
			{
				bEstado = false;
				this.cmdEntidad.Enabled = bEstado;
				this.cmbEstado.Enabled = bEstado;
				this.cmdEstadoSolicitud.Enabled = bEstado;
				this.cmdCalificacion.Enabled = bEstado;
				this.txtScore.Enabled = bEstado;
				this.txtCupo.Enabled = bEstado;
				this.chkNuovo.Enabled = bEstado;
			}
			if (!miAcceso.BExportar) 
			{
				bEstado = false;
				this.ChkOmite.Enabled = false;
			}
			if (!miAcceso.BAsignaCobrador) 
			{
				bEstado = false;
				this.cmbidTipoCliente.Enabled = false;
			}

		}

		private void Limpia()
		{		
			#region Controles
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbEstado.Value = System.DBNull.Value;
			this.cmdCalificacion.Value = System.DBNull.Value;
			this.cmdCalificacion1.Value = System.DBNull.Value;
			this.cmdEntidad.Value = System.DBNull.Value;
			
			this.txtDocumento.Text = "";
			this.txtPnombre.Text = "";
			this.txtSnombre.Text = "";
			this.txtApellidoP.Text = "";
			this.txtApellidoM.Text = "";
			this.txtCelular.Text = "";
			this.txtScore.Value = 0;
			this.txtPichicncha.Value = 0;
			this.txtCupo.Value = 0;
			this.cmdEstadoSolicitud.Value = 0;	
			this.txtObservaciones.Text = "";
			this.txtCuotaValor.Value = 0;
			this.txtidCre_DatosGenerales.Value =0;
			this.txtidCre_solicitud.Value =0;

			this.gdHistorial.DataSource = ultraDataSource4;
			this.gdObservaciones.DataSource = ultraDataSource9;
			#endregion Controles
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion		
	
		


			if (!Validacion.vbIdentificacion(this.txtDocumento, 5, cdsSeteoF)) return;

			if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione elTipo  de Documento")) return;

			if (!Validacion.vbTexto(this.txtApellidoP, 2, 50, "Apellido Paterno")) return;

			//if (!Validacion.vbTexto(this.txtApellidoM, 3, 50, "Ingrese Apellido Materno")) return;

			if (!Validacion.vbTexto(this.txtPnombre, 2, 50, " Primer Nombre")) return;

			//if (!Validacion.vbTexto(this.txtSnombre, 3, 15, "Ingrese Segundo Nombre")) return;			
			if (miAcceso.Editar) 
			{
				if (!Validacion.vbComboVacio(this.cmdEntidad, "Seleccione una Entidad Financiera")) return;
				if (!(bool)this.chkNuovo.Checked)
				{
					if (!Validacion.vbComboVacio(this.cmbEstado, "Seleccione un Estado")) return;
				}
			
				if (!Validacion.vbComboVacio(this.cmdEstadoSolicitud, "Seleccione el Estado de la Solicitud")) return;
				if (!Validacion.vbComboVacio(this.cmdCalificacion, "Seleccione una Calificacion")) return;
				if (!Validacion.vbCampoEnteroVacio(this.txtScore,"Ingrese el Score",1,999))return;
				if (!Validacion.vbCampoDecimalVacio(this.txtCupo,"Ingrese el Cupo",200,2000))return;

			}
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un Local")) return;			

			if (!Validacion.vbTexto(this.txtDocumento, 5, 10, "Ingrese la Cedula")) return;		

			if (!Validacion.vbValidaTelefono(this.txtCelular, 2, false)) return;

			if (!Validacion.vbTexto(this.txtObservaciones,20,1000,"Ingrese una Observacion")) return;

      if (!Validacion.vbTexto(this.txtFuenteIngresos,10,100,"Ingrese la Fuente de Ingreso")) return;
		
		//	if (!Validacion.vbCampoDecimalVacio(this.txtCuotaValor,"Ingrese el valor de la Cuota",30,300))return;
			if (!Validacion.vbComboVacio(this.cmbidTipoCliente, "Seleccione el Tipo ")) return;
			if (!Validacion.vbComboVacio(this.cmbdetallecliente, "Seleccione el Tipo de Cliente")) return;
			if (!Validacion.vbComboVacio(this.cmbTipoVivienda, "Seleccione el Tipo de Vivienda")) return;
			if (!Validacion.vbComboVacio(this.cmbSituacionLaboral, "Seleccione el Tipo de Trabajo")) return;
//			if (!miAcceso.Editar) 
//			{
//				if ((bool)this.chkNuovo.Checked)
//				{
//					if (!Validacion.vbComboVacio(this.cmbModelo, "Seleccione el Modelo del colateral")) return;
//				
//				}
//			}
			int cmdEntidad = 0;
			int cmbEstado = 0;
			int cmdEstadoSolicitud = 0;
			int cmdCalificacion = 0;
			int iModelo = 0;

			if ( this.cmdEntidad.ActiveRow != null) cmdEntidad  = ( int) this.cmdEntidad.Value;
			if ( this.cmbEstado.ActiveRow != null) cmbEstado  = ( int) this.cmbEstado.Value;
			if ( this.cmdEstadoSolicitud.ActiveRow != null) cmdEstadoSolicitud  = ( int) this.cmdEstadoSolicitud.Value;
			if ( this.cmdCalificacion.ActiveRow != null) cmdCalificacion  = ( int) this.cmdCalificacion.Value;
			if ( this.cmbModelo.ActiveRow != null) iModelo  = ( int) this.cmbModelo.Value;

			#endregion Validacion

			#region Valida Fecha Ingreso - Inicio de la fuente de ingresos
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaIngreso, "Seleccione la Fecha De Ingreso", false, "", cdsSeteoF)) return;

			if (DateTime.Today.AddYears(-52) > Convert.ToDateTime(this.dtFechaIngreso.Value))
			{
				MessageBox.Show("La Fecha de ingreso o inicio de la fuente de ingresos no es correcta", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);                    
				this.dtFechaIngreso.Focus();
				return;
			}

			#endregion Valida Fecha Ingreso - Inicio de la fuente de ingresos
			#region Fecha Vivienda
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaVivienda, "Seleccione la Fecha De vivienda", false, "", cdsSeteoF)) return;
			#endregion Fecha Vivienda

			#region Fecha Nacimiento
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaNacimiento, "Seleccione la Fecha De Nacimiento", false, "", cdsSeteoF)) return;
			
			#endregion Fecha Vivienda

			#region Valida celular cliente web
			if ((bool)this.chkWeb.Checked)
			{
				if (!Validacion.vbValidaTelefono(this.txtCelular, 2, true)) return;
				
			}
			#endregion Valida celular cliente web


			#region Guarda
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

						
					

					string sSQLDetalle = string.Format("Exec Cre_GuardaSolRevisionCedulas {0}, '{1}', {2}, {3}, {4}, '{5}', '{6}', '{7}', '{8}', '{9}', {10}, {11},{12}, {13}, '{14}', '{15}', {16}, {17}, {18}, {19}, '{20}', '{21}'",
						(int)this.txtidCre_solicitud.Value, 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
						cmdEntidad, 
						(int)this.cmbBodega.Value, 0," "," "," ",
						this.txtObservaciones.Text.ToString(), " ",cmdEstadoSolicitud,
						Convert.ToDecimal(this.txtCupo.Value), (bool)this.chkWeb.Checked, 
						
						(bool) this.optDependiente.Value,
						this.txtFuenteIngresos.Text.ToString(),
						Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd"),
						(int)this.cmbidTipoCliente.Value,
						(int)this.cmbdetallecliente.Value,
						(int)this.cmbTipoVivienda.Value,
						(int)this.cmbSituacionLaboral.Value,
						Convert.ToDateTime(this.dtFechaNacimiento.Value).ToString("yyyyMMdd"), 
						Convert.ToDateTime(this.dtFechaVivienda.Value).ToString("yyyyMMdd")
						
						); 
					oCmdActualiza.CommandText = sSQLDetalle;
					this.txtidCre_solicitud.Value = (int)oCmdActualiza.ExecuteScalar();	
										
					string sSQLDGenerales = string.Format("Exec Cre_GuardaDGRevisionCedulas  {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}",
						(int)this.txtidCre_DatosGenerales.Value, 0, (int)this.txtidCre_solicitud.Value, (int)this.cmbTipoRuc.Value, 
						this.txtDocumento.Text.ToString().Trim(), this.txtApellidoP.Text.ToString().Trim(), this.txtApellidoM.Text.ToString().Trim(), this.txtPnombre.Text.ToString().Trim(), this.txtSnombre.Text.ToString().Trim(), "", "",
						cmdCalificacion, 0, (int)this.txtScore.Value, 0, cmbEstado, iModelo,Convert.ToInt32((bool) this.chkNuovo.Checked), Convert.ToInt32((bool) this.ChkOmite.Checked));
					oCmdActualiza.CommandText = sSQLDGenerales;
					this.txtidCre_DatosGenerales.Value =(int)oCmdActualiza.ExecuteScalar();



					string  sSQLDet = string.Format("Exec Cre_GuardarObservaciones {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}, {11}" ,
						(int) this.txtidCre_solicitud.Value, 
						cmdEntidad, 
						cmdCalificacion,
						0,
						(int) this.txtScore.Value,
						0,
						cmbEstado,
						(int) this.cmbBodega.Value,
						this.txtObservaciones.Text,
						0,
					//	0,
						 Convert.ToDecimal (this.txtCupo.Value),
						(bool)this.chkWeb.Checked);
					//	(bool)this.chkWeb.Checked );
					oCmdActualiza.CommandText = sSQLDet;
					oCmdActualiza.ExecuteNonQuery();	

					#region Lista Clientes de Gestion - Web
					if (bVistaReiterativo || this.chkWeb.Checked)
					{
						

						string  sSQLDetwe = string.Format("Exec Cre_ListaClientesGestionActualiza {0}, {1}, {2}, '{3}', '{4}', '{5}', '{6}', {7}, {8}, {9}" ,
							(int) this.txtidCre_solicitud.Value, 
							(int) this.cmbBodega.Value,
							cmdEntidad, 
							this.txtDocumento.Text.ToString(),
							this.txtApellidoP.Text.ToString() + ' ' + this.txtApellidoM.Text.ToString() + ' ' + this.txtPnombre.Text.ToString() + ' ' + this.txtSnombre.Text.ToString(),
              "2111111", this.txtCelular.Text.ToString(),
              cmbEstado,
              	Convert.ToDecimal(this.txtCupo.Value),
							cmdCalificacion,
								(bool)this.chkWeb.Checked  );

						oCmdActualiza.CommandText = sSQLDetwe;
						oCmdActualiza.ExecuteNonQuery();	



					}
					#endregion Lista Clientes de Gestion - Web


					string  sSQLDetwecopia = string.Format("Exec Cre_CopiaCliente '{0}'" ,
						this.txtDocumento.Text.ToString()  );

					oCmdActualiza.CommandText = sSQLDetwecopia;
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

					if (scedula.Length > 0)
					{
						this.DialogResult = DialogResult.OK;
					}
					MessageBox.Show("Registro Guardado Correctamente ", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information); 
					

					this.gdHistorial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaHistorialCreditos {0}", this.txtDocumento.Text));
					this.gdObservaciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaObservacionesSolicitud '{0}'",txtidCre_solicitud.Value));
			

				}
			}
			#endregion Guarda			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
		
		private void txtDocumento_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.txtDocumento.Text.Length == 0) Limpia();
		}

		private void txtApellidoP_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtApellidoM.Focus();
		}

		private void txtApellidoM_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtPnombre.Focus();
		}

		private void txtPnombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtSnombre.Focus();
		}

		private void txtSnombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCelular.Focus();
		}

		private void cmdEntidadFinanciera_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtApellidoP.Focus();
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
	
			if (this.cmbBodega.ActiveRow != null)
			{			
				int ibodega = (int)this.cmbBodega.Value;        				
				this.cmdCalificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"Exec Listas 8, 0, 0");
				this.cmdCalificacion.DisplayMember = "Codigo";
				this.cmdCalificacion.ValueMember = "idTipoCalificacion";				
			}
		}

		private void cmdCalificacion_ValueChanged(object sender, System.EventArgs e)
		{
			

			if (this.cmdCalificacion.ActiveRow != null)
			{
				
				int  iCalificacion = (int)this.cmdCalificacion.Value;		
				
				this.cmdCalificacion1.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"Exec Listas 8, 0, 0");
				this.cmdCalificacion1.DisplayMember = "Codigo";
				this.cmdCalificacion1.ValueMember = "idTipoCalificacion";
			}
		
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmdCalificacion.Focus();
		}

		private void cmdCalificacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtScore.Focus();
		}



		private void txtScore_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
		}

		private void txtCupo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
    
		}

		private void cmbEstado_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void cmbEstado_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmdCalificacion.ActiveRow != null)
			{
				
				int  iCalificacion = (int)this.cmdCalificacion.Value;		
				
				this.cmdEstadoSolicitud.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"Exec Listas 24, 0, 0");
				this.cmdEstadoSolicitud.DisplayMember = "Nombre";
				this.cmdEstadoSolicitud.ValueMember = "idNegadoPendiente";
			}
		}

		private void txtScore_Click(object sender, System.EventArgs e)
		{
			this.txtScore.SelectAll();
		}

		private void txtScore_Enter(object sender, System.EventArgs e)
		{
			this.txtScore.SelectAll();
		}

		private void txtCupo_Click(object sender, System.EventArgs e)
		{
			this.txtCupo.SelectAll();
		}

		private void txtCupo_Enter(object sender, System.EventArgs e)
		{
			this.txtCupo.SelectAll();
		}

		private void txtCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void cmbBodega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void cmdEntidadFinanciera_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbEstado.Focus();
		}

		private void cmbEstado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
	
		}

		private void txtScore_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCupo.Focus();
		}

		private void txtCupo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCuotaValor.Focus();
		}

		private void txtCuotaValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkWeb.Focus();
		}

		private void txtObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkWeb.Focus();
		}

		private void cmdCalificacion1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		
		}

		private void txtCelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmdEntidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)this.cmbEstado.Focus();
		}

		private void cmbBodega_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)this.cmdEntidad.Focus();
		}

		private void cmbEstado_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmdCalificacion.Focus();
		}

		private void cmdCalificacion_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtScore.Focus();
		}

		private void chkWeb_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtObservaciones.Focus();
		}

		private void txtCuotaValor_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{			
			if (scedula.Length > 0) 
				this.DialogResult = DialogResult.OK;
			else this.Close();
		}

		private void txtCuotaValor_Click(object sender, System.EventArgs e)
		{
			this.txtCuotaValor.SelectAll();
		}

		private void txtCuotaValor_Enter(object sender, System.EventArgs e)
		{
			this.txtCuotaValor.SelectAll();
		}

		private void frmCred_RevisiondeCedulas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape)
			{
				if (scedula.Length > 0) 
					this.DialogResult = DialogResult.OK;
				else this.Close();		
			}

			if (e.KeyCode == Keys.F12) 
				this.btnGuardar_Click(sender, e);
		}

		private void ultraCalendarCombo2_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbEstado_ValueChanged_1(object sender, System.EventArgs e)
		{
			if (this.cmbEstado.ActiveRow != null)
			{
				this.cmdEstadoSolicitud.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 24, {0}, 0", (int) this.cmbEstado.Value));
				
			}
		}

		private void cmbMarca_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbMarca.ActiveRow != null)
			{
				this.cmbModelo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaKnoxMarcaModelo 2, {0}", (int) this.cmbMarca.Value));
				
			}
		}

		private void cmbModelo_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbModelo.ActiveRow != null)
			{
				if ((int) this.cmbModelo.Value > 0 )
				{
					int idMarca = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec ConsultaKnoxMarcaModelo  {0}, {1}, {2}",5, 0, (int) this.cmbModelo.Value));		

					this.cmbMarca.Value = idMarca;	
				}
			}
		}

		private void cmbModelo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			System.Diagnostics.Process.Start("chrome.exe", "https://www.samsungknox.com/en/knox-platform/supported-devices");
		}

		private void linkLabel2_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			System.Diagnostics.Process.Start("chrome.exe", "https://storage.googleapis.com/sparta_bucket/Principal/DESAPOINT/0916878820001_000001213_1279118/QR_NUOVO.PNG");
		
		}	
	} 
}



