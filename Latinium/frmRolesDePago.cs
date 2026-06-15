using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;

using Infragistics.Win.UltraWinToolbars;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRolesDePago.
	/// </summary>
	public class frmRolesDePago : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdRolNomina;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridRoles;
		private System.Windows.Forms.Button btnGenerar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpPeriodo;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnContabilizar;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		public System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnAcreditacion;
		private System.Windows.Forms.Button btnAutorizar;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.Label lblDiscapacitados;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEDepartamento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECentroCosto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.ContextMenu mnuSubeArchivos;
		private System.Windows.Forms.MenuItem mnuSAPQ;
		private System.Windows.Forms.MenuItem mnuSAPH;
		private System.Windows.Forms.MenuItem mnuSAF;
		private System.Windows.Forms.Button btnSubirArchivos;
		private System.Windows.Forms.Button btnSeleccionar;
		private System.Windows.Forms.ContextMenu mnuAcreditacion;
		private System.Windows.Forms.MenuItem mnuAcreditacionTransferencia;
		private System.Windows.Forms.MenuItem mnuAcreditacionCheques;
		private System.Windows.Forms.MenuItem mnuAcreditacionRJ;
		private System.Windows.Forms.Button btnRevisado;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesAgregar;
		private System.Windows.Forms.MenuItem mnuAccionesActualizar;
		private System.Windows.Forms.NotifyIcon notifyIcon1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAntipoUtilidades;
		private System.Windows.Forms.Button btnUtilidades;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.MenuItem menuItem4;
		private System.Windows.Forms.MenuItem menuItem5;
		private System.Windows.Forms.MenuItem menuItem6;
		private System.Windows.Forms.MenuItem menuItem7;
		private System.ComponentModel.IContainer components;

		public frmRolesDePago()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Revisado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleRolNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellidos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Salida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasTrabajados");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mujer");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Discapacitado");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDePago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEAsumeEmpleado");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEAsumeEmpresa");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEGastoSueldo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BonoMaster");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bonos");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AnticipoUtilidades");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BonoAdicional");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BonoBeneficios");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comisiones");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alimentacion");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movilizacion");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Spiff");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VHOrdinaria");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NHOrdinarias");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HOrdinarias");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VHExtraOrdinaria");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NHExtraOrdinarias");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HExtraOrdinarias");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionReal");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AnticipoUtilidadesComisiones");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Refrigerios");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DerechoFondosReserva");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FondosReservaAcumula");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FRAcumula");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FRNoAcumula");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalIngresos");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseIESS");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecAcumula");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldo");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldo");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vacaciones");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AportePatronal");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoIESS");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Extension");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrestamoQuirografario");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrestamoHipotecario");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PagoIESS");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComprasCredito");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraCP");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComprasCrediPoint");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FaltantesCaja");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Multas");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Uniformes");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AnticipoComisiones");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionesPagadas");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BonosPagados");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AnticipoSueldos");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SeguroMedico");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Farmacia");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Enfermedad");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrosDescuentos");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Faltas");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pasantes");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RetencionJudicial");
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalEgresos");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesFC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesMultas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesAnticipos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesOtros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SSN");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoTercerSueldoAcumulado");
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DecimoCuartoSueldoAcumulado");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CajaChica");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AsistenciaIngreso");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AsistenciaEgreso");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMultas", 0);
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRolesDePago));
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnFC", 1);
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnOtros", 2);
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnticipos", 3);
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCredito", 4);
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comision", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comision", 11, true);
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Hombre", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Hombre", 12, true);
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Mujer", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Mujer", 13, true);
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Discapacitado", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Discapacitado", 14, true);
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Sueldo", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Sueldo", 21, true);
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bonos", 27, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bonos", 27, true);
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comisiones", 31, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comisiones", 31, true);
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Alimentacion", 32, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Alimentacion", 32, true);
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Movilizacion", 33, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Movilizacion", 33, true);
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Spiff", 34, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Spiff", 34, true);
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NHOrdinarias", 36, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NHOrdinarias", 36, true);
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "HOrdinarias", 37, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "HOrdinarias", 37, true);
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NHExtraOrdinarias", 39, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NHExtraOrdinarias", 39, true);
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "HExtraOrdinarias", 40, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "HExtraOrdinarias", 40, true);
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ComisionReal", 41, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ComisionReal", 41, true);
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FondosReservaAcumula", 45, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FondosReservaAcumula", 45, true);
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FRAcumula", 46, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FRAcumula", 46, true);
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FRNoAcumula", 47, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FRNoAcumula", 47, true);
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DecimoTercerSueldo", 51, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DecimoTercerSueldo", 51, true);
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DecimoCuartoSueldo", 52, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DecimoCuartoSueldo", 52, true);
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalIngresos", 48, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalIngresos", 48, true);
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BaseIESS", 49, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BaseIESS", 49, true);
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DescuentoIESS", 55, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DescuentoIESS", 55, true);
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ComprasCredito", 61, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ComprasCredito", 61, true);
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ComprasCrediPoint", 64, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ComprasCrediPoint", 64, true);
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FaltantesCaja", 65, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FaltantesCaja", 65, true);
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings27 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Multas", 66, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Multas", 66, true);
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings28 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AnticipoComisiones", 68, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AnticipoComisiones", 68, true);
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings29 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AnticipoSueldos", 71, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AnticipoSueldos", 71, true);
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings30 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SeguroMedico", 72, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SeguroMedico", 72, true);
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings31 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Farmacia", 73, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Farmacia", 73, true);
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings32 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PrestamoQuirografario", 57, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PrestamoQuirografario", 57, true);
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings33 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PrestamoHipotecario", 58, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PrestamoHipotecario", 58, true);
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings34 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OtrosDescuentos", 76, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OtrosDescuentos", 76, true);
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings35 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalEgresos", 80, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalEgresos", 80, true);
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings36 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 81, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 81, true);
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings37 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DecAcumula", 50, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DecAcumula", 50, true);
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings38 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "RetencionJudicial", 79, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "RetencionJudicial", 79, true);
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings39 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BonoAdicional", 29, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BonoAdicional", 29, true);
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings40 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DerechoFondosReserva", 44, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DerechoFondosReserva", 44, true);
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings41 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DecimoTercerSueldo", 51, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DecimoTercerSueldo", 51, true);
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings42 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DecimoCuartoSueldo", 52, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DecimoCuartoSueldo", 52, true);
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings43 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Vacaciones", 53, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Vacaciones", 53, true);
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings44 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AportePatronal", 54, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AportePatronal", 54, true);
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings45 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PagoIESS", 59, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PagoIESS", 59, true);
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings46 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Faltas", 77, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Faltas", 77, true);
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings47 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Extension", 56, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Extension", 56, true);
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings48 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Uniformes", 67, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Uniformes", 67, true);
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings49 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ComisionesPagadas", 69, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ComisionesPagadas", 69, true);
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings50 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Refrigerios", 43, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Refrigerios", 43, true);
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings51 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Enfermedad", 74, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Enfermedad", 74, true);
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings52 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Depreciacion", 75, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Depreciacion", 75, true);
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings53 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BonoBeneficios", 30, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BonoBeneficios", 30, true);
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings54 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BonosPagados", 70, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BonosPagados", 70, true);
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings55 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Alimentacion", 32, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Alimentacion", 32, true);
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings56 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Pasantes", 78, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Pasantes", 78, true);
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings57 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AnticipoUtilidades", 28, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AnticipoUtilidades", 28, true);
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings58 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AnticipoUtilidadesComisiones", 42, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AnticipoUtilidadesComisiones", 42, true);
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings59 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 25, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 25, true);
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings60 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AnticipoUtilidades", 28, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AnticipoUtilidades", 28, true);
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings61 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEAsumeEmpleado", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEAsumeEmpleado", 22, true);
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings62 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEAsumeEmpresa", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEAsumeEmpresa", 23, true);
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings63 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEGastoSueldo", 24, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEGastoSueldo", 24, true);
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings64 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BonoMaster", 26, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BonoMaster", 26, true);
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings65 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CajaChica", 91, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CajaChica", 91, true);
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings66 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AsistenciaIngreso", 92, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AsistenciaIngreso", 92, true);
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings67 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AsistenciaEgreso", 93, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AsistenciaEgreso", 93, true);
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Revisado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleRolNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Apellidos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Salida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasTrabajados");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mujer");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Discapacitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDePago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEAsumeEmpleado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEAsumeEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEGastoSueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BonoMaster");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bonos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AnticipoUtilidades");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BonoAdicional");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BonoBeneficios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comisiones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Alimentacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Movilizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Spiff");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VHOrdinaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NHOrdinarias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HOrdinarias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VHExtraOrdinaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NHExtraOrdinarias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HExtraOrdinarias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionReal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AnticipoUtilidadesComisiones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Refrigerios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DerechoFondosReserva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FondosReservaAcumula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FRAcumula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FRNoAcumula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalIngresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BaseIESS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecAcumula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vacaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AportePatronal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoIESS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Extension");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrestamoQuirografario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrestamoHipotecario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PagoIESS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComprasCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompraCP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComprasCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FaltantesCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Multas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Uniformes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AnticipoComisiones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionesPagadas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BonosPagados");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AnticipoSueldos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SeguroMedico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Farmacia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Enfermedad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Depreciacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OtrosDescuentos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Faltas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pasantes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RetencionJudicial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalEgresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesFC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesMultas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesAnticipos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesOtros");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SSN");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoTercerSueldoAcumulado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DecimoCuartoSueldoAcumulado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CajaChica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AsistenciaIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AsistenciaEgreso");
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn98 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.uGridRoles = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtpPeriodo = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.btnContabilizar = new System.Windows.Forms.Button();
			this.txtIdRolNomina = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnAcreditacion = new System.Windows.Forms.Button();
			this.btnAutorizar = new System.Windows.Forms.Button();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.lblDiscapacitados = new System.Windows.Forms.Label();
			this.cmbNEDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label33 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.cmbNECentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnSubirArchivos = new System.Windows.Forms.Button();
			this.mnuSubeArchivos = new System.Windows.Forms.ContextMenu();
			this.mnuSAPQ = new System.Windows.Forms.MenuItem();
			this.mnuSAPH = new System.Windows.Forms.MenuItem();
			this.mnuSAF = new System.Windows.Forms.MenuItem();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this.menuItem4 = new System.Windows.Forms.MenuItem();
			this.menuItem5 = new System.Windows.Forms.MenuItem();
			this.menuItem6 = new System.Windows.Forms.MenuItem();
			this.menuItem7 = new System.Windows.Forms.MenuItem();
			this.btnSeleccionar = new System.Windows.Forms.Button();
			this.mnuAcreditacion = new System.Windows.Forms.ContextMenu();
			this.mnuAcreditacionTransferencia = new System.Windows.Forms.MenuItem();
			this.mnuAcreditacionCheques = new System.Windows.Forms.MenuItem();
			this.mnuAcreditacionRJ = new System.Windows.Forms.MenuItem();
			this.btnRevisado = new System.Windows.Forms.Button();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesAgregar = new System.Windows.Forms.MenuItem();
			this.mnuAccionesActualizar = new System.Windows.Forms.MenuItem();
			this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
			this.txtIdAntipoUtilidades = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnUtilidades = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridRoles)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpPeriodo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolNomina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAntipoUtilidades)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridRoles
			// 
			this.uGridRoles.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridRoles.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridRoles.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridRoles.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.Caption = "R";
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 50;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 150;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 150;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 70;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 70;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellAppearance = appearance3;
			ultraGridColumn11.Header.Caption = "Dias";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance4;
			ultraGridColumn12.Header.Caption = "C";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 40;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance5;
			ultraGridColumn13.Header.Caption = "H";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 40;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance6;
			ultraGridColumn14.Header.Caption = "M";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 40;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance7;
			ultraGridColumn15.Header.Caption = "D";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 40;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.Caption = "Pago";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 58;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.VisiblePosition = 16;
			ultraGridColumn18.Width = 120;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 17;
			ultraGridColumn19.Width = 80;
			ultraGridColumn20.Header.Caption = "Centro De Costo";
			ultraGridColumn20.Header.VisiblePosition = 18;
			ultraGridColumn20.Width = 200;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Header.VisiblePosition = 19;
			ultraGridColumn21.Width = 150;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance8;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 20;
			ultraGridColumn22.Width = 80;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance9;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "S E Asume Empleado";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 120;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance10;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "S E Asume Empresa";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 120;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance11;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.Caption = "Gasto Sueldo";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 120;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance12;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 80;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance13;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.Caption = "Bono Master";
			ultraGridColumn27.Header.VisiblePosition = 27;
			ultraGridColumn27.Width = 90;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance14;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 26;
			ultraGridColumn28.PromptChar = ' ';
			ultraGridColumn28.Width = 80;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance15;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.Caption = "Ant Utilidades";
			ultraGridColumn29.Header.VisiblePosition = 39;
			ultraGridColumn29.Width = 100;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance16;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "Adicional";
			ultraGridColumn30.Header.VisiblePosition = 28;
			ultraGridColumn30.Width = 80;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance17;
			ultraGridColumn31.Format = "#,##0.00";
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.FormulaErrorAppearance = appearance18;
			ultraGridColumn31.Header.Caption = "Beneficios";
			ultraGridColumn31.Header.VisiblePosition = 29;
			ultraGridColumn31.Width = 80;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance19;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.PromptChar = ' ';
			ultraGridColumn32.Width = 80;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance20;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.PromptChar = ' ';
			ultraGridColumn33.Width = 90;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance21;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.PromptChar = ' ';
			ultraGridColumn34.Width = 80;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance22;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.PromptChar = ' ';
			ultraGridColumn35.Width = 80;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance23;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.Caption = "V H Ord";
			ultraGridColumn36.Header.VisiblePosition = 35;
			ultraGridColumn36.Width = 80;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance24;
			ultraGridColumn37.Header.Caption = "N H O";
			ultraGridColumn37.Header.VisiblePosition = 36;
			ultraGridColumn37.Width = 50;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance25;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "H Ordinarias";
			ultraGridColumn38.Header.VisiblePosition = 37;
			ultraGridColumn38.Width = 80;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance26;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.Caption = "V H ExtraOrd";
			ultraGridColumn39.Header.VisiblePosition = 38;
			ultraGridColumn39.Width = 80;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance27;
			ultraGridColumn40.Header.Caption = "N H E";
			ultraGridColumn40.Header.VisiblePosition = 40;
			ultraGridColumn40.Width = 50;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance28;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.Caption = "H ExtraOrdinarias";
			ultraGridColumn41.Header.VisiblePosition = 41;
			ultraGridColumn41.Width = 80;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance29;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.Caption = "Com. Real";
			ultraGridColumn42.Header.VisiblePosition = 42;
			ultraGridColumn42.PromptChar = ' ';
			ultraGridColumn42.Width = 80;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance30;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.Caption = "Ant. Util. Comisiones";
			ultraGridColumn43.Header.VisiblePosition = 43;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 120;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance31;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 44;
			ultraGridColumn44.Width = 80;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn45.CellAppearance = appearance32;
			ultraGridColumn45.Header.Caption = "D F R";
			ultraGridColumn45.Header.VisiblePosition = 49;
			ultraGridColumn45.Width = 40;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn46.CellAppearance = appearance33;
			ultraGridColumn46.Header.Caption = "F R A";
			ultraGridColumn46.Header.VisiblePosition = 50;
			ultraGridColumn46.Width = 40;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance34;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.Caption = "F R Acum";
			ultraGridColumn47.Header.VisiblePosition = 51;
			ultraGridColumn47.Width = 80;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance35;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.Caption = "F R No Acum";
			ultraGridColumn48.Header.VisiblePosition = 52;
			ultraGridColumn48.Width = 80;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance36.BackColor = System.Drawing.Color.LightBlue;
			appearance36.BackColor2 = System.Drawing.Color.LightBlue;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance36;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.Caption = "Ingresos";
			ultraGridColumn49.Header.VisiblePosition = 53;
			ultraGridColumn49.Width = 80;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance37;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.Caption = "Base IESS";
			ultraGridColumn50.Header.VisiblePosition = 54;
			ultraGridColumn50.Width = 80;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn51.CellAppearance = appearance38;
			ultraGridColumn51.Header.Caption = "D A";
			ultraGridColumn51.Header.VisiblePosition = 55;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 40;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance39;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.Caption = "13er Sueldo";
			ultraGridColumn52.Header.VisiblePosition = 46;
			ultraGridColumn52.Width = 80;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance40;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.Caption = "14to Sueldo";
			ultraGridColumn53.Header.VisiblePosition = 48;
			ultraGridColumn53.Width = 80;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance41;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.VisiblePosition = 88;
			ultraGridColumn54.Width = 80;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance42;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.Caption = "Ap. Patronal";
			ultraGridColumn55.Header.VisiblePosition = 87;
			ultraGridColumn55.Width = 80;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance43;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.Caption = "Ap. Personal";
			ultraGridColumn56.Header.VisiblePosition = 56;
			ultraGridColumn56.Width = 80;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn57.CellAppearance = appearance44;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.Caption = "Extensión";
			ultraGridColumn57.Header.VisiblePosition = 60;
			ultraGridColumn57.Width = 80;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance45;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.Caption = "P Quirografario";
			ultraGridColumn58.Header.VisiblePosition = 62;
			ultraGridColumn58.PromptChar = ' ';
			ultraGridColumn58.Width = 80;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance46;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.Caption = "P Hipotecario";
			ultraGridColumn59.Header.VisiblePosition = 61;
			ultraGridColumn59.PromptChar = ' ';
			ultraGridColumn59.Width = 80;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance47;
			ultraGridColumn60.Format = "#,##0.00";
			ultraGridColumn60.Header.Caption = "Pago I.E.S.S.";
			ultraGridColumn60.Header.VisiblePosition = 90;
			ultraGridColumn60.Width = 80;
			ultraGridColumn61.Header.VisiblePosition = 92;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance48;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.Caption = "Credito";
			ultraGridColumn62.Header.VisiblePosition = 63;
			ultraGridColumn62.PromptChar = ' ';
			ultraGridColumn62.Width = 80;
			ultraGridColumn63.Header.VisiblePosition = 95;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 93;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance49;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.Caption = "Credi Point";
			ultraGridColumn65.Header.VisiblePosition = 65;
			ultraGridColumn65.PromptChar = ' ';
			ultraGridColumn65.Width = 80;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance50;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.Caption = "Falt. Caja";
			ultraGridColumn66.Header.VisiblePosition = 66;
			ultraGridColumn66.PromptChar = ' ';
			ultraGridColumn66.Width = 80;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance51;
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.VisiblePosition = 68;
			ultraGridColumn67.PromptChar = ' ';
			ultraGridColumn67.Width = 80;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance52;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.VisiblePosition = 71;
			ultraGridColumn68.Width = 80;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance53;
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.VisiblePosition = 73;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.PromptChar = ' ';
			ultraGridColumn69.Width = 80;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance54;
			ultraGridColumn70.Format = "#,##0.00";
			ultraGridColumn70.Header.Caption = "Com Pagadas";
			ultraGridColumn70.Header.VisiblePosition = 72;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance55;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellButtonAppearance = appearance56;
			ultraGridColumn71.Format = "#,##0.00";
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.FormulaErrorAppearance = appearance57;
			ultraGridColumn71.Header.Caption = "Bonos Pagados";
			ultraGridColumn71.Header.VisiblePosition = 57;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.MaskLiteralsAppearance = appearance58;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance59;
			ultraGridColumn72.Format = "#,##0.00";
			ultraGridColumn72.Header.Caption = "Anticipos";
			ultraGridColumn72.Header.VisiblePosition = 74;
			ultraGridColumn72.PromptChar = ' ';
			ultraGridColumn72.Width = 80;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance60;
			ultraGridColumn73.Format = "#,##0.00";
			ultraGridColumn73.Header.Caption = "Seguro";
			ultraGridColumn73.Header.VisiblePosition = 76;
			ultraGridColumn73.PromptChar = ' ';
			ultraGridColumn73.Width = 80;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance61;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.Caption = "Impuesto A La Renta";
			ultraGridColumn74.Header.VisiblePosition = 77;
			ultraGridColumn74.PromptChar = ' ';
			ultraGridColumn74.Width = 80;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance62;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellButtonAppearance = appearance63;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.VisiblePosition = 79;
			ultraGridColumn75.Width = 80;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance64;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellButtonAppearance = appearance65;
			ultraGridColumn76.Format = "#,##0.00";
			ultraGridColumn76.Header.VisiblePosition = 80;
			ultraGridColumn76.Width = 80;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance66;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.Caption = "Otros";
			ultraGridColumn77.Header.VisiblePosition = 81;
			ultraGridColumn77.PromptChar = ' ';
			ultraGridColumn77.Width = 80;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance67;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.VisiblePosition = 83;
			ultraGridColumn79.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance68;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.VisiblePosition = 84;
			appearance69.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			appearance69.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance69;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.Caption = "Ret Judicial";
			ultraGridColumn80.Header.VisiblePosition = 78;
			ultraGridColumn80.Width = 80;
			ultraGridColumn81.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance70.BackColor = System.Drawing.Color.LightGreen;
			appearance70.BackColor2 = System.Drawing.Color.LightGreen;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn81.CellAppearance = appearance70;
			ultraGridColumn81.Format = "#,##0.00";
			ultraGridColumn81.Header.Caption = "Egresos";
			ultraGridColumn81.Header.VisiblePosition = 85;
			ultraGridColumn81.Width = 80;
			ultraGridColumn82.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance71.BackColor = System.Drawing.Color.CadetBlue;
			appearance71.BackColor2 = System.Drawing.Color.CadetBlue;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn82.CellAppearance = appearance71;
			ultraGridColumn82.Format = "#,##0.00";
			ultraGridColumn82.Header.VisiblePosition = 86;
			ultraGridColumn82.Width = 80;
			ultraGridColumn83.Header.VisiblePosition = 97;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 89;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 96;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn86.Header.VisiblePosition = 91;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 98;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn88.Header.VisiblePosition = 21;
			ultraGridColumn89.Header.VisiblePosition = 94;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn90.CellAppearance = appearance72;
			ultraGridColumn90.Header.Caption = "13 A";
			ultraGridColumn90.Header.VisiblePosition = 45;
			ultraGridColumn90.Width = 50;
			ultraGridColumn91.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance73.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn91.CellAppearance = appearance73;
			ultraGridColumn91.Header.Caption = "14 A";
			ultraGridColumn91.Header.VisiblePosition = 47;
			ultraGridColumn91.Width = 50;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn92.CellAppearance = appearance74;
			ultraGridColumn92.Format = "#,##0.00";
			ultraGridColumn92.Header.Caption = "Caja Chica";
			ultraGridColumn92.Header.VisiblePosition = 69;
			ultraGridColumn93.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn93.CellAppearance = appearance75;
			ultraGridColumn93.Format = "#,##0.00";
			ultraGridColumn93.Header.VisiblePosition = 30;
			ultraGridColumn94.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn94.CellAppearance = appearance76;
			ultraGridColumn94.Format = "#,##0.00";
			ultraGridColumn94.Header.VisiblePosition = 59;
			ultraGridColumn95.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance77.Image = ((object)(resources.GetObject("appearance77.Image")));
			ultraGridColumn95.CellButtonAppearance = appearance77;
			ultraGridColumn95.Header.Caption = "";
			ultraGridColumn95.Header.VisiblePosition = 70;
			ultraGridColumn95.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn95.Width = 30;
			ultraGridColumn96.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance78.Image = ((object)(resources.GetObject("appearance78.Image")));
			ultraGridColumn96.CellButtonAppearance = appearance78;
			ultraGridColumn96.Header.Caption = "";
			ultraGridColumn96.Header.VisiblePosition = 67;
			ultraGridColumn96.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn96.Width = 30;
			ultraGridColumn97.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance79.Image = ((object)(resources.GetObject("appearance79.Image")));
			ultraGridColumn97.CellButtonAppearance = appearance79;
			ultraGridColumn97.Header.Caption = "";
			ultraGridColumn97.Header.VisiblePosition = 82;
			ultraGridColumn97.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn97.Width = 30;
			ultraGridColumn98.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance80.Image = ((object)(resources.GetObject("appearance80.Image")));
			ultraGridColumn98.CellButtonAppearance = appearance80;
			ultraGridColumn98.Header.Caption = "";
			ultraGridColumn98.Header.VisiblePosition = 75;
			ultraGridColumn98.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn98.Width = 30;
			ultraGridColumn99.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance81.Image = ((object)(resources.GetObject("appearance81.Image")));
			ultraGridColumn99.CellButtonAppearance = appearance81;
			ultraGridColumn99.Header.Caption = "";
			ultraGridColumn99.Header.VisiblePosition = 64;
			ultraGridColumn99.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn99.Width = 30;
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
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68,
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73,
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79,
																										 ultraGridColumn80,
																										 ultraGridColumn81,
																										 ultraGridColumn82,
																										 ultraGridColumn83,
																										 ultraGridColumn84,
																										 ultraGridColumn85,
																										 ultraGridColumn86,
																										 ultraGridColumn87,
																										 ultraGridColumn88,
																										 ultraGridColumn89,
																										 ultraGridColumn90,
																										 ultraGridColumn91,
																										 ultraGridColumn92,
																										 ultraGridColumn93,
																										 ultraGridColumn94,
																										 ultraGridColumn95,
																										 ultraGridColumn96,
																										 ultraGridColumn97,
																										 ultraGridColumn98,
																										 ultraGridColumn99});
			appearance82.TextHAlign = Infragistics.Win.HAlign.Center;
			summarySettings1.Appearance = appearance82;
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Center;
			summarySettings2.Appearance = appearance83;
			summarySettings2.DisplayFormat = "{0: #,##0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Center;
			summarySettings3.Appearance = appearance84;
			summarySettings3.DisplayFormat = "{0: #,##0}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Center;
			summarySettings4.Appearance = appearance85;
			summarySettings4.DisplayFormat = "{0: #,##0}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance86.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance86;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance87.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance87;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance88.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance88;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance89.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance89;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.Hidden = true;
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance90.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance90;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance91.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance91;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.Hidden = true;
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance92.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance92;
			summarySettings11.DisplayFormat = "{0: #,##0}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance93.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance93;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance94;
			summarySettings13.DisplayFormat = "{0: #,##0}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance95;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance96.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance96;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance97.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance97;
			summarySettings16.DisplayFormat = "{0: #,##0}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance98.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance98;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance99.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance99;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance100.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance100;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.Hidden = true;
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance101.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance101;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.Hidden = true;
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance102.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance102;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance103.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance103;
			summarySettings22.DisplayFormat = "{0: #,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance104.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance104;
			summarySettings23.DisplayFormat = "{0: #,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance105.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance105;
			summarySettings24.DisplayFormat = "{0: #,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance106.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance106;
			summarySettings25.DisplayFormat = "{0: #,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance107.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance107;
			summarySettings26.DisplayFormat = "{0: #,##0.00}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance108.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings27.Appearance = appearance108;
			summarySettings27.DisplayFormat = "{0: #,##0.00}";
			summarySettings27.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance109.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings28.Appearance = appearance109;
			summarySettings28.DisplayFormat = "{0: #,##0.00}";
			summarySettings28.Hidden = true;
			summarySettings28.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings29.Appearance = appearance110;
			summarySettings29.DisplayFormat = "{0: #,##0.00}";
			summarySettings29.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance111.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings30.Appearance = appearance111;
			summarySettings30.DisplayFormat = "{0: #,##0.00}";
			summarySettings30.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance112.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings31.Appearance = appearance112;
			summarySettings31.DisplayFormat = "{0: #,##0.00}";
			summarySettings31.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance113.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings32.Appearance = appearance113;
			summarySettings32.DisplayFormat = "{0: #,##0.00}";
			summarySettings32.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance114.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings33.Appearance = appearance114;
			summarySettings33.DisplayFormat = "{0: #,##0.00}";
			summarySettings33.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance115.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings34.Appearance = appearance115;
			summarySettings34.DisplayFormat = "{0: #,##0.00}";
			summarySettings34.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance116.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings35.Appearance = appearance116;
			summarySettings35.DisplayFormat = "{0: #,##0.00}";
			summarySettings35.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance117.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings36.Appearance = appearance117;
			summarySettings36.DisplayFormat = "{0: #,##0.00}";
			summarySettings36.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance118.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings37.Appearance = appearance118;
			summarySettings37.DisplayFormat = "{0: #,##0}";
			summarySettings37.Hidden = true;
			summarySettings37.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance119.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings38.Appearance = appearance119;
			summarySettings38.DisplayFormat = "{0: #,##0.00}";
			summarySettings38.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance120.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings39.Appearance = appearance120;
			summarySettings39.DisplayFormat = "{0: #,##0.00}";
			summarySettings39.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance121.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings40.Appearance = appearance121;
			summarySettings40.DisplayFormat = "{0: #,##0}";
			summarySettings40.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance122.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings41.Appearance = appearance122;
			summarySettings41.DisplayFormat = "{0: #,##0.00}";
			summarySettings41.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance123.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings42.Appearance = appearance123;
			summarySettings42.DisplayFormat = "{0: #,##0.00}";
			summarySettings42.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance124.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings43.Appearance = appearance124;
			summarySettings43.DisplayFormat = "{0: #,##0.00}";
			summarySettings43.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance125.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings44.Appearance = appearance125;
			summarySettings44.DisplayFormat = "{0: #,##0.00}";
			summarySettings44.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance126.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings45.Appearance = appearance126;
			summarySettings45.DisplayFormat = "{0: #,##0.00}";
			summarySettings45.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance127.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings46.Appearance = appearance127;
			summarySettings46.DisplayFormat = "{0: #,##0.00}";
			summarySettings46.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance128.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings47.Appearance = appearance128;
			summarySettings47.DisplayFormat = "{0: #,##0.00}";
			summarySettings47.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance129.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings48.Appearance = appearance129;
			summarySettings48.DisplayFormat = "{0: #,##0.00}";
			summarySettings48.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance130.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings49.Appearance = appearance130;
			summarySettings49.DisplayFormat = "{0: #,##0.00}";
			summarySettings49.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance131.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings50.Appearance = appearance131;
			summarySettings50.DisplayFormat = "{0: #,##0.00}";
			summarySettings50.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance132.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings51.Appearance = appearance132;
			summarySettings51.DisplayFormat = "{0: #,##0.00}";
			summarySettings51.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance133.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings52.Appearance = appearance133;
			summarySettings52.DisplayFormat = "{0: #,##0.00}";
			summarySettings52.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance134.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings53.Appearance = appearance134;
			summarySettings53.DisplayFormat = "{0: #,##0.00}";
			summarySettings53.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance135.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings54.Appearance = appearance135;
			summarySettings54.DisplayFormat = "{0: #,##0.00}";
			summarySettings54.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance136.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings55.Appearance = appearance136;
			summarySettings55.DisplayFormat = "{0: #,##0.00}";
			summarySettings55.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance137.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings56.Appearance = appearance137;
			summarySettings56.DisplayFormat = "{0: #,##0.00}";
			summarySettings56.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance138.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings57.Appearance = appearance138;
			summarySettings57.DisplayFormat = "{0: #,##0.00}";
			summarySettings57.Hidden = true;
			summarySettings57.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance139.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings58.Appearance = appearance139;
			summarySettings58.DisplayFormat = "{0: #,##0.00}";
			summarySettings58.Hidden = true;
			summarySettings58.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance140.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings59.Appearance = appearance140;
			summarySettings59.DisplayFormat = "{0: #,##0.00}";
			summarySettings59.Hidden = true;
			summarySettings59.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance141.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings60.Appearance = appearance141;
			summarySettings60.DisplayFormat = "{0: #,##0.00}";
			summarySettings60.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance142.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings61.Appearance = appearance142;
			summarySettings61.DisplayFormat = "{0: #,##0.00}";
			summarySettings61.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance143.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings62.Appearance = appearance143;
			summarySettings62.DisplayFormat = "{0: #,##0.00}";
			summarySettings62.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance144.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings63.Appearance = appearance144;
			summarySettings63.DisplayFormat = "{0: #,##0.00}";
			summarySettings63.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance145.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings64.Appearance = appearance145;
			summarySettings64.DisplayFormat = "{0: #,##0.00}";
			summarySettings64.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance146.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings65.Appearance = appearance146;
			summarySettings65.DisplayFormat = "{0: #,##0.00}";
			summarySettings65.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance147.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings66.Appearance = appearance147;
			summarySettings66.DisplayFormat = "{0: #,##0.00}";
			summarySettings66.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance148.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings67.Appearance = appearance148;
			summarySettings67.DisplayFormat = "{0: #,##0.00}";
			summarySettings67.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
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
																																															summarySettings27,
																																															summarySettings28,
																																															summarySettings29,
																																															summarySettings30,
																																															summarySettings31,
																																															summarySettings32,
																																															summarySettings33,
																																															summarySettings34,
																																															summarySettings35,
																																															summarySettings36,
																																															summarySettings37,
																																															summarySettings38,
																																															summarySettings39,
																																															summarySettings40,
																																															summarySettings41,
																																															summarySettings42,
																																															summarySettings43,
																																															summarySettings44,
																																															summarySettings45,
																																															summarySettings46,
																																															summarySettings47,
																																															summarySettings48,
																																															summarySettings49,
																																															summarySettings50,
																																															summarySettings51,
																																															summarySettings52,
																																															summarySettings53,
																																															summarySettings54,
																																															summarySettings55,
																																															summarySettings56,
																																															summarySettings57,
																																															summarySettings58,
																																															summarySettings59,
																																															summarySettings60,
																																															summarySettings61,
																																															summarySettings62,
																																															summarySettings63,
																																															summarySettings64,
																																															summarySettings65,
																																															summarySettings66,
																																															summarySettings67});
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.uGridRoles.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridRoles.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridRoles.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridRoles.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance149.BackColor = System.Drawing.Color.Transparent;
			this.uGridRoles.DisplayLayout.Override.CardAreaAppearance = appearance149;
			appearance150.ForeColor = System.Drawing.Color.Black;
			appearance150.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridRoles.DisplayLayout.Override.CellAppearance = appearance150;
			appearance151.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance151.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance151.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance151.FontData.BoldAsString = "True";
			appearance151.FontData.Name = "Arial";
			appearance151.FontData.SizeInPoints = 8.5F;
			appearance151.ForeColor = System.Drawing.Color.White;
			appearance151.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRoles.DisplayLayout.Override.HeaderAppearance = appearance151;
			this.uGridRoles.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance152.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance152.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance152.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRoles.DisplayLayout.Override.RowSelectorAppearance = appearance152;
			appearance153.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance153.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance153.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRoles.DisplayLayout.Override.SelectedRowAppearance = appearance153;
			this.uGridRoles.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRoles.Location = new System.Drawing.Point(8, 120);
			this.uGridRoles.Name = "uGridRoles";
			this.uGridRoles.Size = new System.Drawing.Size(1230, 256);
			this.uGridRoles.TabIndex = 5;
			this.uGridRoles.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridRoles_BeforeRowsDeleted);
			this.uGridRoles.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridRoles_KeyDown);
			this.uGridRoles.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridRoles_ClickCellButton);
			this.uGridRoles.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridRoles_AfterCellUpdate);
			this.uGridRoles.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridRoles_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn22.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
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
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn30.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn31.DataType = typeof(System.Decimal);
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
			ultraDataColumn37.DataType = typeof(int);
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
			ultraDataColumn40.DataType = typeof(int);
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
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn44.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(System.Decimal);
			ultraDataColumn47.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn48.DataType = typeof(System.Decimal);
			ultraDataColumn48.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn49.DataType = typeof(System.Decimal);
			ultraDataColumn49.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn50.DataType = typeof(System.Decimal);
			ultraDataColumn50.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn51.DataType = typeof(int);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn52.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn53.DataType = typeof(System.Decimal);
			ultraDataColumn53.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn54.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn55.DataType = typeof(System.Decimal);
			ultraDataColumn55.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn56.DataType = typeof(System.Decimal);
			ultraDataColumn56.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn57.DataType = typeof(System.Decimal);
			ultraDataColumn58.DataType = typeof(System.Decimal);
			ultraDataColumn58.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn59.DataType = typeof(System.Decimal);
			ultraDataColumn59.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn60.DataType = typeof(System.Decimal);
			ultraDataColumn60.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn61.DataType = typeof(int);
			ultraDataColumn62.DataType = typeof(System.Decimal);
			ultraDataColumn62.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn64.DataType = typeof(int);
			ultraDataColumn65.DataType = typeof(System.Decimal);
			ultraDataColumn65.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn66.DataType = typeof(System.Decimal);
			ultraDataColumn66.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn67.DataType = typeof(System.Decimal);
			ultraDataColumn67.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn68.DataType = typeof(System.Decimal);
			ultraDataColumn68.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn69.DataType = typeof(System.Decimal);
			ultraDataColumn69.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn70.DataType = typeof(System.Decimal);
			ultraDataColumn70.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn71.DataType = typeof(int);
			ultraDataColumn72.DataType = typeof(System.Decimal);
			ultraDataColumn72.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn73.DataType = typeof(System.Decimal);
			ultraDataColumn73.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn74.DataType = typeof(System.Decimal);
			ultraDataColumn74.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn75.DataType = typeof(System.Decimal);
			ultraDataColumn75.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn76.DataType = typeof(System.Decimal);
			ultraDataColumn76.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn77.DataType = typeof(System.Decimal);
			ultraDataColumn77.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn78.DataType = typeof(System.Decimal);
			ultraDataColumn78.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn79.DataType = typeof(System.Decimal);
			ultraDataColumn80.DataType = typeof(System.Decimal);
			ultraDataColumn80.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn81.DataType = typeof(System.Decimal);
			ultraDataColumn81.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn82.DataType = typeof(System.Decimal);
			ultraDataColumn82.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn83.DataType = typeof(bool);
			ultraDataColumn83.DefaultValue = false;
			ultraDataColumn89.DataType = typeof(int);
			ultraDataColumn90.DataType = typeof(int);
			ultraDataColumn91.DataType = typeof(int);
			ultraDataColumn92.DataType = typeof(System.Decimal);
			ultraDataColumn93.DataType = typeof(System.Decimal);
			ultraDataColumn94.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn54,
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58,
																																 ultraDataColumn59,
																																 ultraDataColumn60,
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63,
																																 ultraDataColumn64,
																																 ultraDataColumn65,
																																 ultraDataColumn66,
																																 ultraDataColumn67,
																																 ultraDataColumn68,
																																 ultraDataColumn69,
																																 ultraDataColumn70,
																																 ultraDataColumn71,
																																 ultraDataColumn72,
																																 ultraDataColumn73,
																																 ultraDataColumn74,
																																 ultraDataColumn75,
																																 ultraDataColumn76,
																																 ultraDataColumn77,
																																 ultraDataColumn78,
																																 ultraDataColumn79,
																																 ultraDataColumn80,
																																 ultraDataColumn81,
																																 ultraDataColumn82,
																																 ultraDataColumn83,
																																 ultraDataColumn84,
																																 ultraDataColumn85,
																																 ultraDataColumn86,
																																 ultraDataColumn87,
																																 ultraDataColumn88,
																																 ultraDataColumn89,
																																 ultraDataColumn90,
																																 ultraDataColumn91,
																																 ultraDataColumn92,
																																 ultraDataColumn93,
																																 ultraDataColumn94});
			// 
			// dtpPeriodo
			// 
			appearance154.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpPeriodo.Appearance = appearance154;
			this.dtpPeriodo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpPeriodo.DateButtons.Add(dateButton1);
			this.dtpPeriodo.Format = "MM/yyyy";
			this.dtpPeriodo.Location = new System.Drawing.Point(104, 8);
			this.dtpPeriodo.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpPeriodo.Name = "dtpPeriodo";
			this.dtpPeriodo.NonAutoSizeHeight = 23;
			this.dtpPeriodo.Size = new System.Drawing.Size(112, 21);
			this.dtpPeriodo.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtpPeriodo.SpinButtonsVisible = true;
			this.dtpPeriodo.TabIndex = 687;
			this.dtpPeriodo.Value = ((object)(resources.GetObject("dtpPeriodo.Value")));
			this.dtpPeriodo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpPeriodo_KeyDown);
			this.dtpPeriodo.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtpPeriodo_BeforeDropDown);
			this.dtpPeriodo.ValueChanged += new System.EventHandler(this.dtpPeriodo_ValueChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(42, 17);
			this.label6.TabIndex = 686;
			this.label6.Text = "Periodo";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Enabled = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(328, 72);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 707;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 75);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(82, 17);
			this.label7.TabIndex = 709;
			this.label7.Text = "Cedula/Nombre";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(104, 73);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(208, 20);
			this.txtBuscar.TabIndex = 708;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			this.txtBuscar.ValueChanged += new System.EventHandler(this.txtBuscar_ValueChanged);
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Enabled = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(672, 8);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(88, 23);
			this.btnGenerar.TabIndex = 710;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// btnContabilizar
			// 
			this.btnContabilizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnContabilizar.CausesValidation = false;
			this.btnContabilizar.Enabled = false;
			this.btnContabilizar.Image = ((System.Drawing.Image)(resources.GetObject("btnContabilizar.Image")));
			this.btnContabilizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnContabilizar.Location = new System.Drawing.Point(864, 40);
			this.btnContabilizar.Name = "btnContabilizar";
			this.btnContabilizar.Size = new System.Drawing.Size(88, 23);
			this.btnContabilizar.TabIndex = 711;
			this.btnContabilizar.Text = "Contabilizar";
			this.btnContabilizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnContabilizar.Click += new System.EventHandler(this.btnContabilizar_Click);
			// 
			// txtIdRolNomina
			// 
			this.txtIdRolNomina.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdRolNomina.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdRolNomina.Enabled = false;
			this.txtIdRolNomina.Location = new System.Drawing.Point(502, 7);
			this.txtIdRolNomina.Name = "txtIdRolNomina";
			this.txtIdRolNomina.PromptChar = ' ';
			this.txtIdRolNomina.Size = new System.Drawing.Size(16, 22);
			this.txtIdRolNomina.TabIndex = 712;
			this.txtIdRolNomina.Visible = false;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1152, 40);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(88, 23);
			this.btnSalir.TabIndex = 706;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1152, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 713;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1246, 8);
			this.groupBox1.TabIndex = 714;
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
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(408, 9);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 18);
			this.lblEstado.TabIndex = 715;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(664, 74);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 18);
			this.lblContador.TabIndex = 716;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(478, 7);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 717;
			this.txtIdAsiento.Visible = false;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(1056, 40);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(88, 23);
			this.btnImprimir.TabIndex = 718;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnAcreditacion
			// 
			this.btnAcreditacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAcreditacion.CausesValidation = false;
			this.btnAcreditacion.Enabled = false;
			this.btnAcreditacion.Image = ((System.Drawing.Image)(resources.GetObject("btnAcreditacion.Image")));
			this.btnAcreditacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAcreditacion.Location = new System.Drawing.Point(768, 40);
			this.btnAcreditacion.Name = "btnAcreditacion";
			this.btnAcreditacion.Size = new System.Drawing.Size(88, 23);
			this.btnAcreditacion.TabIndex = 719;
			this.btnAcreditacion.Text = "Acreditación";
			this.btnAcreditacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAcreditacion.Click += new System.EventHandler(this.btnAcreditacion_Click);
			// 
			// btnAutorizar
			// 
			this.btnAutorizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAutorizar.CausesValidation = false;
			this.btnAutorizar.Enabled = false;
			this.btnAutorizar.Image = ((System.Drawing.Image)(resources.GetObject("btnAutorizar.Image")));
			this.btnAutorizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAutorizar.Location = new System.Drawing.Point(960, 8);
			this.btnAutorizar.Name = "btnAutorizar";
			this.btnAutorizar.Size = new System.Drawing.Size(88, 23);
			this.btnAutorizar.TabIndex = 722;
			this.btnAutorizar.Text = "Autorizar";
			this.btnAutorizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAutorizar.Click += new System.EventHandler(this.btnAutorizar_Click);
			// 
			// ultraGridExcelExporter1
			// 
			this.ultraGridExcelExporter1.BeginExport += new Infragistics.Win.UltraWinGrid.ExcelExport.BeginExportEventHandler(this.ultraGridExcelExporter1_BeginExport);
			// 
			// saveFileDialog1
			// 
			this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Enabled = false;
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizar.Location = new System.Drawing.Point(1056, 8);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(88, 23);
			this.btnActualizar.TabIndex = 724;
			this.btnActualizar.Text = "Actualizar";
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// lblDiscapacitados
			// 
			this.lblDiscapacitados.AutoSize = true;
			this.lblDiscapacitados.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDiscapacitados.ForeColor = System.Drawing.Color.Firebrick;
			this.lblDiscapacitados.Location = new System.Drawing.Point(944, 74);
			this.lblDiscapacitados.Name = "lblDiscapacitados";
			this.lblDiscapacitados.Size = new System.Drawing.Size(0, 18);
			this.lblDiscapacitados.TabIndex = 725;
			this.lblDiscapacitados.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbNEDepartamento
			// 
			this.cmbNEDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEDepartamento.DataSource = this.ultraDataSource8;
			ultraGridBand2.AddButtonCaption = "RolTipo";
			ultraGridColumn100.Header.VisiblePosition = 0;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 1;
			ultraGridColumn101.Width = 208;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn100,
																										 ultraGridColumn101});
			this.cmbNEDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNEDepartamento.DisplayMember = "Descripcion";
			this.cmbNEDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEDepartamento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEDepartamento.Location = new System.Drawing.Point(416, 40);
			this.cmbNEDepartamento.Name = "cmbNEDepartamento";
			this.cmbNEDepartamento.Size = new System.Drawing.Size(208, 22);
			this.cmbNEDepartamento.TabIndex = 729;
			this.cmbNEDepartamento.ValueMember = "idDepartamento";
			this.cmbNEDepartamento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEDepartamento_KeyDown);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn95.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn95,
																																 ultraDataColumn96});
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(8, 43);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(86, 17);
			this.label33.TabIndex = 726;
			this.label33.Text = "Centro De Costo";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(328, 43);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(76, 17);
			this.label12.TabIndex = 728;
			this.label12.Text = "Departamento";
			// 
			// cmbNECentroCosto
			// 
			this.cmbNECentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECentroCosto.DataSource = this.ultraDataSource7;
			ultraGridBand3.AddButtonCaption = "RolTipo";
			ultraGridColumn102.Header.VisiblePosition = 1;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn102.Width = 85;
			ultraGridColumn103.Header.VisiblePosition = 0;
			ultraGridColumn103.Width = 208;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn102,
																										 ultraGridColumn103});
			this.cmbNECentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbNECentroCosto.DisplayMember = "Nombre";
			this.cmbNECentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECentroCosto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECentroCosto.Location = new System.Drawing.Point(104, 40);
			this.cmbNECentroCosto.Name = "cmbNECentroCosto";
			this.cmbNECentroCosto.Size = new System.Drawing.Size(208, 22);
			this.cmbNECentroCosto.TabIndex = 727;
			this.cmbNECentroCosto.ValueMember = "idProyecto";
			this.cmbNECentroCosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNECentroCosto_KeyDown);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn97.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn97,
																																 ultraDataColumn98});
			// 
			// btnSubirArchivos
			// 
			this.btnSubirArchivos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSubirArchivos.CausesValidation = false;
			this.btnSubirArchivos.Enabled = false;
			this.btnSubirArchivos.Image = ((System.Drawing.Image)(resources.GetObject("btnSubirArchivos.Image")));
			this.btnSubirArchivos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSubirArchivos.Location = new System.Drawing.Point(768, 8);
			this.btnSubirArchivos.Name = "btnSubirArchivos";
			this.btnSubirArchivos.Size = new System.Drawing.Size(88, 23);
			this.btnSubirArchivos.TabIndex = 730;
			this.btnSubirArchivos.Text = "Importar";
			this.btnSubirArchivos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSubirArchivos.Click += new System.EventHandler(this.btnSubirArchivos_Click);
			// 
			// mnuSubeArchivos
			// 
			this.mnuSubeArchivos.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																										this.mnuSAPQ,
																																										this.mnuSAPH,
																																										this.mnuSAF,
																																										this.menuItem1,
																																										this.menuItem2,
																																										this.menuItem3,
																																										this.menuItem4,
																																										this.menuItem5,
																																										this.menuItem6,
																																										this.menuItem7});
			// 
			// mnuSAPQ
			// 
			this.mnuSAPQ.Index = 0;
			this.mnuSAPQ.Text = "Préstamos Quirografarios";
			this.mnuSAPQ.Click += new System.EventHandler(this.mnuSAPQ_Click);
			// 
			// mnuSAPH
			// 
			this.mnuSAPH.Index = 1;
			this.mnuSAPH.Text = "Préstamos Hipotecarios";
			this.mnuSAPH.Click += new System.EventHandler(this.mnuSAPH_Click);
			// 
			// mnuSAF
			// 
			this.mnuSAF.Index = 2;
			this.mnuSAF.Text = "Farmacias";
			this.mnuSAF.Click += new System.EventHandler(this.mnuSAF_Click);
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 3;
			this.menuItem1.Text = "Caja Chica";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 4;
			this.menuItem2.Text = "Depreciacion";
			this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 5;
			this.menuItem3.Text = "Multas";
			this.menuItem3.Click += new System.EventHandler(this.menuItem3_Click);
			// 
			// menuItem4
			// 
			this.menuItem4.Index = 6;
			this.menuItem4.Text = "Seguros";
			this.menuItem4.Click += new System.EventHandler(this.menuItem4_Click);
			// 
			// menuItem5
			// 
			this.menuItem5.Index = 7;
			this.menuItem5.Text = "Otros";
			this.menuItem5.Click += new System.EventHandler(this.menuItem5_Click);
			// 
			// menuItem6
			// 
			this.menuItem6.Index = 8;
			this.menuItem6.Text = "Credito";
			this.menuItem6.Click += new System.EventHandler(this.menuItem6_Click);
			// 
			// menuItem7
			// 
			this.menuItem7.Index = 9;
			this.menuItem7.Text = "CrediPoint";
			this.menuItem7.Click += new System.EventHandler(this.menuItem7_Click);
			// 
			// btnSeleccionar
			// 
			this.btnSeleccionar.CausesValidation = false;
			this.btnSeleccionar.Image = ((System.Drawing.Image)(resources.GetObject("btnSeleccionar.Image")));
			this.btnSeleccionar.Location = new System.Drawing.Point(27, 120);
			this.btnSeleccionar.Name = "btnSeleccionar";
			this.btnSeleccionar.Size = new System.Drawing.Size(30, 22);
			this.btnSeleccionar.TabIndex = 731;
			this.btnSeleccionar.Visible = false;
			this.btnSeleccionar.Click += new System.EventHandler(this.btnSeleccionar_Click);
			// 
			// mnuAcreditacion
			// 
			this.mnuAcreditacion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																										this.mnuAcreditacionTransferencia,
																																										this.mnuAcreditacionCheques,
																																										this.mnuAcreditacionRJ});
			// 
			// mnuAcreditacionTransferencia
			// 
			this.mnuAcreditacionTransferencia.Index = 0;
			this.mnuAcreditacionTransferencia.Text = "Transferencia";
			this.mnuAcreditacionTransferencia.Click += new System.EventHandler(this.mnuAcreditacionTransferencia_Click);
			// 
			// mnuAcreditacionCheques
			// 
			this.mnuAcreditacionCheques.Index = 1;
			this.mnuAcreditacionCheques.Text = "Cheques";
			this.mnuAcreditacionCheques.Click += new System.EventHandler(this.mnuAcreditacionCheques_Click);
			// 
			// mnuAcreditacionRJ
			// 
			this.mnuAcreditacionRJ.Index = 2;
			this.mnuAcreditacionRJ.Text = "Retención Judicial";
			this.mnuAcreditacionRJ.Click += new System.EventHandler(this.mnuAcreditacionRJ_Click);
			// 
			// btnRevisado
			// 
			this.btnRevisado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnRevisado.CausesValidation = false;
			this.btnRevisado.Enabled = false;
			this.btnRevisado.Image = ((System.Drawing.Image)(resources.GetObject("btnRevisado.Image")));
			this.btnRevisado.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRevisado.Location = new System.Drawing.Point(864, 8);
			this.btnRevisado.Name = "btnRevisado";
			this.btnRevisado.Size = new System.Drawing.Size(88, 23);
			this.btnRevisado.TabIndex = 732;
			this.btnRevisado.Text = "Revisión";
			this.btnRevisado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRevisado.Click += new System.EventHandler(this.btnRevisado_Click);
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesAgregar,
																																								this.mnuAccionesActualizar});
			// 
			// mnuAccionesAgregar
			// 
			this.mnuAccionesAgregar.Index = 0;
			this.mnuAccionesAgregar.Text = "Agregar";
			this.mnuAccionesAgregar.Click += new System.EventHandler(this.mnuAccionesAgregar_Click);
			// 
			// mnuAccionesActualizar
			// 
			this.mnuAccionesActualizar.Index = 1;
			this.mnuAccionesActualizar.Text = "Actualizar";
			this.mnuAccionesActualizar.Click += new System.EventHandler(this.mnuAccionesActualizar_Click);
			// 
			// notifyIcon1
			// 
			this.notifyIcon1.Text = "notifyIcon1";
			this.notifyIcon1.Visible = true;
			// 
			// txtIdAntipoUtilidades
			// 
			this.txtIdAntipoUtilidades.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAntipoUtilidades.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAntipoUtilidades.Enabled = false;
			this.txtIdAntipoUtilidades.Location = new System.Drawing.Point(462, 8);
			this.txtIdAntipoUtilidades.Name = "txtIdAntipoUtilidades";
			this.txtIdAntipoUtilidades.PromptChar = ' ';
			this.txtIdAntipoUtilidades.Size = new System.Drawing.Size(8, 22);
			this.txtIdAntipoUtilidades.TabIndex = 749;
			this.txtIdAntipoUtilidades.Visible = false;
			// 
			// btnUtilidades
			// 
			this.btnUtilidades.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnUtilidades.CausesValidation = false;
			this.btnUtilidades.Enabled = false;
			this.btnUtilidades.Image = ((System.Drawing.Image)(resources.GetObject("btnUtilidades.Image")));
			this.btnUtilidades.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnUtilidades.Location = new System.Drawing.Point(960, 40);
			this.btnUtilidades.Name = "btnUtilidades";
			this.btnUtilidades.Size = new System.Drawing.Size(88, 23);
			this.btnUtilidades.TabIndex = 750;
			this.btnUtilidades.Text = "Utilidades";
			this.btnUtilidades.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnUtilidades.Click += new System.EventHandler(this.btnUtilidades_Click);
			// 
			// frmRolesDePago
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1246, 382);
			this.Controls.Add(this.btnUtilidades);
			this.Controls.Add(this.txtIdAntipoUtilidades);
			this.Controls.Add(this.btnRevisado);
			this.Controls.Add(this.btnSeleccionar);
			this.Controls.Add(this.btnSubirArchivos);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label33);
			this.Controls.Add(this.cmbNECentroCosto);
			this.Controls.Add(this.lblDiscapacitados);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.btnAutorizar);
			this.Controls.Add(this.btnAcreditacion);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.txtIdRolNomina);
			this.Controls.Add(this.btnContabilizar);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.dtpPeriodo);
			this.Controls.Add(this.uGridRoles);
			this.Controls.Add(this.cmbNEDepartamento);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmRolesDePago";
			this.Text = "Rol De Pagos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmRolesDePago_KeyDown);
			this.Load += new System.EventHandler(this.frmRolesDePago_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridRoles)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpPeriodo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolNomina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAntipoUtilidades)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int Estado = 0;
		private Acceso miAcceso;
		decimal dSueldoBasico = 0.00m;
		decimal dAportePersonal = 0.00m;
		decimal dAportePatronal = 0.00m;
		int	iDiasTrabajaMes = 0;
		bool bCredito_Fac = false;
		bool bCredipoint_Fac = false;
		bool bDecMovilizacion = false;
		bool DecAlimentacion = false;

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			//			if (bAnticipoUtilidades)
			//			{
			//				this.uGridRoles.DisplayLayout.Bands[0].Columns["Bonos"].CellActivation = Activation.Disabled;
			//				this.uGridRoles.DisplayLayout.Bands[0].Columns["Bono"].Hidden = false;
			//				this.uGridRoles.DisplayLayout.Bands[0].Columns["AnticipoUtilidades"].Hidden = false;
			//				this.uGridRoles.DisplayLayout.Bands[0].Columns["AnticipoUtilidadesComisiones"].Hidden = false;
			//				this.uGridRoles.DisplayLayout.Bands[0].Columns["AnticipoUtilidades"].CellActivation = Activation.Disabled;
			//			}

			int idProyecto = 0;
			int idDepartamento = 0;

			if (this.cmbNECentroCosto.ActiveRow != null) idProyecto = (int)this.cmbNECentroCosto.Value;
			if (this.cmbNEDepartamento.ActiveRow != null) idDepartamento = (int)this.cmbNEDepartamento.Value;

			string sSQL = string.Format("Exec NominaConsultaRol {0}, '{1}', {2}, {3}", this.txtIdRolNomina.Value, this.txtBuscar.Text.ToString(), idProyecto, idDepartamento);
			this.uGridRoles.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.lblContador.Text = this.uGridRoles.Rows.Count + " REGISTROS ENCONTRADOS";

			int iContadorDiscapacitados = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridRoles.Rows.All)
				iContadorDiscapacitados = iContadorDiscapacitados + (int)dr.Cells["Discapacitado"].Value;

			this.lblDiscapacitados.Text = iContadorDiscapacitados.ToString() + " PERSONAS CON DISCAPACIDAD";
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec NominaValidaPagosCheques '{0}'", Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"))) > 0)
			{
				MessageBox.Show("Existen cheques pendientes por acreditar.\n\nPara generar la nómina, asegúrese de acreditar todos los cheques pendientes generados por el modulo de bonos semanales.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de GENERAR ROL DEL PERIODO DE '{0}'", 
				Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				using (frmRolPorcentajes Porcentajes = new frmRolPorcentajes(Convert.ToDateTime(this.dtpPeriodo.Value)))
				{				
					if (DialogResult.OK == Porcentajes.ShowDialog())
					{
						this.Cursor = Cursors.WaitCursor;

						string sSQL = string.Format("Exec GuardaRol {0}, '{1}'", 
							(int)this.txtIdRolNomina.Value, Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"));
						this.txtIdRolNomina.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
										
						string sSQLGenera = string.Format("Exec GeneraRolesDePagoNomina '{0}', {1}", 
							Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"), (int)this.txtIdRolNomina.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLGenera, true);

						string sSQLEnfermedad = string.Format("Exec NominaActualizaEnfermedad '{0}', {1}", 
							Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"), (int)this.txtIdRolNomina.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLEnfermedad, true);
				
						string sSQLPorcentajes = string.Format("Exec NominaGuardaRolPorcentajes {0}, {1}, {2}, {3}", (int)this.txtIdRolNomina.Value,
							Convert.ToDecimal(Porcentajes.txtBonos.Value), Convert.ToDecimal(Porcentajes.txtHE.Value), Convert.ToDecimal(Porcentajes.txtHO.Value));
						Funcion.EjecutaSQL(cdsSeteoF, sSQLPorcentajes, true);
						
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaCalculaBonosHorasExtras {0}, {1}, {2}, {3}", 
							(int)this.txtIdRolNomina.Value, Convert.ToDecimal(Porcentajes.txtBonos.Value), Convert.ToDecimal(Porcentajes.txtHE.Value), Convert.ToDecimal(Porcentajes.txtHO.Value)));	
						
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));

						this.Cursor = Cursors.Default;
					
						MessageBox.Show(string.Format("Rol del Periodo de '{0}' Generado Correctamente", Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

						this.lblEstado.Text = "GENERADO";

						this.btnGenerar.Enabled = false;	

						this.dtpPeriodo_ValueChanged(sender, e);
						this.btnVer_Click(sender, e);
					}
				}
			}
		}

		private void dtpPeriodo_ValueChanged(object sender, System.EventArgs e)
		{			
			#region Deshabilita Controles
			this.btnGenerar.Enabled = false;
			this.btnSubirArchivos.Enabled = false;
			this.btnRevisado.Enabled = false;
			this.btnAutorizar.Enabled = false;
			this.btnAcreditacion.Enabled = false;			
			this.btnContabilizar.Enabled = false;
			this.btnActualizar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnExportar.Enabled = false;
			this.btnUtilidades.Enabled = false;
			this.btnSeleccionar.Visible = true;
			this.uGridRoles.DisplayLayout.Bands[0].Columns["Revisado"].Hidden = true;
			this.txtIdAsiento.Value = 0;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridRoles);
			#endregion Deshabilita Controles	

			string sSQL = string.Format("Exec ValidaRolExiste '{0}'", Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"));			
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdRolNomina.Value = dr.GetInt32(0);
				this.lblEstado.Text = dr.GetString(1);
				Estado = dr.GetInt32(2);	
				this.txtIdAsiento.Value = dr.GetInt32(3);
				this.txtIdAntipoUtilidades.Value = dr.GetInt32(4);
				dr.Close();

				if (miAcceso.BExportar) this.btnExportar.Enabled = true;

				if (miAcceso.Buscar) 
				{
					this.btnVer.Enabled = true;
					this.btnVer_Click(sender, e);
				}

				if (Estado == 1) //GENERADO
				{						
					if (miAcceso.BRolSubirArchivos) this.btnSubirArchivos.Enabled = true;

					if (miAcceso.Editar) 
					{
						FuncionesProcedimientos.EstadoGrids(true, this.uGridRoles);						
						this.uGridRoles.DisplayLayout.Bands[0].Columns["Revisado"].Hidden = false;
						this.btnActualizar.Enabled = true;						
					}
					if (miAcceso.BRolRevision) this.btnRevisado.Enabled = true;
				}
				else if (Estado == 2) //EN REVISIÓN
				{					
					if (miAcceso.BRolAutorizar) 
					{
						this.btnAutorizar.Enabled = true;										
						this.btnSeleccionar.Visible = true;
						this.uGridRoles.DisplayLayout.Bands[0].Columns["Revisado"].Hidden = false;
					}					
					this.btnSeleccionar.Enabled = true;
					if (miAcceso.Editar) FuncionesProcedimientos.EstadoGrids(true, this.uGridRoles);
				}
				else if (Estado == 3) //AUTORIZADO
				{	
					if (miAcceso.BRolAcreditar) this.btnAcreditacion.Enabled = true;
					if (miAcceso.Editar) FuncionesProcedimientos.EstadoGrids(true, this.uGridRoles);
					this.uGridRoles.DisplayLayout.Bands[0].Columns["Revisado"].Hidden = false;
				}
				else if (Estado == 4) //ACREDITADO
				{				
					if (miAcceso.BRolContabilizar) this.btnContabilizar.Enabled = true;
					this.uGridRoles.DisplayLayout.Bands[0].Columns["Revisado"].Hidden = false;
				}
				else if (Estado == 5) //CONTABILIZADO
				{				
					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;		
					this.uGridRoles.DisplayLayout.Bands[0].Columns["Revisado"].Hidden = false;	
					if (MenuLatinium.stIdDB == 1 || MenuLatinium.stIdDB == 33|| MenuLatinium.stIdDB == 63)	
					{
						this.btnUtilidades.Enabled = true;
					}
				}
			}
			else
			{
				dr.Close();
				this.lblEstado.Text = "PENDIENTE DE GENERAR";
				this.txtIdRolNomina.Value = 0;
				this.txtIdAsiento.Value = 0;

				if (miAcceso.BRolGenerar) this.btnGenerar.Enabled = true;
			}
			dr.Close();

			this.btnVer_Click(sender, e);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void UnloadMe()
		{
			this.Close();
		}

		//bool bAnticipoUtilidades = false;

		private void frmRolesDePago_Load(object sender, System.EventArgs e)
		{			
			miAcceso = new Acceso(cdsSeteoF, "0505");
			
			if (!Funcion.Permiso("83", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Rol de Pagos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				DateTime dtFecha = DateTime.Parse("01/04/2016");			
				this.dtpPeriodo.CalendarInfo.MinDate = dtFecha.Date;
				this.dtpPeriodo.CalendarInfo.MaxDate = Convert.ToDateTime(DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month).ToString() + "/" + DateTime.Today.Month.ToString() + "/" + DateTime.Today.Year.ToString());

				this.cmbNECentroCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 45, 0");
				this.cmbNEDepartamento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 48, 0");

				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select SueldoBasico, AportePersonal, AportePatronal, DiasTrabajaMes From Seteo", true);
				dr.Read();
				if (dr.HasRows)
				{
					dSueldoBasico = dr.GetDecimal(0);
					dAportePersonal = dr.GetDecimal(1);
					dAportePatronal = dr.GetDecimal(2);
					iDiasTrabajaMes = dr.GetInt32(3);
				}
				dr.Close();

				SqlDataReader dr1 = Funcion.rEscalarSQL(cdsSeteoF, "Select Credito_Fac, Credipoint_Fac From SeteoRol", true);
				dr1.Read();
				if (dr1.HasRows)
				{
					bCredito_Fac = dr1.GetBoolean(0);
					bCredipoint_Fac = dr1.GetBoolean(1);
				}
				dr1.Close();

				SqlDataReader dr2 = Funcion.rEscalarSQL(cdsSeteoF, "SELECT DecMovilizacion, DecAlimentacion FROM SeteoC", true);
				dr2.Read();
				if (dr2.HasRows)
				{
					bDecMovilizacion = dr2.GetBoolean(0);
					DecAlimentacion = dr2.GetBoolean(1);
				}
				dr2.Close();

				if (!bCredito_Fac)
				{
					this.uGridRoles.DisplayLayout.Bands[0].Columns["ComprasCredito"].CellActivation = Activation.AllowEdit;
					this.uGridRoles.DisplayLayout.Bands[0].Columns["btnCredito"].Hidden = true;
				}

				if (!bCredipoint_Fac) this.uGridRoles.DisplayLayout.Bands[0].Columns["ComprasCrediPoint"].CellActivation = Activation.AllowEdit;
			}

			//bAnticipoUtilidades = Funcion.bEscalarSQL(cdsSeteoF, "Select AnticipoUtilidades From Seteo", true);

			#region Cabecera
			if (MenuLatinium.stDirFacturacion == "GRETASANTORINI" || MenuLatinium.stDirFacturacion == "PROSERVICIOS")
			{
				this.uGridRoles.DisplayLayout.Bands[0].Columns["ComprasCrediPoint"].Header.Caption = "Desc. Bares";
				this.uGridRoles.DisplayLayout.Bands[0].Columns["ComprasCredito"].Header.Caption = "Prestamo Empleado";
				this.uGridRoles.DisplayLayout.Bands[0].Columns["AnticipoSueldos"].Header.Caption = "Anticipo Quincena";
				this.uGridRoles.DisplayLayout.Bands[0].Columns["SeguroMedico"].Header.Caption = "Seguro Privado";
				this.uGridRoles.DisplayLayout.Bands[0].Columns["Extension"].Header.Caption = "Ext. Conyuge";
			}
			#endregion Cabecera

			//			if (bAnticipoUtilidades)
			//			{
			//				this.uGridRoles.DisplayLayout.Bands[0].Columns["Bonos"].CellActivation = Activation.Disabled;
			//				this.uGridRoles.DisplayLayout.Bands[0].Columns["Bonos"].Hidden = false;
			//				this.uGridRoles.DisplayLayout.Bands[0].Columns["AnticipoUtilidades"].Hidden = false;
			//				this.uGridRoles.DisplayLayout.Bands[0].Columns["AnticipoUtilidadesComisiones"].Hidden = false;
			//			}
			//Infragistics.Win.UltraWinToolbars.UltraToolbarsManager TB = new Infragistics.Win.UltraWinToolbars.UltraToolbarsManager();
			//			Infragistics.Win.UltraWinToolbars.UltraToolbarsManager TB = new Infragistics.Win.UltraWinToolbars.UltraToolbarsManager();	
			//			Infragistics.Win.UltraWinToolbars.ToolsCollection TC;
			//
			//			TC.AddTool("Notificacion");
			//			TB.Toolbars.AddToolbar("Notificacion");
			//			LabelTool lblMensaje = new LabelTool("LBLM");
			//			//TB.Toolbars["Notificación"].Tools.AddTool.Add(lblMensaje);
			//			TB.Toolbars["Notificacion"].Tools.AddTool("LBLM");
			//			TB.Tools["LBLM"].SharedProps.Caption = string.Format("Tiene {0} Cedula(s) Por Revisar", 10); 
			//						
			//			
			//			TB.ShowPopup("Notificacion");
			//			//TB.Text = string.Format("Tiene {0} Cedula(s) Por Revisar", 10);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void uGridRoles_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{			
			//			FuncionesProcedimientos.FormatoGrid(e, "Bono", 2);
			//			FuncionesProcedimientos.FormatoGrid(e, "AnticipoUtilidades", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEAsumeEmpleado", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEAsumeEmpresa", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEGastoSueldo", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Bonos", 2);
			FuncionesProcedimientos.FormatoGrid(e, "BonoMaster", 2);
			FuncionesProcedimientos.FormatoGrid(e, "BonoAdicional", 2);
			FuncionesProcedimientos.FormatoGrid(e, "BonoBeneficios", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Comisiones", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Alimentacion", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Movilizacion", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Spiff", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ComisionReal", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ComprasCredito", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ComprasCrediPoint", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FaltantesCaja", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Multas", 2);
			FuncionesProcedimientos.FormatoGrid(e, "CajaChica", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AnticipoComisiones", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Uniformes", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Refrigerios", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ComisionesPagadas", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AnticipoSueldos", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SeguroMedico", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Farmacia", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Extension", 2);
			FuncionesProcedimientos.FormatoGrid(e, "PrestamoQuirografario", 2);
			FuncionesProcedimientos.FormatoGrid(e, "PrestamoHipotecario", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Enfermedad", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Depreciacion", 2);
			FuncionesProcedimientos.FormatoGrid(e, "OtrosDescuentos", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Faltas", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Faltas", 2);
			FuncionesProcedimientos.FormatoGrid(e, "RetencionJudicial", 2);			
		}

		private void dtpPeriodo_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void Sumatorias(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{	
			//			if (bAnticipoUtilidades)
			//			{
			//				e.Cell.Row.Cells["Bonos"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Bono"].Value), 2) * 50.00m / 100.00m;
			//				e.Cell.Row.Cells["AnticipoUtilidades"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Bono"].Value), 2) * 50.00m / 100.00m;				
			//			}
		
			e.Cell.Row.Cells["SEGastoSueldo"].Value = 
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Sueldo"].Value), 2) -
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["SEAsumeEmpleado"].Value), 2) -
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["SEAsumeEmpresa"].Value), 2);

			e.Cell.Row.Cells["BaseIESS"].Value = 
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Sueldo"].Value), 2) +
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Bonos"].Value), 2)	+	
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["BonoMaster"].Value), 2) + 
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Comisiones"].Value), 2) +
				//Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Alimentacion"].Value), 2) +
				(MenuLatinium.stIdDB != 33 ? Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Movilizacion"].Value), 2) : 0) +
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Spiff"].Value), 2)	+
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["HOrdinarias"].Value), 2) +
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Refrigerios"].Value), 2) +
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["HExtraOrdinarias"].Value), 2);
					
			e.Cell.Row.Cells["DescuentoIESS"].Value = Math.Round((Convert.ToDecimal(e.Cell.Row.Cells["BaseIESS"].Value) * dAportePersonal)/100.00m, 2);

			e.Cell.Row.Cells["AportePatronal"].Value = Math.Round((Convert.ToDecimal(e.Cell.Row.Cells["BaseIESS"].Value) * dAportePatronal)/100.00m, 2);

			e.Cell.Row.Cells["Vacaciones"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["BaseIESS"].Value)/24, 2);

			e.Cell.Row.Cells["DecimoTercerSueldo"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["BaseIESS"].Value)/12, 2);

			e.Cell.Row.Cells["DecimoCuartoSueldo"].Value = ((dSueldoBasico/iDiasTrabajaMes) * (int)e.Cell.Row.Cells["DiasTrabajados"].Value)/12.00m; 
						 
			if (Convert.ToBoolean(e.Cell.Row.Cells["DerechoFondosReserva"].Value))
			{
				if (Convert.ToBoolean(e.Cell.Row.Cells["FondosReservaAcumula"].Value))
				{
					e.Cell.Row.Cells["FRNoAcumula"].Value = 0.00m;
					e.Cell.Row.Cells["FRAcumula"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["BaseIESS"].Value) / 12, 2);					
				}
				else
				{
					e.Cell.Row.Cells["FRAcumula"].Value = 0.00m;
					e.Cell.Row.Cells["FRNoAcumula"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["BaseIESS"].Value) / 12, 2);
				}
			}
			
			e.Cell.Row.Cells["TotalIngresos"].Value = 
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Sueldo"].Value), 2) +
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Bonos"].Value), 2)	+ 
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["BonoMaster"].Value), 2) + 
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["BonoAdicional"].Value), 2)	+
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["BonoBeneficios"].Value), 2)	+
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Comisiones"].Value), 2)	+				
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Alimentacion"].Value), 2) +
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Movilizacion"].Value), 2) +
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Spiff"].Value), 2)	+
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["HOrdinarias"].Value), 2) +
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["HExtraOrdinarias"].Value), 2) + 
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["Refrigerios"].Value), 2) +
				//((MenuLatinium.stIdDB == 33 || MenuLatinium.stIdDB == 1) ? Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["DecimoTercerSueldo"].Value), 2) : 0) +
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["FRNoAcumula"].Value), 2) +
				(((int)e.Cell.Row.Cells["DecimoTercerSueldoAcumulado"].Value == 0) ? Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["DecimoTercerSueldo"].Value), 2) : 0) +
				(((int)e.Cell.Row.Cells["DecimoCuartoSueldoAcumulado"].Value == 0) ? Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["DecimoCuartoSueldo"].Value), 2) : 0) -
				Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["SEAsumeEmpleado"].Value), 2);				

			e.Cell.Row.Cells["PagoIESS"].Value =
				Convert.ToDecimal(e.Cell.Row.Cells["FRAcumula"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["AportePatronal"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["DescuentoIESS"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["PrestamoHipotecario"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["PrestamoQuirografario"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["Extension"].Value);

			e.Cell.Row.Cells["TotalEgresos"].Value = 
				//Convert.ToDecimal(e.Cell.Row.Cells["DescuentoIESS"].Value) + 
				((Convert.ToInt32(e.Cell.Row.Cells["SSN"].Value) == 1) ? Convert.ToDecimal(e.Cell.Row.Cells["DescuentoIESS"].Value) : 0)+
				Convert.ToDecimal(e.Cell.Row.Cells["Extension"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["PrestamoHipotecario"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["PrestamoQuirografario"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["ComprasCredito"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["ComprasCrediPoint"].Value) +				
				Convert.ToDecimal(e.Cell.Row.Cells["FaltantesCaja"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Multas"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["CajaChica"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Uniformes"].Value) +				
				Convert.ToDecimal(e.Cell.Row.Cells["ComisionesPagadas"].Value) +				
				Convert.ToDecimal(e.Cell.Row.Cells["AnticipoSueldos"].Value) +				
				Convert.ToDecimal(e.Cell.Row.Cells["SeguroMedico"].Value) +
				((Convert.ToInt32(e.Cell.Row.Cells["SSN"].Value) == 1) ? Convert.ToDecimal(e.Cell.Row.Cells["Farmacia"].Value) : 0)+
				//Convert.ToDecimal(e.Cell.Row.Cells["Farmacia"].Value) +				
				Convert.ToDecimal(e.Cell.Row.Cells["RetencionJudicial"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["OtrosDescuentos"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["Enfermedad"].Value) +
				Convert.ToDecimal(e.Cell.Row.Cells["Depreciacion"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["BonosPagados"].Value) + 
				Convert.ToDecimal(e.Cell.Row.Cells["Faltas"].Value);				

			e.Cell.Row.Cells["Total"].Value = Convert.ToDecimal(e.Cell.Row.Cells["TotalIngresos"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["TotalEgresos"].Value);		
			
			ActualizaValores("BaseIESS", Convert.ToDecimal(e.Cell.Row.Cells["BaseIESS"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			ActualizaValores("DescuentoIESS", Convert.ToDecimal(e.Cell.Row.Cells["DescuentoIESS"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			ActualizaValores("AportePatronal", Convert.ToDecimal(e.Cell.Row.Cells["AportePatronal"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			ActualizaValores("Vacaciones", Convert.ToDecimal(e.Cell.Row.Cells["Vacaciones"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			ActualizaValores("DecimoTercerSueldo", Convert.ToDecimal(e.Cell.Row.Cells["DecimoTercerSueldo"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			ActualizaValores("DecimoCuartoSueldo", Convert.ToDecimal(e.Cell.Row.Cells["DecimoCuartoSueldo"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			ActualizaValores("FRAcumula", Convert.ToDecimal(e.Cell.Row.Cells["FRAcumula"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			ActualizaValores("FRNoAcumula", Convert.ToDecimal(e.Cell.Row.Cells["FRNoAcumula"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);						
			ActualizaValores("TotalIngresos", Convert.ToDecimal(e.Cell.Row.Cells["TotalIngresos"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			ActualizaValores("PagoIESS", Convert.ToDecimal(e.Cell.Row.Cells["PagoIESS"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			ActualizaValores("TotalEgresos", Convert.ToDecimal(e.Cell.Row.Cells["TotalEgresos"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			ActualizaValores("Total", Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
		}

		private void ActualizaValores(string sCampo, decimal dValor, int idDetalleRolNomina)
		{
			string sSQL = string.Format("Exec RolActualizaValores '{0}', {1}, {2}", sCampo, dValor, idDetalleRolNomina);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void uGridRoles_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{	
			if (e.Cell.Column.ToString() == "Revisado") 
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaActualizaRevisado {0}, {1}, 0", (bool)e.Cell.Row.Cells["Revisado"].Value, (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value));				
			}

			//			if (e.Cell.Column.ToString() == "Bono") 
			//			{
			//				Sumatorias(e);
			//				ActualizaValores("Bono", Convert.ToDecimal(e.Cell.Row.Cells["Bono"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			//			}

			if (e.Cell.Column.ToString() == "SEAsumeEmpleado") 
			{
				Sumatorias(e);
				ActualizaValores("SEAsumeEmpleado", Convert.ToDecimal(e.Cell.Row.Cells["SEAsumeEmpleado"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "SEAsumeEmpresa") 
			{
				Sumatorias(e);
				ActualizaValores("SEAsumeEmpresa", Convert.ToDecimal(e.Cell.Row.Cells["SEAsumeEmpresa"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "SEGastoSueldo") 
			{
				Sumatorias(e);
				ActualizaValores("SEGastoSueldo", Convert.ToDecimal(e.Cell.Row.Cells["SEGastoSueldo"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "Bonos") 
			{
				Sumatorias(e);
				ActualizaValores("Bonos", Convert.ToDecimal(e.Cell.Row.Cells["Bonos"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "BonoMaster") 
			{
				Sumatorias(e);
				ActualizaValores("BonoMaster", Convert.ToDecimal(e.Cell.Row.Cells["BonoMaster"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "BonosPagados") 
			{
				Sumatorias(e);
				ActualizaValores("BonosPagados", Convert.ToDecimal(e.Cell.Row.Cells["BonosPagados"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "AnticipoUtilidades") 
			{
				Sumatorias(e);
				ActualizaValores("AnticipoUtilidades", Convert.ToDecimal(e.Cell.Row.Cells["AnticipoUtilidades"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "BonoAdicional") 
			{
				Sumatorias(e);
				ActualizaValores("BonoAdicional", Convert.ToDecimal(e.Cell.Row.Cells["BonoAdicional"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}
			if (e.Cell.Column.ToString() == "BonoBeneficios") 
			{
				Sumatorias(e);
				ActualizaValores("BonoBeneficios", Convert.ToDecimal(e.Cell.Row.Cells["BonoBeneficios"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "Comisiones") 
			{
				Sumatorias(e);
				ActualizaValores("Comisiones", Convert.ToDecimal(e.Cell.Row.Cells["Comisiones"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "Alimentacion") 
			{
				Sumatorias(e);
				ActualizaValores("Alimentacion", Convert.ToDecimal(e.Cell.Row.Cells["Alimentacion"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "Movilizacion") 
			{
				Sumatorias(e);
				ActualizaValores("Movilizacion", Convert.ToDecimal(e.Cell.Row.Cells["Movilizacion"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "Spiff") 
			{
				Sumatorias(e);						
				ActualizaValores("Spiff", Convert.ToDecimal(e.Cell.Row.Cells["Spiff"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}
			
			if (e.Cell.Column.ToString() == "NHOrdinarias")
			{
				e.Cell.Row.Cells["HOrdinarias"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["VHOrdinaria"].Value)	* Convert.ToDecimal(e.Cell.Row.Cells["NHOrdinarias"].Value), 2);
			}

			if (e.Cell.Column.ToString() == "HOrdinarias")
			{
				Sumatorias(e);
				ActualizaValores("HOrdinarias", Convert.ToDecimal(e.Cell.Row.Cells["HOrdinarias"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "NHExtraOrdinarias")
			{
				e.Cell.Row.Cells["HExtraOrdinarias"].Value = Math.Round(Convert.ToDecimal(e.Cell.Row.Cells["VHExtraOrdinaria"].Value) * Convert.ToDecimal(e.Cell.Row.Cells["NHExtraOrdinarias"].Value), 2);
			}

			if (e.Cell.Column.ToString() == "HExtraOrdinarias")
			{
				Sumatorias(e);
				ActualizaValores("HExtraOrdinarias", Convert.ToDecimal(e.Cell.Row.Cells["HExtraOrdinarias"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "ComisionReal") 
			{
				Sumatorias(e);
				ActualizaValores("ComisionReal", Convert.ToDecimal(e.Cell.Row.Cells["ComisionReal"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "Extension") 
			{
				Sumatorias(e);
				ActualizaValores("Extension", Convert.ToDecimal(e.Cell.Row.Cells["Extension"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "PrestamoQuirografario") 
			{
				Sumatorias(e);
				ActualizaValores("PrestamoQuirografario", Convert.ToDecimal(e.Cell.Row.Cells["PrestamoQuirografario"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "PrestamoHipotecario") 
			{
				Sumatorias(e);
				ActualizaValores("PrestamoHipotecario", Convert.ToDecimal(e.Cell.Row.Cells["PrestamoHipotecario"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "ComprasCredito") 
			{
				Sumatorias(e);
				ActualizaValores("ComprasCredito", Convert.ToDecimal(e.Cell.Row.Cells["ComprasCredito"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "ComprasCrediPoint") 
			{
				Sumatorias(e);
				ActualizaValores("ComprasCrediPoint", Convert.ToDecimal(e.Cell.Row.Cells["ComprasCrediPoint"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "FaltantesCaja") 
			{
				Sumatorias(e);
				ActualizaValores("FaltantesCaja", Convert.ToDecimal(e.Cell.Row.Cells["FaltantesCaja"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "Multas") 
			{
				Sumatorias(e);
				ActualizaValores("Multas", Convert.ToDecimal(e.Cell.Row.Cells["Multas"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "ComisionesPagadas") 
			{
				Sumatorias(e);
				ActualizaValores("ComisionesPagadas", Convert.ToDecimal(e.Cell.Row.Cells["ComisionesPagadas"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "Refrigerios") 
			{
				Sumatorias(e);
				ActualizaValores("Refrigerios", Convert.ToDecimal(e.Cell.Row.Cells["Refrigerios"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "AnticipoComisiones") 
			{
				Sumatorias(e);
				ActualizaValores("AnticipoComisiones", Convert.ToDecimal(e.Cell.Row.Cells["AnticipoComisiones"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			//			if (e.Cell.Column.ToString() == "AnticipoComisiones") 
			//			{
			//				Sumatorias(e);
			//				ActualizaValores("AnticipoComisiones", Convert.ToDecimal(e.Cell.Row.Cells["AnticipoComisiones"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			//			}

			if (e.Cell.Column.ToString() == "AnticipoSueldos") 
			{
				Sumatorias(e);
				ActualizaValores("AnticipoSueldos", Convert.ToDecimal(e.Cell.Row.Cells["AnticipoSueldos"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "SeguroMedico") 
			{
				Sumatorias(e);
				ActualizaValores("SeguroMedico", Convert.ToDecimal(e.Cell.Row.Cells["SeguroMedico"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "Farmacia") 
			{
				Sumatorias(e);
				ActualizaValores("Farmacia", Convert.ToDecimal(e.Cell.Row.Cells["Farmacia"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}
			
			if (e.Cell.Column.ToString() == "Enfermedad") 
			{
				Sumatorias(e);
				ActualizaValores("Enfermedad", Convert.ToDecimal(e.Cell.Row.Cells["Enfermedad"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}
			
			if (e.Cell.Column.ToString() == "Depreciacion") 
			{
				Sumatorias(e);
				ActualizaValores("Depreciacion", Convert.ToDecimal(e.Cell.Row.Cells["Depreciacion"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "OtrosDescuentos") 
			{
				Sumatorias(e);
				ActualizaValores("OtrosDescuentos", Convert.ToDecimal(e.Cell.Row.Cells["OtrosDescuentos"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "Faltas") 
			{
				Sumatorias(e);
				ActualizaValores("Faltas", Convert.ToDecimal(e.Cell.Row.Cells["Faltas"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "Uniformes") 
			{
				Sumatorias(e);
				ActualizaValores("Uniformes", Convert.ToDecimal(e.Cell.Row.Cells["Uniformes"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}

			if (e.Cell.Column.ToString() == "RetencionJudicial") 
			{
				Sumatorias(e);
				ActualizaValores("RetencionJudicial", Convert.ToDecimal(e.Cell.Row.Cells["RetencionJudicial"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}			
			if (e.Cell.Column.ToString() == "CajaChica") 
			{
				Sumatorias(e);
				ActualizaValores("CajaChica", Convert.ToDecimal(e.Cell.Row.Cells["CajaChica"].Value), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value);
			}			
		}

		private void uGridRoles_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridRoles);
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridRoles);
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void frmRolesDePago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void dtpPeriodo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.cmbNECentroCosto.Focus();
		}

		private void CreaNotasDeDebito(int idCliente, string sRuc, string sNombre, string sNotas, decimal dValor, int idProyecto, int idArticulo)
		{			
			DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());

			#region Numeracion
			string sSQLNumero = string.Format("Exec NumeracionLocales 28, 1, 0");			
			string sNumero = Funcion.sEscalarSQL(cdsSeteoF, sSQLNumero);
			#endregion Numeracion

			#region Graba Registro
			string sSQLEncabezado = string.Format("Exec GrabaNotaDeDebito {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', {8}, '{9}', {10}, {11}",
				0, 28, 0, idCliente, sRuc, sNombre, dtFecha.ToString("yyyyMMdd"), sNumero, 1, sNotas, dValor, dValor);
			int idCompra = Funcion.iEscalarSQL(cdsSeteoF, sSQLEncabezado);
						
			string sSQLDetalle = string.Format("Exec GrabaDetalleNotaDeDebito {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}",
				0, idCompra, idArticulo, "", "", 1, dValor, dValor, idProyecto);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			#endregion Graba Registro

			#region Crea Asiento
			string sSQLAsiento = string.Format("Exec CreaAsiNotaDevCompra {0}", idCompra);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento);
			#endregion Crea Asiento

			#region Autorizacion de Pagos
			string sSQLAP = string.Format("Exec CreaAutorizacionesDePago 15, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', '{4}', 0", 
				idCompra, dtFecha.ToString("yyyyMMdd"), dtFecha.ToString("yyyyMMdd"), dValor, sNotas);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLAP);			
			#endregion Autorizacion de Pagos
		}

		private void CreaPagos(int idAnticipo, int idCliente, decimal dValor, int iTipo, int idDetalleRolNomina, int idCompra)
		{
			#region Guarda Anticipo
			DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());;
			
			string sNumero = "";
			
			if (idAnticipo == 0) sNumero = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 24, {0}, 0", 1));

			string sSQLGraba = string.Format("Exec GuardaAnticipos {0}, {1}, '{2}', {3}, '{4}', {5}, {6}, {7}, '{8}', {9}, {10}, {11}, '{12}', {13}, 0", 
				idAnticipo, 0, dtFecha.ToString("yyyyMMdd"), 1, sNumero, idCliente, 18, 0, "ABONO FACTURA DE CREDITO PENDIENTE", 1, dValor, iTipo, "", idCompra);
			idAnticipo = Funcion.iEscalarSQL(cdsSeteoF, sSQLGraba);			

			#region Guarda Cobro
			int iOrigen = 0;
			if (iTipo == 3) iOrigen = 4;

			string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, 0", 
				0, idAnticipo, iOrigen, 18, dValor, 0, 0, "", "", dtFecha.ToString("yyyyMMdd"),	"", "", 0, "", "", 0, 0, 0);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro);
			#endregion Guarda Cobro

			#region Asiento
			if (iTipo == 3)
			{
				string sSQLAsiento = string.Format("Exec AsientoCXCAbono {0}, {1}", idAnticipo, 0);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);

				//int idCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec NominaRetornaIdCompraFacturas {0}", idCliente));
				
				string sSQLInsPag = string.Format("Insert Into AnticiposCompras (idAnticipo, idCompra, Valor, Fecha, idTipoFactura) Values ({0}, {1}, {2}, '{3}', {4})", 
					idAnticipo, idCompra, dValor, dtFecha.ToString("yyyyMMdd"), 1);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLInsPag);
								
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoFacturas {0}, 1, 6", idCompra));
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleRolNomina Set idAnticipoCredito = {0} Where idDetalleRolNomina = {1} Update Anticipos Set Estado = 3, Saldo = 0 Where idAnticipo = {0}", idAnticipo, idDetalleRolNomina));
			}
			else if (iTipo == 12)
			{
				string sSQLAsiento = string.Format("Exec CreaAsientoAnticipo {0}, {1}, {2}, {3}", 0, idAnticipo, 12, 1);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);				

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleRolNomina Set idAnticipoCrediPoint = {0} Where idDetalleRolNomina = {1}", idAnticipo, idDetalleRolNomina));
			}
			#endregion Asiento
			#endregion Guarda Anticipo
		}

		#region Genera Archivos Comentados
		/*
				private void Internacional()
				{
					region Exporta
					try
					{
						SaveFileDialog sArchivo = new SaveFileDialog();

						sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
						sArchivo.FileName = sArchivo.FileName;
						sArchivo.OverwritePrompt = true;
						sArchivo.Title = "Guardar Lista Acreditaciones Banco Internacional";		
                                
						if (sArchivo.ShowDialog() == DialogResult.OK)
						{
							using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
							{						
								ion Conexion
								SqlCommand oCmdActualiza = oConexion.CreateCommand();
								oCmdActualiza.Connection = oConexion;
								oCmdActualiza.CommandTimeout = 0;
								string sSQL = string.Format("Exec RolListaPersonalInternacional {0}", (int)this.txtIdRolNomina.Value);
								oCmdActualiza.CommandText = sSQL;
								SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
								DataTable oDTDetalle = new DataTable();
								oDADetalle.Fill(oDTDetalle);	
								region Conexion

								ion Pagos
								using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
								{						
									int i = 1;
									foreach (DataRow dr in oDTDetalle.Rows)
									{							
										string iLinea = dr["Pago"].ToString() + "\t" + i.ToString() + "\t" + dr["Moneda"].ToString() + "\t" + dr["Valor"].ToString() + "\t"
											+ dr["Titulo"].ToString() + "\t" + dr["TipoDeCuenta"].ToString() + "\t" + dr["Cuenta"].ToString() + "\t"
											+ dr["Mensaje"].ToString() + "\t" + dr["TipoDoc"].ToString() + "\t" + dr["Cedula"].ToString() + "\t" + dr["Nombre"].ToString();

										file.WriteLine(iLinea);

										i++;
									}					
								}
								region Pagos					
							}
						}
					}
					catch (Exception exc)
					{
						MessageBox.Show(exc.Message);
					}
					endregion Exporta
				}

				private void Loja()
				{			
					SaveFileDialog sArchivo = new SaveFileDialog();

					sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";
					sArchivo.FileName = sArchivo.FileName;
					sArchivo.OverwritePrompt = true;
					sArchivo.Title = "Guardar Lista Acreditaciones Banco De Loja";
                                
					if (sArchivo.ShowDialog() == DialogResult.OK)
					{
						this.uGridLoja.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RolListaPersonalLoja {0}", (int)this.txtIdRolNomina.Value));

						ultraGridExcelExporter1.Export(this.uGridLoja, sArchivo.FileName);
						System.Diagnostics.Process.Start(sArchivo.FileName);
					}
				}
		*/
		#endregion Genera Archivos Comentados

		private void btnContabilizar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de CONTABILIZAR EL ROL DEL PERIODO DE '{0}'", 
				Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				this.Cursor = Cursors.WaitCursor;
				
				DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());
	
				string sSQL = string.Format("Exec AsientoEncabezadoRolNomina {0}, {1}", (int)this.txtIdAsiento.Value, (int)this.txtIdRolNomina.Value);
				this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);				
					
				#region Farmacias - Abonos Facturas
				int iContadorCrediPoint = 0;
				decimal dValorFarmacias = 0.00m;
				decimal dPagoIESS = 0.00m;
				decimal dCredito = 0.00m;
		
				//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridRoles.Rows.All)
				//					dValorFarmacias = dValorFarmacias + Convert.ToDecimal(dr.Cells["Farmacia"].Value);	
				//
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridRoles.Rows.All)
				{
					//					dValorFarmacias = dValorFarmacias + Convert.ToDecimal(dr.Cells["Farmacia"].Value);	
					//
					dPagoIESS = dPagoIESS + Convert.ToDecimal(dr.Cells["PagoIESS"].Value);
					//					dCredito = dCredito + Convert.ToDecimal(dr.Cells["ComprasCredito"].Value);
					//
					//					#region Credito - CrediPoint
					//					if (Convert.ToDecimal(dr.Cells["ComprasCredito"].Value) > 0.00m)
					//					{
					//						int idCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCliente From Compra Where idCompra = {0}", Convert.ToInt32(dr.Cells["idCompraC"].Value)));//Select idCliente From Cliente Where Proveedor = 0 And Ruc = '{0}'", dr.Cells["Cedula"].Value.ToString()));
					//					
					//						CreaPagos(0, idCliente, Convert.ToDecimal(dr.Cells["ComprasCredito"].Value), 3, Convert.ToInt32(dr.Cells["idDetalleRolNomina"].Value), Convert.ToInt32(dr.Cells["idCompraC"].Value));
					//					}
					//				
					////					if (Convert.ToDecimal(dr.Cells["ComprasCrediPoint"].Value) > 0.00m) 
					////					{
					////						iContadorCrediPoint++;
					////
					////						int idCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCliente From Cliente Where Proveedor = 0 And Ruc = '{0}'", dr.Cells["Cedula"].Value.ToString()));
					////
					////						CreaPagos(0, idCliente, Convert.ToDecimal(dr.Cells["ComprasCrediPoint"].Value), 12, Convert.ToInt32(dr.Cells["idDetalleRolNomina"].Value));
					////					}					
					//					#endregion Credito - CrediPoint
				}				
				#endregion Farmacias - Abonos Facturas						

				//				if (dCredito > 0.00m)
				//					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RolCruzaPagosCreditoCorporativo {0}", (int)this.txtIdRolNomina.Value));

				#region Autorización de Pagos
				string sSQLAP = string.Format("Exec CreaAutorizacionesDePago 19, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', 'PAGO IESS - ROL DEL MES DE {4}', 0", 
					(int)this.txtIdRolNomina.Value, dtFecha.AddDays(15).ToString("yyyyMMdd"), dtFecha.AddDays(15).ToString("yyyyMMdd"), dPagoIESS, dtFecha.ToString("MMMM").ToUpper());
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAP, true);
				#endregion Autorización de Pagos

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaActualizaEstadoOtrosDescuentos '{0}'", dtFecha.ToString("yyyyMMdd")));

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaActualizaEstadoUniformes {0}", (int)this.txtIdRolNomina.Value));
				
				#region Utilidades
				//				string sSQLUtilidades = string.Format("Exec GuardaEncabezadoAnticipoUtilidades  '{0}'",  Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"));
				//				this.txtIdAntipoUtilidades.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLUtilidades);
				//	
				//				string sSQLDetalleUtilidades = string.Format("Exec GuardaDetalleAnticipoUtilidades {0},{1}",(int)this.txtIdRolNomina.Value, (int)this.txtIdAntipoUtilidades.Value, MenuLatinium.bComisionQuincena);
				//				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalleUtilidades);
				#endregion Utilidades

				Estado = 5;

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaActualizaEstadoNomina {0}",  (int)this.txtIdRolNomina.Value));

				this.lblEstado.Text = "CONTABILIZADO";
								
				MessageBox.Show(string.Format("Rol de Pagos del Periodo '{0}' Contabilizado Correctamente", Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				
				this.btnContabilizar.Enabled = false;
				this.btnUtilidades.Enabled =true;
				this.btnImprimir.Enabled = true;	

				#region Impresion
				string strFiltro = "{Asiento.idAsiento} = " + this.txtIdAsiento.Value.ToString();				
				Reporte Reporte = new Reporte("Asientos.rpt", strFiltro);
				Reporte.MdiParent = this.MdiParent;				
				Reporte.Show();
				#endregion Impresion

				this.Cursor = Cursors.Default;        
			}
		}

		private void btnAcreditacion_Click(object sender, System.EventArgs e)
		{
			this.mnuAcreditacion.Show(this.btnAcreditacion, new Point(1, 25));			
		}

		private void btnAutorizar_Click(object sender, System.EventArgs e)
		{
			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridRoles.Rows.All)
			{	
				if (!(bool)dr.Cells["Revisado"].Value) iCont++;
			}

			if (iCont > 0)
			{
				MessageBox.Show(string.Format("Existen {0} Registros Sin Revisar", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de AUTORIZAR EL ROL DEL PERIODO DE '{0}'", 
				Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaAutorizaRol {0}, 3", (int)this.txtIdRolNomina.Value));

				FuncionesProcedimientos.EstadoGrids(false, this.uGridRoles);

				this.btnAcreditacion.Enabled = true;				
				this.btnContabilizar.Enabled = true;
				this.btnAutorizar.Enabled = false;

				Estado = 3;

				this.lblEstado.Text = "AUTORIZADO";

				this.dtpPeriodo_ValueChanged(sender, e);
			}
		}

		private void cmbNECentroCosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.cmbNEDepartamento.Focus();
		}

		private void cmbNEDepartamento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.txtBuscar.Focus();
		}

		private void txtBuscar_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{	
			this.mnuAcciones.Show(this.btnActualizar, new Point(1, 25));
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de Imprimir los Roles de Pago individuales del periodo '{0}',\n\nPrepare {1} hojas para Iniciar la Impresión", 
				Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper(), this.uGridRoles.Rows.Count), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				#region Impresion Pantalla Bloqueada
				//								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridRoles.Rows.All)
				//								{
				//									ParameterFields paramFields = new ParameterFields ();
				//																					
				//									ParameterField paramFieldCompra = new ParameterField ();
				//									ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				//									paramFieldCompra.ParameterFieldName = "@idRolNomina";
				//									ValIdCompra.Value = (int)this.txtIdRolNomina.Value;
				//									paramFieldCompra.CurrentValues.Add (ValIdCompra);
				//									paramFields.Add (paramFieldCompra);
				//																	
				//									ParameterField paramFieldFecha = new ParameterField ();
				//									ParameterDiscreteValue ValIdFecha = new ParameterDiscreteValue ();
				//									paramFieldFecha.ParameterFieldName = "@idNomina";
				//									ValIdFecha.Value = Convert.ToInt32(dr.Cells["idNomina"].Value);
				//									paramFieldFecha.CurrentValues.Add (ValIdFecha);
				//									paramFields.Add (paramFieldFecha);
				//																					
				//									Reporte miRep = new Reporte("RolIndividual.rpt", "");
				//									miRep.MdiParent = this.MdiParent;
				//									miRep.Titulo("Rol Individual");
				//									miRep.crVista.ParameterFieldInfo = paramFields;
				//									miRep.Show();
				//								}					
				//							
				//			return;
				#endregion Impresion Pantalla Bloqueada

				using (frmImprimeRolesDePago IRP = new frmImprimeRolesDePago(this.uGridRoles.Rows.Count))
				{
					if (DialogResult.OK == IRP.ShowDialog())
					{
						int iCont = 0;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridRoles.Rows.All)
						{	
							//							if (!(bool)dr.Cells["Impreso"].Value)
							//							{
							if ((int)dr.Cells["Numero"].Value >= (int)IRP.txtDesde.Value && (int)dr.Cells["Numero"].Value <= (int)IRP.txtHasta.Value)
							{
								if(MenuLatinium.stIdDB == 1 || MenuLatinium.stIdDB == 33)
								{
									#region Impresion
									string sReporte = "RolesIndividuales.rpt";
									if (MenuLatinium.stDirFacturacion == "LOTESERVICIO") sReporte = "RolIndividualPeony.rpt";
									ReportDocument Report1 = new ReportDocument();
									Report1.FileName = MenuLatinium.stDirReportes + sReporte;//"RolesIndividuales.rpt";
								
									ConnectionInfo crConnectionInfo = new ConnectionInfo (); 
								{
									string strBaseDato = MenuLatinium.stDirNumero;
									string strUsuario = MenuLatinium.stDirConta;
									string strServidor = MenuLatinium.stDirServidor;
									string strImpresora = MenuLatinium.ImpresoraNombre;

									crConnectionInfo.ServerName = strServidor;
									crConnectionInfo.DatabaseName = strBaseDato; 
									crConnectionInfo.UserID = "infoelect"; 
									crConnectionInfo.Password = "Bl45o6$9"; 
								}
			
									TableLogOnInfo crTableLogOnInfo; 
									Database crDatabase = Report1.Database; 
									Tables crTables = crDatabase.Tables; 
									Table crTable;

									for (int i=0; i< crTables.Count; i++)
									{
									{
										crTable = crTables [i]; 
										crTableLogOnInfo = crTable.LogOnInfo; 
										crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
										crTable.ApplyLogOnInfo(crTableLogOnInfo);
									}
									}
					
									Report1.PrintOptions.PrinterName = MenuLatinium.ImpresoraNombre;
									Report1.SetParameterValue("@idRolNomina", (int)this.txtIdRolNomina.Value);
									Report1.SetParameterValue("@idNomina", Convert.ToInt32(dr.Cells["idNomina"].Value));
					
									Report1.PrintToPrinter(1, false, 0, 0);	
									#endregion Impresion	
								}
								else
								{									
									#region Impresion Otras Empresas
									ParameterFields paramFields = new ParameterFields ();
																														
									ParameterField paramFieldCompra = new ParameterField ();
									ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
									paramFieldCompra.ParameterFieldName = "@idRolNomina";
									ValIdCompra.Value = (int)this.txtIdRolNomina.Value;
									paramFieldCompra.CurrentValues.Add (ValIdCompra);
									paramFields.Add (paramFieldCompra);
																										
									ParameterField paramFieldFecha = new ParameterField ();
									ParameterDiscreteValue ValIdFecha = new ParameterDiscreteValue ();
									paramFieldFecha.ParameterFieldName = "@idNomina";
									ValIdFecha.Value = Convert.ToInt32(dr.Cells["idNomina"].Value);
									paramFieldFecha.CurrentValues.Add (ValIdFecha);
									paramFields.Add (paramFieldFecha);
																														
									Reporte miRep = new Reporte("RolesIndividuales.rpt", "");
									miRep.MdiParent = this.MdiParent;
									miRep.Titulo("Rol Individual");
									miRep.crVista.ParameterFieldInfo = paramFields;
									miRep.Show();
									#endregion Impresion Otras Empresas
								}
								iCont++;
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleRolNomina Set Impreso = 1 Where idDetalleRolNomina = {0}", Convert.ToInt32(dr.Cells["idDetalleRolNomina"].Value)));
							}
						}
					}
				}
			}
		}

		private void mnuSAPQ_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;                                  

				int iCont = 0;

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CEDULA, SUM(VALOR) As Valor FROM [QUIROGRAFARIOS$] GROUP BY CEDULA", "QUIROGRAFARIOS").Tables[0].Rows)
				{
					decimal dValor = 0.00m;
					string sCedula = "";
					
					if (row["Cedula"] != System.DBNull.Value) sCedula = row["Cedula"].ToString();
					if (row["Valor"] != System.DBNull.Value) dValor = Convert.ToDecimal(row["Valor"]);
					
					if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
					
					string sSQL = string.Format("Exec NominaActualizaDescuentos {0}, 1, '{1}', {2}", (int)this.txtIdRolNomina.Value, sCedula, dValor);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					

					iCont++;
				}				

				MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));
			
				this.btnVer_Click(sender, e);
				this.Cursor = Cursors.Default;
			}
		}

		private void btnSubirArchivos_Click(object sender, System.EventArgs e)
		{
			this.mnuSubeArchivos.Show(this.btnSubirArchivos, new Point(1, 25));
		}

		private void mnuSAPH_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;                                  

				int iCont = 0;

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CEDULA, SUM(VALOR) As Valor FROM [HIPOTECARIOS$] GROUP BY CEDULA", "HIPOTECARIOS").Tables[0].Rows)
				{
					decimal dValor = 0.00m;
					string sCedula = "";
					
					if (row["Cedula"] != System.DBNull.Value) sCedula = row["Cedula"].ToString();
					if (row["Valor"] != System.DBNull.Value) dValor = Convert.ToDecimal(row["Valor"]);
					
					if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
					
					string sSQL = string.Format("Exec NominaActualizaDescuentos {0}, 2, '{1}', {2}", (int)this.txtIdRolNomina.Value, sCedula, dValor);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					

					iCont++;
				}				

				MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));
			
				this.btnVer_Click(sender, e);

				this.Cursor = Cursors.Default;
			}
		}

		private void mnuSAF_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;                                  

				int iCont = 0;

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CEDULA, SUM(VALOR) As Valor FROM [FARMACIA$] GROUP BY CEDULA", "FARMACIA").Tables[0].Rows)
				{
					decimal dValor = 0.00m;
					string sCedula = "";
					
					if (row["Cedula"] != System.DBNull.Value) sCedula = row["Cedula"].ToString();
					if (row["Valor"] != System.DBNull.Value) dValor = Convert.ToDecimal(row["Valor"]);
					
					if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
					
					string sSQL = string.Format("Exec NominaActualizaDescuentos {0}, 3, '{1}', {2}", (int)this.txtIdRolNomina.Value, sCedula, dValor);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					

					iCont++;
				}				

				MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));
			
				this.btnVer_Click(sender, e);

				this.Cursor = Cursors.Default;
			}
		}

		private void btnSeleccionar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridRoles.Rows.All)
			{	
				dr.Cells["Revisado"].Value = !(bool)dr.Cells["Revisado"].Value;
			}
		}

		private void uGridRoles_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnCredito")
			{		
				using (frmCreditoEmpleado Busqueda = new frmCreditoEmpleado((int)e.Cell.Row.Cells["idNomina"].Value, (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value, Estado))
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{
						e.Cell.Row.Cells["ComprasCredito"].Value = Busqueda.txtTotal.Value;
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

				using (frmNominaObservaciones Observaciones = new frmNominaObservaciones(1, e.Cell.Row.Cells["ObservacionesFC"].Value.ToString().Trim()))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleRolNomina Set ObservacionesFC = '{0}' Where idDetalleRolNomina = {1}", 
							Observaciones.txtObservaciones.Text.ToString().Trim(), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value));

						e.Cell.Row.Cells["ObservacionesFC"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
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
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleRolNomina Set ObservacionesMultas = '{0}' Where idDetalleRolNomina = {1}", 
							Observaciones.txtObservaciones.Text.ToString().Trim(), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value));

						e.Cell.Row.Cells["ObservacionesMultas"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
					}
				}
			}
			#endregion Multas

			#region Anticipos
			if (e.Cell.Column.ToString() == "btnAnticipos")
			{	
				if (Convert.ToDecimal(e.Cell.Row.Cells["AnticipoSueldos"].Value) == 0)
				{
					MessageBox.Show("El valor de Anticipos es CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}

				using (frmNominaObservaciones Observaciones = new frmNominaObservaciones(4, e.Cell.Row.Cells["ObservacionesAnticipos"].Value.ToString().Trim()))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleRolNomina Set ObservacionesAnticipos = '{0}' Where idDetalleRolNomina = {1}", 
							Observaciones.txtObservaciones.Text.ToString().Trim(), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value));

						e.Cell.Row.Cells["ObservacionesAnticipos"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
					}
				}
			}
			#endregion Anticipos

			#region Otros Descuentos
			if (e.Cell.Column.ToString() == "btnOtros")
			{	
				if (Convert.ToDecimal(e.Cell.Row.Cells["OtrosDescuentos"].Value) == 0)
				{
					MessageBox.Show("El valor de Otros descuentos es CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}

				using (frmNominaObservaciones Observaciones = new frmNominaObservaciones(3, e.Cell.Row.Cells["ObservacionesOtros"].Value.ToString().Trim()))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DetalleRolNomina Set ObservacionesOtros = '{0}' Where idDetalleRolNomina = {1}", 
							Observaciones.txtObservaciones.Text.ToString().Trim(), (int)e.Cell.Row.Cells["idDetalleRolNomina"].Value));

						e.Cell.Row.Cells["ObservacionesOtros"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
					}
				}
			}
			#endregion Otros Descuentos
		}

		private void mnuAcreditacionTransferencia_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());

			using (frmRolAcreditacion RAC = new frmRolAcreditacion((int)this.txtIdRolNomina.Value, 0, dtFecha, 0))
			{
				if (DialogResult.OK == RAC.ShowDialog())
				{
		
				}

				this.dtpPeriodo_ValueChanged(sender, e);
			}		
		}

		private void mnuAcreditacionCheques_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());

			using (frmRolAutorizacionPagoCheques RAPC = new frmRolAutorizacionPagoCheques((int)this.txtIdRolNomina.Value, dtFecha, 0))
			{
				if (DialogResult.OK == RAPC.ShowDialog())
				{
		
				}

				this.dtpPeriodo_ValueChanged(sender, e);
			}
		}

		private void mnuAcreditacionRJ_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = Convert.ToDateTime(DateTime.DaysInMonth(Convert.ToDateTime(this.dtpPeriodo.Value).Year, Convert.ToDateTime(this.dtpPeriodo.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtpPeriodo.Value).Year.ToString());

			using (frmRolRetencionJudicial RRJ = new frmRolRetencionJudicial((int)this.txtIdRolNomina.Value, dtFecha))
			{
				if (DialogResult.OK == RRJ.ShowDialog())
				{
		
				}

				this.dtpPeriodo_ValueChanged(sender, e);
			}
		}

		private void btnRevisado_Click(object sender, System.EventArgs e)
		{
			int iContH = 0;
			int iContQ = 0;
			int iContF = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridRoles.Rows.All)
			{	
				if (Convert.ToDecimal(dr.Cells["PrestamoQuirografario"].Value) > 0.00m) iContQ++;
				if (Convert.ToDecimal(dr.Cells["PrestamoHipotecario"].Value) > 0.00m) iContH++;
				if (Convert.ToDecimal(dr.Cells["Farmacia"].Value) > 0.00m) iContF++;
			}

			if (iContH == 0 && DialogResult.No == MessageBox.Show("No hay empleados con Prestamos Hipotecarios\n\n¿Esta seguro de Continuar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			if (iContQ == 0 && DialogResult.No == MessageBox.Show("No hay empleados con Prestamos Quirografarios\n\n¿Esta seguro de Continuar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			if (iContF == 0 && DialogResult.No == MessageBox.Show("No hay empleados con Consumos de Farmacia\n\n¿Esta seguro de Continuar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			
			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de PONER EN REVISION EL ROL DEL PERIODO DE '{0}'", 
				Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaAutorizaRol {0}, 2", (int)this.txtIdRolNomina.Value));

				FuncionesProcedimientos.EstadoGrids(false, this.uGridRoles);

				this.btnSubirArchivos.Enabled = false;
				this.btnRevisado.Enabled = false;
				this.btnAutorizar.Enabled = true;
				
				Estado = 2;

				this.lblEstado.Text = "EN REVISIÓN";

				this.dtpPeriodo_ValueChanged(sender, e);
			}
		}

		private void ultraGridExcelExporter1_BeginExport(object sender, Infragistics.Win.UltraWinGrid.ExcelExport.BeginExportEventArgs e)
		{
		
		}

		private void saveFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void mnuAccionesAgregar_Click(object sender, System.EventArgs e)
		{
			string sSQLC = string.Format("Exec NominaCuentaPesonalFaltanteRol '{0}', {1}", 
				Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"), (int)this.txtIdRolNomina.Value);
			int iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLC, true);		
			
			if (iCont == 0)
			{
				MessageBox.Show("No existe personal nuevo para agregar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.OK == MessageBox.Show(string.Format("Agregar {0} empleados al rol del periodo '{1}'", iCont,
				Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				this.Cursor = Cursors.WaitCursor;
						
				string sSQL = string.Format("Exec NominaAgregaPersonalRol '{0}', {1}", 
					Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"), (int)this.txtIdRolNomina.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				
				#region retorna Porcentajes de Bonos HE HEO
				decimal dPorcentajeBonos = 0.00m; 
				decimal dPorcentajeHorasExtras = 0.00m;
				decimal dPorcentajeHorasOrdinarias	= 0.00m;

				SqlDataReader drPorc = Funcion.rEscalarSQL(cdsSeteoF, 
					string.Format("Select PorcentajeBonos, PorcentajeHorasExtras, PorcentajeHorasOrdinarias From RolPorcentajes Where idRolNomina = {0}", 
					(int)this.txtIdRolNomina.Value), true);
				drPorc.Read();
				if (drPorc.HasRows)
				{
					dPorcentajeBonos = drPorc.GetDecimal(0);
					dPorcentajeHorasExtras = drPorc.GetDecimal(1);
					dPorcentajeHorasOrdinarias = drPorc.GetDecimal(2);
				}
				drPorc.Close();

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaAgregaCalculaBonosHorasExtras {0}, {1}, {2}, {3}", 
					(int)this.txtIdRolNomina.Value, dPorcentajeBonos, dPorcentajeHorasExtras, dPorcentajeHorasOrdinarias));	
				#endregion retorna Porcentajes de Bonos HE HEO

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));

				MessageBox.Show(string.Format("Rol del Periodo de '{0}' Actualizado Correctamente", Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnVer_Click(sender, e);

				this.Cursor = Cursors.Default;
			}
		}

		private void mnuAccionesActualizar_Click(object sender, System.EventArgs e)
		{			
			if (!Validacion.vbValidaSelGrids(this.uGridRoles, "Revisado", "Seleccione al menos un registro para actualizar")) return;

			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta seguro de ACTUALIZAR LA INFORMACIÓN DEL PERSONAL SELECCIONADO DEL PERIODO DE '{0}'", 
				Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				this.Cursor = Cursors.WaitCursor;
						
				string sSQL = string.Format("Exec NominaActualizaInformacionPersonalRol '{0}', {1}", 
					Convert.ToDateTime(this.dtpPeriodo.Value).ToString("yyyyMMdd"), (int)this.txtIdRolNomina.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				
				#region retorna Porcentajes de Bonos HE HEO
				decimal dPorcentajeBonos = 0.00m; 
				decimal dPorcentajeHorasExtras = 0.00m;
				decimal dPorcentajeHorasOrdinarias	= 0.00m;

				SqlDataReader drPorc = Funcion.rEscalarSQL(cdsSeteoF, 
					string.Format("Select PorcentajeBonos, PorcentajeHorasExtras, PorcentajeHorasOrdinarias From RolPorcentajes Where idRolNomina = {0}", 
					(int)this.txtIdRolNomina.Value), true);
				drPorc.Read();
				if (drPorc.HasRows)
				{
					dPorcentajeBonos = drPorc.GetDecimal(0);
					dPorcentajeHorasExtras = drPorc.GetDecimal(1);
					dPorcentajeHorasOrdinarias = drPorc.GetDecimal(2);
				}
				drPorc.Close();

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaActualizaCalculaBonosHorasExtras {0}, {1}, {2}, {3}", 
					(int)this.txtIdRolNomina.Value, dPorcentajeBonos, dPorcentajeHorasExtras, dPorcentajeHorasOrdinarias));	
				#endregion retorna Porcentajes de Bonos HE HEO

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));

				MessageBox.Show(string.Format("Rol del Periodo de '{0}' Actualizado Correctamente", Convert.ToDateTime(this.dtpPeriodo.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnVer_Click(sender, e);

				this.Cursor = Cursors.Default;
			}
		}

		private void uGridRoles_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (Estado == 1)
			{
				e.DisplayPromptMsg = false;
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta Seguro de Borrar este registro?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				{
					for (int i=0; i<e.Rows.Length; i++)
					{
						string sSQL = string.Format("Delete From DetalleRolNomina Where idDetalleRolNomina = {0} And idRolNomina = {1}", 
							(int)e.Rows[i].Cells["idDetalleRolNomina"].Value, (int)e.Rows[i].Cells["idRolNomina"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
					}
				}
				else e.Cancel = true;
			}
			else e.Cancel = true;
		}

		private void btnUtilidades_Click(object sender, System.EventArgs e)
		{
			using (frmRolAnticipoUtilidades AntUti = new frmRolAnticipoUtilidades((int)this.txtIdAntipoUtilidades.Value,6))
			{
				if (DialogResult.OK == AntUti.ShowDialog())
				{
		
				}

				this.dtpPeriodo_ValueChanged(sender, e);
			}
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;                                  

				int iCont = 0;

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CEDULA, SUM(VALOR) As Valor FROM [CAJACHICA$] GROUP BY CEDULA", "CAJACHICA").Tables[0].Rows)
				{
					decimal dValor = 0.00m;
					string sCedula = "";
					
					if (row["Cedula"] != System.DBNull.Value) sCedula = row["Cedula"].ToString();
					if (row["Valor"] != System.DBNull.Value) dValor = Convert.ToDecimal(row["Valor"]);
					
					if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
					
					string sSQL = string.Format("Exec NominaActualizaDescuentos {0}, 4, '{1}', {2}", (int)this.txtIdRolNomina.Value, sCedula, dValor);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					

					iCont++;
				}				

				MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));
			
				this.btnVer_Click(sender, e);
				this.Cursor = Cursors.Default;
			}
		}

		private void menuItem2_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;                                  

				int iCont = 0;

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CEDULA, SUM(VALOR) As Valor FROM [DEPRECIACION$] GROUP BY CEDULA", "DEPRECIACION").Tables[0].Rows)
				{
					decimal dValor = 0.00m;
					string sCedula = "";
					
					if (row["Cedula"] != System.DBNull.Value) sCedula = row["Cedula"].ToString();
					if (row["Valor"] != System.DBNull.Value) dValor = Convert.ToDecimal(row["Valor"]);
					
					if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
					
					string sSQL = string.Format("Exec NominaActualizaDescuentos {0}, 6, '{1}', {2}", (int)this.txtIdRolNomina.Value, sCedula, dValor);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					

					iCont++;
				}				

				MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));
			
				this.btnVer_Click(sender, e);
				this.Cursor = Cursors.Default;
			}
		}

		private void menuItem3_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;                                  

				int iCont = 0;

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CEDULA, SUM(VALOR) As Valor FROM [MULTAS$] GROUP BY CEDULA", "MULTAS").Tables[0].Rows)
				{
					decimal dValor = 0.00m;
					string sCedula = "";
					
					if (row["Cedula"] != System.DBNull.Value) sCedula = row["Cedula"].ToString();
					if (row["Valor"] != System.DBNull.Value) dValor = Convert.ToDecimal(row["Valor"]);
					
					if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
					
					string sSQL = string.Format("Exec NominaActualizaDescuentos {0}, 5, '{1}', {2}", (int)this.txtIdRolNomina.Value, sCedula, dValor);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					

					iCont++;
				}				

				MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));
			
				this.btnVer_Click(sender, e);
				this.Cursor = Cursors.Default;
			}
		
		}

		private void menuItem4_Click(object sender, System.EventArgs e)
		{
		
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;                                  

				int iCont = 0;

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CEDULA, SUM(VALOR) As Valor FROM [SEGUROS$] GROUP BY CEDULA", "SEGUROS").Tables[0].Rows)
				{
					decimal dValor = 0.00m;
					string sCedula = "";
					
					if (row["Cedula"] != System.DBNull.Value) sCedula = row["Cedula"].ToString();
					if (row["Valor"] != System.DBNull.Value) dValor = Convert.ToDecimal(row["Valor"]);
					
					if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
					
					string sSQL = string.Format("Exec NominaActualizaDescuentos {0}, 7, '{1}', {2}", (int)this.txtIdRolNomina.Value, sCedula, dValor);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					

					iCont++;
				}				

				MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));
			
				this.btnVer_Click(sender, e);
				this.Cursor = Cursors.Default;
			}
		}

		private void menuItem5_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;                                  

				int iCont = 0;

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CEDULA, SUM(VALOR) As Valor FROM [OTROS$] GROUP BY CEDULA", "OTROS").Tables[0].Rows)
				{
					decimal dValor = 0.00m;
					string sCedula = "";
					
					if (row["Cedula"] != System.DBNull.Value) sCedula = row["Cedula"].ToString();
					if (row["Valor"] != System.DBNull.Value) dValor = Convert.ToDecimal(row["Valor"]);
					
					if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
					
					string sSQL = string.Format("Exec NominaActualizaDescuentos {0}, 8, '{1}', {2}", (int)this.txtIdRolNomina.Value, sCedula, dValor);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					

					iCont++;
				}				

				MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));
			
				this.btnVer_Click(sender, e);
				this.Cursor = Cursors.Default;
			}
		
		}

		private void menuItem6_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;                                  

				int iCont = 0;

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CEDULA, SUM(VALOR) As Valor FROM [CREDITO$] GROUP BY CEDULA", "CREDITO").Tables[0].Rows)
				{
					decimal dValor = 0.00m;
					string sCedula = "";
					
					if (row["Cedula"] != System.DBNull.Value) sCedula = row["Cedula"].ToString();
					if (row["Valor"] != System.DBNull.Value) dValor = Convert.ToDecimal(row["Valor"]);
					
					if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
					
					string sSQL = string.Format("Exec NominaActualizaDescuentos {0}, 9, '{1}', {2}", (int)this.txtIdRolNomina.Value, sCedula, dValor);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					

					iCont++;
				}				

				MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));
			
				this.btnVer_Click(sender, e);
				this.Cursor = Cursors.Default;
			}
		}

		private void menuItem7_Click(object sender, System.EventArgs e)
		{
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;                                  

				int iCont = 0;

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CEDULA, SUM(VALOR) As Valor FROM [CREDIPOINT$] GROUP BY CEDULA", "CREDIPOINT").Tables[0].Rows)
				{
					decimal dValor = 0.00m;
					string sCedula = "";
					
					if (row["Cedula"] != System.DBNull.Value) sCedula = row["Cedula"].ToString();
					if (row["Valor"] != System.DBNull.Value) dValor = Convert.ToDecimal(row["Valor"]);
					
					if (sCedula.Trim().Length == 9) sCedula = "0" + sCedula.Trim();
					
					string sSQL = string.Format("Exec NominaActualizaDescuentos {0}, 10, '{1}', {2}", (int)this.txtIdRolNomina.Value, sCedula, dValor);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);					

					iCont++;
				}				

				MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec NominaTotalizaValores {0}", (int)this.txtIdRolNomina.Value));
			
				this.btnVer_Click(sender, e);
				this.Cursor = Cursors.Default;
			}
		}
	}
}
