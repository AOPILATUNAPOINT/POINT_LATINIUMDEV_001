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

using DevExpress.XtraEditors;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmConsultaCbo_GestorDeCobranzas.
	/// </summary>
	public class frmConsultaCbo_GestorDeCobranzas : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbParentesco;
		private System.Windows.Forms.Label lblParentesco;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbGestion;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbGestionCredito;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Timer timer2;
		private System.Windows.Forms.Timer timer1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCobrador;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDMDesde;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDMHasta;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optAsignados;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbEntidadFinanciera;
		private System.Windows.Forms.Label label8;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBonificacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentaje;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobrado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProyectado;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCobradores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtfechas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobroExterno;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Button button1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSupervisor;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optGestion;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbOperadora;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstadosGestio;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbResultadoGestion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource13;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbZona;
		private System.Windows.Forms.Button button2;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultgconsulta;
		private System.Windows.Forms.Button btnImportar;
		private System.Windows.Forms.Button btnImpOperador;
		private System.ComponentModel.IContainer components;

		public frmConsultaCbo_GestorDeCobranzas()
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
		int iGrupo = 0;
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_GestorDeCobranzas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOperador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idZona");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCobrador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_GestionesDeCobranzas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Ultima_Gestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_EstadoGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_ResultadoGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero_Documento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Barrio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto_Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Venta_Cartera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital_Vendido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operacion_Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo_Deuda_Actual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo_Capital_Vencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo_Interes_Vencido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes_Mora_Actual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gastos_Cobranza_Actual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Cobrar_Proyectado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Actual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Retirado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Proyectado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasPendientes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasPagadas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cartera_Asignada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Recover_Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desempleo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_En_Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NDC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSupervisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rescate_Anteriores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IEmpleado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Cobrado_Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiaPago");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParentesco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParentesco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_ResultadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_ResultadoGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_EstadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_EstadoGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_GestionesDeCobranzas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operador_Cobrador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ESTADO_PAGO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DireccionUrlWSP");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConsultaCbo_GestorDeCobranzas));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCobrador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOperador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operadora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOperador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operadora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCobrador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_EstadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_EstadoGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_ResultadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Resultado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCbo_ResultadoGestion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Resultado");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_GestorDeCobranzas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOperador", -1, "cmbOperadora");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idZona", -1, "cmbZona");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobrador", -1, "cmbCobradores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_GestionesDeCobranzas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Ultima_Gestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_EstadoGestion", -1, "cmbEstadosGestio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCbo_ResultadoGestion", -1, "cmbResultadoGestion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero_Documento");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Barrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("telefono");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto_Credito");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Venta_Cartera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital_Vendido");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("lote");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operacion_Banco");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo_Deuda_Actual");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo_Capital_Vencido");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo_Interes_Vencido");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes_Mora_Actual");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gastos_Cobranza_Actual");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Cobrar_Proyectado");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Actual");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Cobrado");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Retirado");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Proyectado");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuotas");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasPendientes");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasPagadas");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cartera_Asignada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Recover_Cobrado");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desempleo");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas_En_Mora");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NDC");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor", -1, "cmbSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rescate_Anteriores");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IEmpleado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Cobrado_Total");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiaPago");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ir", 0);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cob", 1);
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Historico", 2);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NC", 3);
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEquifax", 4);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Cobrar_Proyectado", 33, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Cobrar_Proyectado", 33, true);
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Cobrado_Total", 51, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Cobrado_Total", 51, true);
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Retirado", 36, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Retirado", 36, true);
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo_Deuda_Actual", 28, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo_Deuda_Actual", 28, true);
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo_Capital_Vencido", 29, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo_Capital_Vencido", 29, true);
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo_Interes_Vencido", 30, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo_Interes_Vencido", 30, true);
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes_Mora_Actual", 31, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes_Mora_Actual", 31, true);
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Gastos_Cobranza_Actual", 32, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Gastos_Cobranza_Actual", 32, true);
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo_Capital_Vencido", 29, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo_Capital_Vencido", 29, true);
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, "", "Recover_Cobrado", 43, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Recover_Cobrado", 43, true);
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NDC", 46, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NDC", 46, true);
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Rescate_Anteriores", 48, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Rescate_Anteriores", 48, true);
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idZona");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Zona");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idZona");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Zona");
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbParentesco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblParentesco = new System.Windows.Forms.Label();
			this.cmbGestion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbGestionCredito = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.timer2 = new System.Windows.Forms.Timer(this.components);
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbCobrador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtDMDesde = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDMHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.optAsignados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cmbEntidadFinanciera = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label8 = new System.Windows.Forms.Label();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.txtBonificacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcentaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobrado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.txtProyectado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.cmbCobradores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtfechas = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtCobroExterno = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			this.cmbSupervisor = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.label14 = new System.Windows.Forms.Label();
			this.optGestion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label15 = new System.Windows.Forms.Label();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbOperadora = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstadosGestio = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbResultadoGestion = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbZona = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource13 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.button2 = new System.Windows.Forms.Button();
			this.ultgconsulta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnImportar = new System.Windows.Forms.Button();
			this.btnImpOperador = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestionCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDMDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDMHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optAsignados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtBonificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobrado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyectado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobradores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtfechas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobroExterno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSupervisor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optGestion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOperadora)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadosGestio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResultadoGestion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbZona)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgconsulta)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource5
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(System.DateTime);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(System.Decimal);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(System.Decimal);
			ultraDataColumn51.DataType = typeof(int);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn53.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn53});
			// 
			// cmbParentesco
			// 
			this.cmbParentesco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbParentesco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbParentesco.DataSource = this.ultraDataSource3;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Parentesco";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbParentesco.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbParentesco.DisplayMember = "Nombre";
			this.cmbParentesco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbParentesco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbParentesco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbParentesco.Location = new System.Drawing.Point(560, 202);
			this.cmbParentesco.MaxDropDownItems = 30;
			this.cmbParentesco.Name = "cmbParentesco";
			this.cmbParentesco.Size = new System.Drawing.Size(16, 21);
			this.cmbParentesco.TabIndex = 859;
			this.cmbParentesco.ValueMember = "idParentesco";
			this.cmbParentesco.Visible = false;
			// 
			// ultraDataSource3
			// 
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn54,
																																 ultraDataColumn55});
			// 
			// lblParentesco
			// 
			this.lblParentesco.AutoSize = true;
			this.lblParentesco.Location = new System.Drawing.Point(624, 200);
			this.lblParentesco.Name = "lblParentesco";
			this.lblParentesco.Size = new System.Drawing.Size(59, 17);
			this.lblParentesco.TabIndex = 858;
			this.lblParentesco.Text = "Parentesco";
			this.lblParentesco.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblParentesco.Visible = false;
			// 
			// cmbGestion
			// 
			this.cmbGestion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGestion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGestion.DataSource = this.ultraDataSource2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "Gestión";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 278;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbGestion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbGestion.DisplayMember = "Resultado";
			this.cmbGestion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGestion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGestion.Location = new System.Drawing.Point(72, 104);
			this.cmbGestion.MaxDropDownItems = 30;
			this.cmbGestion.Name = "cmbGestion";
			this.cmbGestion.Size = new System.Drawing.Size(280, 21);
			this.cmbGestion.TabIndex = 857;
			this.cmbGestion.ValueMember = "idCbo_ResultadoGestion";
			this.cmbGestion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGestion_KeyDown);
			this.cmbGestion.ValueChanged += new System.EventHandler(this.cmbGestion_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn56.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn56,
																																 ultraDataColumn57});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 72);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 17);
			this.label1.TabIndex = 856;
			this.label1.Text = "Gestión";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbGestionCredito
			// 
			this.cmbGestionCredito.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGestionCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbGestionCredito.DataSource = this.ultraDataSource1;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 150;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 278;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbGestionCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbGestionCredito.DisplayMember = "Estado";
			this.cmbGestionCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGestionCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGestionCredito.Location = new System.Drawing.Point(72, 70);
			this.cmbGestionCredito.MaxDropDownItems = 30;
			this.cmbGestionCredito.Name = "cmbGestionCredito";
			this.cmbGestionCredito.Size = new System.Drawing.Size(280, 21);
			this.cmbGestionCredito.TabIndex = 855;
			this.cmbGestionCredito.ValueMember = "idCbo_EstadoGestion";
			this.cmbGestionCredito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGestionCredito_KeyDown);
			this.cmbGestionCredito.ValueChanged += new System.EventHandler(this.cmbGestionCredito_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn58.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn58,
																																 ultraDataColumn59});
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 104);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(38, 17);
			this.label3.TabIndex = 854;
			this.label3.Text = "Estado";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn60.DataType = typeof(int);
			ultraDataColumn61.DataType = typeof(int);
			ultraDataColumn62.DataType = typeof(System.DateTime);
			ultraDataColumn67.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn75});
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
			// timer2
			// 
			this.timer2.Enabled = true;
			this.timer2.Interval = 1000;
			// 
			// timer1
			// 
			this.timer1.Interval = 1000;
			// 
			// txtBuscar
			// 
			this.txtBuscar.AcceptsReturn = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance1;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(72, 201);
			this.txtBuscar.MaxLength = 1000;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(280, 22);
			this.txtBuscar.TabIndex = 861;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 204);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 17);
			this.label2.TabIndex = 860;
			this.label2.Text = "Buscar";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(42, 17);
			this.label4.TabIndex = 863;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 40);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(51, 17);
			this.label5.TabIndex = 865;
			this.label5.Text = "Operador";
			// 
			// cmbCobrador
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCobrador.Appearance = appearance2;
			this.cmbCobrador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCobrador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobrador.DataSource = this.ultraDataSource6;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 275;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbCobrador.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCobrador.DisplayMember = "Cargo";
			this.cmbCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCobrador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCobrador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobrador.Location = new System.Drawing.Point(72, 40);
			this.cmbCobrador.Name = "cmbCobrador";
			this.cmbCobrador.Size = new System.Drawing.Size(280, 21);
			this.cmbCobrador.TabIndex = 866;
			this.cmbCobrador.ValueMember = "idCargo";
			this.cmbCobrador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCobrador_KeyDown);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn76.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn76,
																																 ultraDataColumn77});
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 232);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1208, 8);
			this.groupBox2.TabIndex = 953;
			this.groupBox2.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtDMDesde);
			this.groupBox1.Controls.Add(this.txtDMHasta);
			this.groupBox1.Controls.Add(this.label6);
			this.groupBox1.Controls.Add(this.label7);
			this.groupBox1.Location = new System.Drawing.Point(400, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(288, 56);
			this.groupBox1.TabIndex = 954;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Dias de Mora";
			// 
			// txtDMDesde
			// 
			this.txtDMDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDMDesde.Location = new System.Drawing.Point(72, 24);
			this.txtDMDesde.MaxValue = 3600000;
			this.txtDMDesde.MinValue = 0;
			this.txtDMDesde.Name = "txtDMDesde";
			this.txtDMDesde.PromptChar = ' ';
			this.txtDMDesde.Size = new System.Drawing.Size(72, 22);
			this.txtDMDesde.TabIndex = 671;
			this.txtDMDesde.Value = 1;
			// 
			// txtDMHasta
			// 
			this.txtDMHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDMHasta.Location = new System.Drawing.Point(200, 24);
			this.txtDMHasta.MaxValue = 3600000;
			this.txtDMHasta.MinValue = 0;
			this.txtDMHasta.Name = "txtDMHasta";
			this.txtDMHasta.PromptChar = ' ';
			this.txtDMHasta.Size = new System.Drawing.Size(72, 22);
			this.txtDMHasta.TabIndex = 673;
			this.txtDMHasta.Value = 150;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(152, 26);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(33, 17);
			this.label6.TabIndex = 674;
			this.label6.Text = "Hasta";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(16, 26);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(35, 17);
			this.label7.TabIndex = 672;
			this.label7.Text = "Desde";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnExcel
			// 
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1136, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 956;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(400, 200);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 955;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// optAsignados
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.optAsignados.Appearance = appearance3;
			this.optAsignados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optAsignados.CheckedIndex = 0;
			this.optAsignados.ItemAppearance = appearance4;
			this.optAsignados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Gestión";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Sin Gestión";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Compromiso De  Pago";
			this.optAsignados.Items.Add(valueListItem1);
			this.optAsignados.Items.Add(valueListItem2);
			this.optAsignados.Items.Add(valueListItem3);
			this.optAsignados.Items.Add(valueListItem4);
			this.optAsignados.ItemSpacingVertical = 8;
			this.optAsignados.Location = new System.Drawing.Point(192, 8);
			this.optAsignados.Name = "optAsignados";
			this.optAsignados.Size = new System.Drawing.Size(360, 26);
			this.optAsignados.TabIndex = 957;
			this.optAsignados.Text = "Todos";
			this.optAsignados.ValueChanged += new System.EventHandler(this.optAsignados_ValueChanged);
			// 
			// cmbEntidadFinanciera
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbEntidadFinanciera.Appearance = appearance5;
			this.cmbEntidadFinanciera.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntidadFinanciera.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 208;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbEntidadFinanciera.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbEntidadFinanciera.DisplayMember = "Nombre";
			this.cmbEntidadFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntidadFinanciera.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntidadFinanciera.Location = new System.Drawing.Point(72, 136);
			this.cmbEntidadFinanciera.Name = "cmbEntidadFinanciera";
			this.cmbEntidadFinanciera.Size = new System.Drawing.Size(280, 21);
			this.cmbEntidadFinanciera.TabIndex = 958;
			this.cmbEntidadFinanciera.ValueMember = "idEntidadFinanciera";
			this.cmbEntidadFinanciera.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEntidadFinanciera_KeyDown);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 136);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 17);
			this.label8.TabIndex = 959;
			this.label8.Text = "Banco";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.txtBonificacion);
			this.groupBox3.Controls.Add(this.txtPorcentaje);
			this.groupBox3.Controls.Add(this.txtCobrado);
			this.groupBox3.Controls.Add(this.label12);
			this.groupBox3.Controls.Add(this.label11);
			this.groupBox3.Controls.Add(this.txtProyectado);
			this.groupBox3.Controls.Add(this.label10);
			this.groupBox3.Controls.Add(this.label9);
			this.groupBox3.Location = new System.Drawing.Point(400, 104);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(744, 48);
			this.groupBox3.TabIndex = 960;
			this.groupBox3.TabStop = false;
			// 
			// txtBonificacion
			// 
			appearance6.BackColor = System.Drawing.Color.Transparent;
			appearance6.BackColor2 = System.Drawing.Color.Transparent;
			appearance6.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance6.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBonificacion.Appearance = appearance6;
			this.txtBonificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBonificacion.Enabled = false;
			this.txtBonificacion.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBonificacion.Location = new System.Drawing.Point(624, 16);
			this.txtBonificacion.MaxValue = 8000000;
			this.txtBonificacion.MinValue = 0;
			this.txtBonificacion.Name = "txtBonificacion";
			this.txtBonificacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtBonificacion.PromptChar = ' ';
			this.txtBonificacion.Size = new System.Drawing.Size(104, 25);
			this.txtBonificacion.TabIndex = 680;
			this.txtBonificacion.Visible = false;
			// 
			// txtPorcentaje
			// 
			appearance7.BackColor = System.Drawing.Color.Transparent;
			appearance7.BackColor2 = System.Drawing.Color.Transparent;
			appearance7.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance7.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcentaje.Appearance = appearance7;
			this.txtPorcentaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentaje.Enabled = false;
			this.txtPorcentaje.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtPorcentaje.Location = new System.Drawing.Point(440, 16);
			this.txtPorcentaje.MaxValue = ((long)(80000000000));
			this.txtPorcentaje.MinValue = 0;
			this.txtPorcentaje.Name = "txtPorcentaje";
			this.txtPorcentaje.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcentaje.PromptChar = ' ';
			this.txtPorcentaje.Size = new System.Drawing.Size(104, 25);
			this.txtPorcentaje.TabIndex = 679;
			// 
			// txtCobrado
			// 
			appearance8.BackColor = System.Drawing.Color.Transparent;
			appearance8.BackColor2 = System.Drawing.Color.Transparent;
			appearance8.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance8.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobrado.Appearance = appearance8;
			this.txtCobrado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobrado.Enabled = false;
			this.txtCobrado.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCobrado.Location = new System.Drawing.Point(248, 16);
			this.txtCobrado.MaxValue = ((long)(100000000000));
			this.txtCobrado.MinValue = 0;
			this.txtCobrado.Name = "txtCobrado";
			this.txtCobrado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobrado.PromptChar = ' ';
			this.txtCobrado.Size = new System.Drawing.Size(104, 25);
			this.txtCobrado.TabIndex = 678;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(552, 16);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(64, 17);
			this.label12.TabIndex = 677;
			this.label12.Text = "Bonificación";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label12.Visible = false;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(360, 16);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(72, 17);
			this.label11.TabIndex = 676;
			this.label11.Text = "% De Avance";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtProyectado
			// 
			appearance9.BackColor = System.Drawing.Color.Transparent;
			appearance9.BackColor2 = System.Drawing.Color.Transparent;
			appearance9.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance9.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProyectado.Appearance = appearance9;
			this.txtProyectado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProyectado.Enabled = false;
			this.txtProyectado.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtProyectado.Location = new System.Drawing.Point(80, 16);
			this.txtProyectado.MaxValue = ((long)(100000000000));
			this.txtProyectado.MinValue = 0;
			this.txtProyectado.Name = "txtProyectado";
			this.txtProyectado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProyectado.PromptChar = ' ';
			this.txtProyectado.Size = new System.Drawing.Size(104, 25);
			this.txtProyectado.TabIndex = 671;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(8, 16);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(60, 17);
			this.label10.TabIndex = 672;
			this.label10.Text = "Proyectado";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(192, 16);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(46, 17);
			this.label9.TabIndex = 674;
			this.label9.Text = "Cobrado";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(496, 201);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 979;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbCobradores
			// 
			this.cmbCobradores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobradores.DataSource = this.ultraDataSource7;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.Caption = "Cobrador";
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 250;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbCobradores.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCobradores.DisplayMember = "Nombre";
			this.cmbCobradores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobradores.Location = new System.Drawing.Point(224, 248);
			this.cmbCobradores.Name = "cmbCobradores";
			this.cmbCobradores.Size = new System.Drawing.Size(272, 88);
			this.cmbCobradores.TabIndex = 980;
			this.cmbCobradores.ValueMember = "idCobrador";
			this.cmbCobradores.Visible = false;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn78.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn78,
																																 ultraDataColumn79});
			// 
			// dtFecha
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance10;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dtFecha.Format = "MM/yyyy ";
			this.dtFecha.Location = new System.Drawing.Point(72, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 862;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// dtfechas
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtfechas.Appearance = appearance11;
			this.dtfechas.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtfechas.DateButtons.Add(dateButton2);
			this.dtfechas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dtfechas.Format = "dd/MM/yyyy ";
			this.dtfechas.Location = new System.Drawing.Point(1016, 8);
			this.dtfechas.Name = "dtfechas";
			this.dtfechas.NonAutoSizeHeight = 23;
			this.dtfechas.Size = new System.Drawing.Size(104, 21);
			this.dtfechas.SpinButtonsVisible = true;
			this.dtfechas.TabIndex = 981;
			this.dtfechas.Value = ((object)(resources.GetObject("dtfechas.Value")));
			this.dtfechas.Visible = false;
			// 
			// txtCobroExterno
			// 
			appearance12.BackColor = System.Drawing.Color.Transparent;
			appearance12.BackColor2 = System.Drawing.Color.Transparent;
			appearance12.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance12.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCobroExterno.Appearance = appearance12;
			this.txtCobroExterno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobroExterno.Enabled = false;
			this.txtCobroExterno.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCobroExterno.Location = new System.Drawing.Point(1024, 160);
			this.txtCobroExterno.MaxValue = 8000000;
			this.txtCobroExterno.MinValue = 0;
			this.txtCobroExterno.Name = "txtCobroExterno";
			this.txtCobroExterno.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobroExterno.PromptChar = ' ';
			this.txtCobroExterno.Size = new System.Drawing.Size(104, 25);
			this.txtCobroExterno.TabIndex = 982;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(936, 164);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(75, 17);
			this.label13.TabIndex = 983;
			this.label13.Text = "Cobro Externo";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// button1
			// 
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(1152, 160);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(32, 24);
			this.button1.TabIndex = 984;
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// cmbSupervisor
			// 
			this.cmbSupervisor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSupervisor.DataSource = this.ultraDataSource7;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.Caption = "Cobrador";
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 250;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbSupervisor.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbSupervisor.DisplayMember = "Nombre";
			this.cmbSupervisor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSupervisor.Location = new System.Drawing.Point(520, 248);
			this.cmbSupervisor.Name = "cmbSupervisor";
			this.cmbSupervisor.Size = new System.Drawing.Size(272, 88);
			this.cmbSupervisor.TabIndex = 985;
			this.cmbSupervisor.ValueMember = "idCobrador";
			this.cmbSupervisor.Visible = false;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Bisque;
			this.label14.Location = new System.Drawing.Point(1088, 40);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(123, 17);
			this.label14.TabIndex = 987;
			this.label14.Text = "Creditos Nomina POINT";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optGestion
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.optGestion.Appearance = appearance13;
			this.optGestion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optGestion.CheckedIndex = 0;
			this.optGestion.ItemAppearance = appearance14;
			this.optGestion.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem5.DataValue = 0;
			valueListItem5.DisplayText = "Gestión";
			valueListItem6.DataValue = 1;
			valueListItem6.DisplayText = "Almacen";
			this.optGestion.Items.Add(valueListItem5);
			this.optGestion.Items.Add(valueListItem6);
			this.optGestion.ItemSpacingVertical = 8;
			this.optGestion.Location = new System.Drawing.Point(560, 8);
			this.optGestion.Name = "optGestion";
			this.optGestion.Size = new System.Drawing.Size(152, 26);
			this.optGestion.TabIndex = 988;
			this.optGestion.Text = "Gestión";
			this.optGestion.ValueChanged += new System.EventHandler(this.optGestion_ValueChanged);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(8, 168);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(47, 17);
			this.label15.TabIndex = 990;
			this.label15.Text = "Almacen";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource8
			// 
			ultraDataColumn80.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn80,
																																 ultraDataColumn81});
			// 
			// cmbBodega
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance15;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource8;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 270;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(72, 168);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(280, 21);
			this.cmbBodega.TabIndex = 991;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// cmbOperadora
			// 
			this.cmbOperadora.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbOperadora.DataSource = this.ultraDataSource10;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.Caption = "Cobrador";
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 250;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbOperadora.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbOperadora.DisplayMember = "Operadora";
			this.cmbOperadora.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbOperadora.Location = new System.Drawing.Point(800, 248);
			this.cmbOperadora.Name = "cmbOperadora";
			this.cmbOperadora.Size = new System.Drawing.Size(272, 88);
			this.cmbOperadora.TabIndex = 992;
			this.cmbOperadora.ValueMember = "idOperador";
			this.cmbOperadora.Visible = false;
			// 
			// ultraDataSource10
			// 
			ultraDataColumn82.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn82,
																																	ultraDataColumn83});
			// 
			// ultraDataSource9
			// 
			ultraDataColumn84.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn84,
																																 ultraDataColumn85});
			// 
			// cmbEstadosGestio
			// 
			this.cmbEstadosGestio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstadosGestio.DataSource = this.ultraDataSource11;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 250;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn19,
																											ultraGridColumn20});
			this.cmbEstadosGestio.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbEstadosGestio.DisplayMember = "Estado";
			this.cmbEstadosGestio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstadosGestio.Location = new System.Drawing.Point(240, 296);
			this.cmbEstadosGestio.Name = "cmbEstadosGestio";
			this.cmbEstadosGestio.Size = new System.Drawing.Size(272, 88);
			this.cmbEstadosGestio.TabIndex = 995;
			this.cmbEstadosGestio.ValueMember = "idCbo_EstadoGestion";
			this.cmbEstadosGestio.Visible = false;
			// 
			// ultraDataSource11
			// 
			ultraDataColumn86.DataType = typeof(int);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn86,
																																	ultraDataColumn87});
			// 
			// cmbResultadoGestion
			// 
			this.cmbResultadoGestion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResultadoGestion.DataSource = this.ultraDataSource12;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 250;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn21,
																											ultraGridColumn22});
			this.cmbResultadoGestion.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbResultadoGestion.DisplayMember = "Resultado";
			this.cmbResultadoGestion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResultadoGestion.Location = new System.Drawing.Point(536, 288);
			this.cmbResultadoGestion.Name = "cmbResultadoGestion";
			this.cmbResultadoGestion.Size = new System.Drawing.Size(272, 88);
			this.cmbResultadoGestion.TabIndex = 996;
			this.cmbResultadoGestion.ValueMember = "idCbo_ResultadoGestion";
			this.cmbResultadoGestion.Visible = false;
			// 
			// ultraDataSource12
			// 
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn88,
																																	ultraDataColumn89});
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource5;
			appearance16.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance16;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 75;
			ultraGridColumn25.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.Caption = "Operador";
			ultraGridColumn25.Header.VisiblePosition = 2;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn25.Width = 250;
			ultraGridColumn26.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn26.Header.Caption = "Zona";
			ultraGridColumn26.Header.VisiblePosition = 7;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn26.Width = 150;
			ultraGridColumn27.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn27.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn27.Header.Caption = "Cobrador";
			ultraGridColumn27.Header.VisiblePosition = 8;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn27.Width = 250;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn28.Header.VisiblePosition = 9;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn29.Format = "yyy/MM/dd HH:mm:ss";
			ultraGridColumn29.Header.Caption = "F. Ultima Gestion";
			ultraGridColumn29.Header.VisiblePosition = 10;
			ultraGridColumn29.Width = 120;
			ultraGridColumn30.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.Caption = "Gestión";
			ultraGridColumn30.Header.VisiblePosition = 11;
			ultraGridColumn30.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn30.Width = 250;
			ultraGridColumn31.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.Caption = "Estado";
			ultraGridColumn31.Header.VisiblePosition = 12;
			ultraGridColumn31.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn31.Width = 250;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn32.Header.VisiblePosition = 13;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn33.Header.VisiblePosition = 14;
			ultraGridColumn33.Width = 100;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn34.Format = "dd/MM/yyyy";
			ultraGridColumn34.Header.Caption = "F. Emisión";
			ultraGridColumn34.Header.VisiblePosition = 15;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Format = "dd/MM/yyyy";
			ultraGridColumn35.Header.Caption = "F. Origen";
			ultraGridColumn35.Header.VisiblePosition = 16;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn36.CellAppearance = appearance17;
			ultraGridColumn36.Header.Caption = "# Documento";
			ultraGridColumn36.Header.VisiblePosition = 18;
			ultraGridColumn36.Width = 100;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance18;
			ultraGridColumn37.Header.VisiblePosition = 19;
			ultraGridColumn37.Width = 90;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn38.Header.VisiblePosition = 20;
			ultraGridColumn38.Width = 300;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn39.Header.VisiblePosition = 21;
			ultraGridColumn39.Width = 100;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn40.Header.VisiblePosition = 22;
			ultraGridColumn40.Width = 250;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn41.Header.Caption = "Dirección";
			ultraGridColumn41.Header.VisiblePosition = 23;
			ultraGridColumn41.Width = 250;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance19;
			ultraGridColumn42.Header.Caption = "Teléfono";
			ultraGridColumn42.Header.VisiblePosition = 26;
			ultraGridColumn42.Width = 100;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance20;
			ultraGridColumn43.Header.VisiblePosition = 25;
			ultraGridColumn43.Width = 100;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance21;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.Caption = "Monto Credito";
			ultraGridColumn44.Header.VisiblePosition = 27;
			ultraGridColumn44.Width = 110;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn45.Header.VisiblePosition = 28;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 90;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn46.Header.VisiblePosition = 29;
			ultraGridColumn46.Width = 90;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn47.Header.Caption = "F.Venta Cartera";
			ultraGridColumn47.Header.VisiblePosition = 31;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 90;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance22;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.Caption = "Capital Vendido";
			ultraGridColumn48.Header.VisiblePosition = 32;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 110;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn49.CellAppearance = appearance23;
			ultraGridColumn49.Header.Caption = "Lote";
			ultraGridColumn49.Header.VisiblePosition = 33;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 80;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance24;
			ultraGridColumn50.Header.Caption = "Operación";
			ultraGridColumn50.Header.VisiblePosition = 34;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 90;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance25;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.Caption = "Saldo Deuda Actual";
			ultraGridColumn51.Header.VisiblePosition = 35;
			ultraGridColumn51.Width = 150;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance26;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.Caption = "Saldo Capital Vencido";
			ultraGridColumn52.Header.VisiblePosition = 36;
			ultraGridColumn52.Width = 150;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance27;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.Caption = "S.Interes Vencido";
			ultraGridColumn53.Header.VisiblePosition = 37;
			ultraGridColumn53.Width = 110;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance28;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.Caption = "Int.Mora Actual";
			ultraGridColumn54.Header.VisiblePosition = 38;
			ultraGridColumn54.Width = 110;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance29;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.Caption = "Gast.Cobranza Actual";
			ultraGridColumn55.Header.VisiblePosition = 43;
			ultraGridColumn55.Width = 110;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance30.BackColor = System.Drawing.Color.LimeGreen;
			appearance30.BackColor2 = System.Drawing.Color.LimeGreen;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance30;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.Caption = "V.Cobr.Proyectado";
			ultraGridColumn56.Header.VisiblePosition = 51;
			ultraGridColumn56.Width = 110;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn57.CellAppearance = appearance31;
			ultraGridColumn57.Header.Caption = "Dias Mora";
			ultraGridColumn57.Header.VisiblePosition = 41;
			ultraGridColumn57.Width = 90;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.BackColorDisabled = System.Drawing.Color.LightBlue;
			appearance32.BackColorDisabled2 = System.Drawing.Color.LightBlue;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance32;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.Caption = "Valor Cobrado";
			ultraGridColumn58.Header.VisiblePosition = 49;
			ultraGridColumn58.Width = 140;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance33.BackColor = System.Drawing.Color.Firebrick;
			appearance33.BackColor2 = System.Drawing.Color.Firebrick;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance33;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.Caption = "V. Retirado";
			ultraGridColumn59.Header.VisiblePosition = 55;
			ultraGridColumn59.Width = 110;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn60.CellAppearance = appearance34;
			ultraGridColumn60.Header.Caption = "Dias Mora Pr.";
			ultraGridColumn60.Header.VisiblePosition = 42;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn61.CellAppearance = appearance35;
			ultraGridColumn61.Header.Caption = "# Cuotas";
			ultraGridColumn61.Header.VisiblePosition = 39;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance36;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.Caption = "Valor Cuota";
			ultraGridColumn62.Header.VisiblePosition = 40;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn63.CellAppearance = appearance37;
			ultraGridColumn63.Header.Caption = "C. Pendientes";
			ultraGridColumn63.Header.VisiblePosition = 44;
			ultraGridColumn63.Width = 80;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn64.CellAppearance = appearance38;
			ultraGridColumn64.Header.Caption = "C. Pagadas";
			ultraGridColumn64.Header.VisiblePosition = 45;
			ultraGridColumn64.Width = 80;
			ultraGridColumn65.Header.VisiblePosition = 46;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance39.BackColor = System.Drawing.Color.Red;
			appearance39.BackColor2 = System.Drawing.Color.Red;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance39;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellButtonAppearance = appearance40;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.Caption = "Cobro Externo ";
			ultraGridColumn66.Header.VisiblePosition = 54;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn67.CellAppearance = appearance41;
			ultraGridColumn67.Header.VisiblePosition = 47;
			ultraGridColumn68.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn68.CellAppearance = appearance42;
			ultraGridColumn68.Header.Caption = "Cuotas En Mora";
			ultraGridColumn68.Header.VisiblePosition = 48;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance43;
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.VisiblePosition = 53;
			ultraGridColumn70.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn70.Header.Caption = "Supervisor";
			ultraGridColumn70.Header.VisiblePosition = 24;
			ultraGridColumn70.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn70.Width = 200;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance44;
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.Caption = "Rescate Anteriores";
			ultraGridColumn71.Header.VisiblePosition = 52;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn72.Header.VisiblePosition = 30;
			ultraGridColumn72.Width = 200;
			ultraGridColumn73.Header.VisiblePosition = 56;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance45.BackColor = System.Drawing.Color.LightBlue;
			appearance45.BackColor2 = System.Drawing.Color.LightBlue;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance45;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.Caption = "Valor Cobrado Total";
			ultraGridColumn74.Header.VisiblePosition = 50;
			ultraGridColumn74.Width = 160;
			ultraGridColumn75.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn75.CellAppearance = appearance46;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn75.FormulaErrorAppearance = appearance47;
			ultraGridColumn75.Header.Caption = "Día de Pago";
			ultraGridColumn75.Header.VisiblePosition = 17;
			ultraGridColumn76.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance48.Image = ((object)(resources.GetObject("appearance48.Image")));
			appearance48.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn76.CellButtonAppearance = appearance48;
			ultraGridColumn76.Header.Caption = "...";
			ultraGridColumn76.Header.VisiblePosition = 3;
			ultraGridColumn76.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn76.Width = 30;
			ultraGridColumn77.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance49.Image = ((object)(resources.GetObject("appearance49.Image")));
			appearance49.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn77.CellButtonAppearance = appearance49;
			ultraGridColumn77.Header.Caption = "...";
			ultraGridColumn77.Header.VisiblePosition = 5;
			ultraGridColumn77.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn77.Width = 30;
			ultraGridColumn78.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance50.Image = ((object)(resources.GetObject("appearance50.Image")));
			appearance50.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn78.CellButtonAppearance = appearance50;
			ultraGridColumn78.Header.Caption = "...";
			ultraGridColumn78.Header.VisiblePosition = 4;
			ultraGridColumn78.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn78.Width = 30;
			ultraGridColumn79.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance51.Image = ((object)(resources.GetObject("appearance51.Image")));
			appearance51.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn79.CellButtonAppearance = appearance51;
			ultraGridColumn79.Header.Caption = "N/C";
			ultraGridColumn79.Header.VisiblePosition = 57;
			ultraGridColumn79.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn79.Width = 30;
			ultraGridColumn80.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn80.Header.Caption = "Equifax";
			ultraGridColumn80.Header.VisiblePosition = 6;
			ultraGridColumn80.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn80.Width = 50;
			ultraGridBand12.Columns.AddRange(new object[] {
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
																											ultraGridColumn80});
			ultraGridBand12.Header.Caption = "Periodo A";
			ultraGridBand12.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand12.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand12.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance52;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance53;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance54;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance55;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance56;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.Hidden = true;
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance57;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.Hidden = true;
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance58;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.Hidden = true;
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance59;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.Hidden = true;
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance60;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.Hidden = true;
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance61;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.Hidden = true;
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance62;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance63;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand12.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															 summarySettings12});
			ultraGridBand12.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance64.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance64;
			appearance65.ForeColor = System.Drawing.Color.Black;
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance65;
			appearance66.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance66.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance66.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance66.FontData.BoldAsString = "True";
			appearance66.FontData.Name = "Arial";
			appearance66.FontData.SizeInPoints = 8F;
			appearance66.ForeColor = System.Drawing.Color.White;
			appearance66.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance66;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance67.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance67.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance67.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance67;
			appearance68.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance68.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance68.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance68;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 248);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1200, 288);
			this.uGridInformacion.TabIndex = 998;
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton_1);
			this.uGridInformacion.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_AfterCellUpdate_1);
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout_1);
			// 
			// cmbZona
			// 
			this.cmbZona.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbZona.DataSource = this.ultraDataSource13;
			ultraGridColumn81.Header.VisiblePosition = 0;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 1;
			ultraGridColumn82.Width = 250;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn81,
																											ultraGridColumn82});
			this.cmbZona.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbZona.DisplayMember = "Zona";
			this.cmbZona.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbZona.Location = new System.Drawing.Point(824, 352);
			this.cmbZona.Name = "cmbZona";
			this.cmbZona.Size = new System.Drawing.Size(272, 88);
			this.cmbZona.TabIndex = 999;
			this.cmbZona.ValueMember = "idZona";
			this.cmbZona.Visible = false;
			// 
			// ultraDataSource13
			// 
			ultraDataColumn90.DataType = typeof(int);
			this.ultraDataSource13.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn90,
																																	ultraDataColumn91});
			// 
			// button2
			// 
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.Location = new System.Drawing.Point(1136, 72);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(75, 24);
			this.button2.TabIndex = 1000;
			this.button2.Text = "&Barrios";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button2.Visible = false;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// ultgconsulta
			// 
			this.ultgconsulta.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultgconsulta.Cursor = System.Windows.Forms.Cursors.Default;
			appearance69.BackColor = System.Drawing.Color.White;
			appearance69.ForegroundAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultgconsulta.DisplayLayout.Appearance = appearance69;
			this.ultgconsulta.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultgconsulta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultgconsulta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance70.BackColor = System.Drawing.Color.Transparent;
			this.ultgconsulta.DisplayLayout.Override.CardAreaAppearance = appearance70;
			appearance71.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance71.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance71.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance71.FontData.BoldAsString = "True";
			appearance71.FontData.Name = "Arial";
			appearance71.FontData.SizeInPoints = 8.5F;
			appearance71.ForeColor = System.Drawing.Color.White;
			appearance71.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultgconsulta.DisplayLayout.Override.HeaderAppearance = appearance71;
			this.ultgconsulta.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance72.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance72.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgconsulta.DisplayLayout.Override.RowSelectorAppearance = appearance72;
			appearance73.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance73.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgconsulta.DisplayLayout.Override.SelectedRowAppearance = appearance73;
			this.ultgconsulta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultgconsulta.Location = new System.Drawing.Point(760, 40);
			this.ultgconsulta.Name = "ultgconsulta";
			this.ultgconsulta.Size = new System.Drawing.Size(160, 40);
			this.ultgconsulta.TabIndex = 1001;
			this.ultgconsulta.Visible = false;
			// 
			// btnImportar
			// 
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(1048, 72);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.Size = new System.Drawing.Size(75, 24);
			this.btnImportar.TabIndex = 1002;
			this.btnImportar.Text = "&Importar";
			this.btnImportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportar.Visible = false;
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// btnImpOperador
			// 
			this.btnImpOperador.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImpOperador.Location = new System.Drawing.Point(920, 72);
			this.btnImpOperador.Name = "btnImpOperador";
			this.btnImpOperador.Size = new System.Drawing.Size(112, 24);
			this.btnImpOperador.TabIndex = 1003;
			this.btnImpOperador.Text = "&Importar Operador";
			this.btnImpOperador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImpOperador.Visible = false;
			this.btnImpOperador.Click += new System.EventHandler(this.btnImpOperador_Click);
			// 
			// frmConsultaCbo_GestorDeCobranzas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1216, 542);
			this.Controls.Add(this.btnImpOperador);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.ultgconsulta);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.cmbZona);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.cmbResultadoGestion);
			this.Controls.Add(this.cmbEstadosGestio);
			this.Controls.Add(this.cmbOperadora);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblParentesco);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.optGestion);
			this.Controls.Add(this.cmbSupervisor);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.txtCobroExterno);
			this.Controls.Add(this.dtfechas);
			this.Controls.Add(this.cmbCobradores);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.cmbEntidadFinanciera);
			this.Controls.Add(this.optAsignados);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.cmbCobrador);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.cmbParentesco);
			this.Controls.Add(this.cmbGestion);
			this.Controls.Add(this.cmbGestionCredito);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmConsultaCbo_GestorDeCobranzas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Informes Operadora / Cobrador";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConsultaCbo_GestorDeCobranzas_KeyDown);
			this.Load += new System.EventHandler(this.frmConsultaCbo_GestorDeCobranzas_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGestionCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDMDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDMHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optAsignados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidadFinanciera)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtBonificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobrado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProyectado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobradores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtfechas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobroExterno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSupervisor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optGestion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOperadora)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadosGestio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResultadoGestion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbZona)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgconsulta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		int iMinimo = 0;
		int iMaximo = 0;

		private void frmConsultaCbo_GestorDeCobranzas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "082916");

			if (!Funcion.Permiso("1194", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Proformas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			DateTime dtFecha ;
			FuncionesProcedimientos.RetornaFechaServidor(this.dtfechas, cdsSeteoF, false);

			int	iMesA = Convert.ToInt32( ( Convert.ToDateTime(this.dtfechas.Value).Month));
			int iYearA = Convert.ToInt32( ( Convert.ToDateTime(this.dtfechas.Value).Year));
			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);

			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
		
			if (miAcceso.BExportar) this.btnExcel.Enabled =  true;
			if (!miAcceso.CambiarFecha) this.dtFecha.Enabled =  false;
			if (miAcceso.Buscar) 
			{
				  this.btnImpOperador.Visible =  true;	
					this.btnImportar.Visible =  true;	
				this.button2.Visible = true;
			}

			if (miAcceso.Editar) 
			{
				int idia = Convert.ToInt32( ( Convert.ToDateTime(this.dtfechas.Value).Day));
				if (idia <= 16 )
				{
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["idCobrador"].CellActivation = Activation.AllowEdit; 
				}
				else

				{
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["idCobrador"].CellActivation = Activation.Disabled; 
				}
			}
			else
			{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["idCobrador"].CellActivation = Activation.Disabled; 
			}
			//		 this.txtProyectado.BackColor = Color.Black;
			//        this.txtProyectado.ForeColor = Color.Red;
			

			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("01/03/2019");
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;		
			this.dtFecha.Value = DateTime.Today;
			
			int iPersonal = 0;

			string ssqlDo = string.Format ("Exec ConsultaCbo_OperadorCobrador {0}, 1", MenuLatinium.IdUsuario ); 
			SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
			drDO.Read();
			if(drDO.HasRows)
			{
				iGrupo = drDO.GetInt32 (0);
			}
			drDO.Close();

			string sSqlID = string.Format ("Exec ConsultaCbo_OperadorCobrador {0}, 2", MenuLatinium.IdUsuario ); 
			SqlDataReader drId = Funcion.rEscalarSQL(cdsSeteoF,sSqlID,true);
			drId.Read();
			if(drId.HasRows)
			{
				iPersonal = drId.GetInt32 (0);
			}
			drId.Close();

			this.cmbGestionCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCbo_EstadoGestion, Estado from dbo.Cbo_EstadosGestion");
			
			//    	this.cmbCobrador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select n.idPersonal,( ApellidoPaterno + ' '+ ApellidoMaterno +' '+ PrimerNombre + ' '+ SegundoNombre) as Nombre from Nomina n Inner Join RolCargo c on c.idCargo =  n.idCargo where n.idCargo in (121,122,65,123,70,91,120,125,97,119) and n.Estado = 1 order by Nombre");
      
			this.cmbEntidadFinanciera.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre From dbo.Cre_EntidadFinanciera  where idEntidadFinanciera not in (6) Order By idEntidadFinanciera");

			#region OperadoCobrador
			if(iGrupo == 19 || iGrupo == 31)
			{
				
				iPersonal = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ConsultaCbo_OperadorCobrador {0}, 2", MenuLatinium.IdUsuario));	
				this.cmbCobrador.Enabled = true ;
				this.cmbCobrador.Value =  iPersonal ;

				string sSqlMin = string.Format ("Exec ConsultaCbo_OperadorCobrador {0}, 3, {1}", MenuLatinium.IdUsuario, iPersonal ); 
				SqlDataReader drMin = Funcion.rEscalarSQL(cdsSeteoF,sSqlMin,true);
				drMin.Read();
				if(drMin.HasRows)
				{
					iMinimo = drMin.GetInt32 (0);
					iMaximo = drMin.GetInt32 (1);
				}
				drMin.Close();

				if (iGrupo == 19)
				{
					this.txtDMDesde.MinValue = iMinimo;
					this.txtDMDesde.MaxValue = iMaximo;
					this.txtDMHasta.MinValue = iMinimo;
					this.txtDMHasta.MaxValue = iMaximo;

					this.txtDMDesde.Value = iMinimo;
					this.txtDMHasta.Value = iMaximo;
				}
				if ( iGrupo == 31 )
				{
					//					this.txtDMDesde.MinValue = 30;
					//					this.txtDMHasta.MinValue = 30;

					this.txtDMDesde.Value = iMinimo;
					this.txtDMHasta.Value = iMaximo;
				}
			}

			int iGrupos = 0;

			string sSqlMina = string.Format (" SELECT idGrupo FROM DBO.Usuario WHERE IDUSUARIO = {0}", MenuLatinium.IdUsuario ); 
			SqlDataReader drIda = Funcion.rEscalarSQL(cdsSeteoF,sSqlMina,true);
			drIda.Read();
			if(drIda.HasRows)
			{
				iGrupos = drIda.GetInt32 (0);
			}
			drIda.Close();

			int iPerOp = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [Valida_OperadorCobrador] {0}, {1},'{2}',2", 
				MenuLatinium.IdUsuario,iGrupos,Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd") ));

			int iHabilita = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [Valida_OperadorCobrador] {0}, {1},'{2}',1 ", 
				MenuLatinium.IdUsuario,iGrupos,Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd") ));


			this.cmbCobrador.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec [ConsultaCbo_OperadorCobradorPersonalComisionValida] {0}, {1},'{2}'", 
				MenuLatinium.IdUsuario,iGrupos,Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd") ));

			if (iHabilita == 0)
			{
				this.cmbCobrador.Enabled = false;
				this.cmbCobrador .Value = iPerOp;
			}
			#endregion OperadoCobrador			
		}


		private void ListaBodegas (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, int Usuario,int iPersonal, int iTipo, C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec ConsultaCbo_OperadorCobrador {0}, 4, {1}", MenuLatinium.IdUsuario, iPersonal));
		}

		private void cmbGestionCredito_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbGestionCredito.ActiveRow != null)
			{
			
				this.cmbGestion.Value = System.DBNull.Value;
				if((int)this.cmbGestionCredito.Value == 1)
				{
					this.cmbGestion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("select idCbo_ResultadoGestion, Resultado from Cbo_ResultadoGestion where idCbo_EstadoGestion = {0}", (int) this.cmbGestionCredito.Value));
				
				}
				if((int)this.cmbGestionCredito.Value == 2)
				{
					this.cmbGestion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("select idCbo_ResultadoGestion, Resultado from Cbo_ResultadoGestion where idCbo_EstadoGestion = {0}", (int) this.cmbGestionCredito.Value));
			
				}
				if((int)this.cmbGestionCredito.Value == 3)
				{
					this.cmbGestion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("select idCbo_ResultadoGestion, Resultado from Cbo_ResultadoGestion where idCbo_EstadoGestion =  {0}", (int) this.cmbGestionCredito.Value));

				}

			}
		}

		private void cmbGestion_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbGestion.ActiveRow != null)
			{

				#region Pariente Domicilio

				if((int)this.cmbGestion.Value == 5 || (int)this.cmbGestion.Value == 6 )
				{
				
					if((int)this.cmbGestion.Value == 5)
					{
						this.cmbParentesco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("SELECT idParentesco, Nombre FROM Cre_Parentesco where idParentesco not in (20,21)"));
					}
					else
					{
						this.cmbParentesco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("SELECT idParentesco, Nombre FROM Cre_Parentesco where idParentesco  in (20,21)"));
					}
				}
				
				#endregion Pariente Domicilio
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
		private void Saldo(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
			{
				
				dr.Cells["Saldo_Deuda_Actual"].Value =  Convert.ToDecimal(dr.Cells["Saldo_Deuda_Actual"].Value) -  Convert.ToDecimal(dr.Cells["Valor_Cobrado_Total"].Value) ;
			
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			
			int iMesM = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Month));
			int iYearM = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Year));

			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(( Convert.ToDateTime(this.dtFecha.Value).Year), ( Convert.ToDateTime(this.dtFecha.Value).Month)).ToString() + "/" + iMesM + "/" + iYearM );
			DateTime dtFecha = DateTime.Today ;
			int idOperador = 0;
			int idCobrador = 0;
			int idCbo_EstadoGestion = 0;
			int idCbo_ResultadoGestion = 0;
			int idEntidadFinanciera  = 0;
			int iBodega = 0;


			if (this.cmbCobrador.ActiveRow != null) idOperador = (int)this.cmbCobrador.Value;
			
			if (this.cmbGestionCredito.ActiveRow != null) idCbo_EstadoGestion = (int)this.cmbGestionCredito.Value;

			if (this.cmbGestion.ActiveRow != null) idCbo_ResultadoGestion = (int)this.cmbGestion.Value;

			if (this.cmbEntidadFinanciera.ActiveRow != null) idEntidadFinanciera = (int)this.cmbEntidadFinanciera.Value;

			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			if (idOperador > 0)
			{

				iGrupo = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" select u.idGrupo from IngresoCobrador n Inner join Usuario u on u.Nombre =  n.Codigo where n.idIngresoCobrador = {0}",
					(int) this.cmbCobrador.Value));	
         
			}

			if ((int) this.optGestion.Value > 0)
			{
				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un almacen")) return;
			}

			string sSQLPa = string.Format("Exec ConsultageneralCbo_GestorDeCobranzas '{0}', {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', {10}, '{11}', {12}, {13}", 
				Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"), (int)this.txtDMDesde.Value, (int) this.txtDMHasta.Value,(int)this.optAsignados.Value,
				idOperador,0, idCbo_EstadoGestion, idCbo_ResultadoGestion, idEntidadFinanciera, this.txtBuscar.Text.ToString (),0, 
				Convert.ToDateTime(dtFecha).ToString("yyyyMMdd"), (int) this.optGestion.Value, iBodega);
			FuncionesProcedimientos.dsGeneralModulo(sSQLPa, this.uGridInformacion);

			int  iOperadorCobrador = 0;

			if (iGrupo == 19)
			{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Ir"].Hidden = false;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cob"].Hidden = true;
				iOperadorCobrador = 1;
			}
			if (iGrupo == 31)
			{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Ir"].Hidden = true;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cob"].Hidden = false;
				iOperadorCobrador = 2;
			}

			if (iGrupo == 33)
			{
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Ir"].Hidden = true;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cob"].Hidden = false;
				iOperadorCobrador = 4;
			}      

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";

			#region valores 
			/*
			if (idOperador > 0)
			{
				string sSQLR = string.Format("Exec Cbo_RetornaValoresComision {0}, {1}, '{2}', {3},{4}", 
					idOperador, iOperadorCobrador, Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"), (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLR, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtProyectado.Value = dr.GetDecimal(0);
					this.txtCobrado.Value = dr.GetDecimal(1);
					this.txtPorcentaje.Value = dr.GetDecimal(2);
					this.txtBonificacion.Value = dr.GetDecimal(3);
					this.txtCobroExterno.Value = dr.GetDecimal(4);
				}
				dr.Close();
			}
			else
			{
				string sSQLRC = string.Format("Exec Cbo_RetornaValoresComision {0}, {1}, '{2}',{3},{4}", idOperador, iOperadorCobrador, 
					Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"), (int) this.txtDMDesde.Value, (int) this.txtDMHasta.Value );
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLRC, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtProyectado.Value = dr.GetDecimal(0);
					this.txtProyectado.Value = dr.GetDecimal(1);
					this.txtPorcentaje.Value = dr.GetDecimal(2);
					this.txtCobroExterno.Value = dr.GetDecimal(4);
					//this.txtBonificacion.Value = dr.GetDecimal(3);
				}
				dr.Close();
			}*/
			#endregion valores 


			decimal iProyectado = 0 ;
			decimal	iCobrado= 0;
			decimal dPorcentaje = 0;

			if ((int) this.uGridInformacion.Rows.Count > 0)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
				{
				
					iProyectado =  Convert.ToDecimal(dr.Cells["Valor_Cobrar_Proyectado"].Value) +  iProyectado ;
					iCobrado =  Convert.ToDecimal(dr.Cells["Valor_Cobrado_Total"].Value) +  iCobrado ;
				}
			}

			this.txtProyectado.Value = iProyectado;
			this.txtCobrado.Value = iCobrado ;

			if (iProyectado > 0)
			{
				this.txtPorcentaje.Value = (iCobrado * 100 ) / iProyectado;
			}
			else
			{
				this.txtPorcentaje.Value = 0;
			}

			this.Saldo(this.uGridInformacion);
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int idOrigen = 0;


			#region Operadora

			if (e.Cell.Column.ToString() == "Ir")
			{
				//				if (e.Cell.Row.Cells["Cartera_Asignada"].Value != System.DBNull.Value)
				//				{
				//					if ((int)e.Cell.Row.Cells["Cartera_Asignada"].Value > 0)
				//					{
				//						MessageBox.Show(string.Format("El cliente fue asignado a una gestion Externa "), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				////						return;
				//          
				//					}
				//				}


				if (e.Cell.Row.Cells["idOperador"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idOperador"].Value > 0)
					{
						using (frmCob_InformeNotCompro Gestion = new frmCob_InformeNotCompro((int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idOperador"].Value, 
										 e.Cell.Row.Cells["Cedula"].Value.ToString(), (int)e.Cell.Row.Cells["idEntidadFinanciera"].Value, 
										 e.Cell.Row.Cells["Numero_Documento"].Value.ToString(), e.Cell.Row.Cells["Cliente"].Value.ToString(), 
										 e.Cell.Row.Cells["Operadora"].Value.ToString(), 0, (int)e.Cell.Row.Cells["idCbo_GestorDeCobranzas"].Value))
						{
							if (DialogResult.OK == Gestion.ShowDialog( ))
							{

	
							}
						}
					}
					else
					{
						MessageBox.Show("Seleccione un Operador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
				else
				{
					MessageBox.Show("Seleccione un Operador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}

			#endregion Operadora

			#region Cobrador

			if (e.Cell.Column.ToString() == "Cob")
			{
				//				if (e.Cell.Row.Cells["Cartera_Asignada"].Value != System.DBNull.Value)
				//				{
				//					if ((int)e.Cell.Row.Cells["Cartera_Asignada"].Value > 0)
				//					{
				//						MessageBox.Show(string.Format("El cliente fue asignado a una gestion Externa "), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				////						return;
				//          
				//					}
				//				}

				if (e.Cell.Row.Cells["idOperador"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idCobrador"].Value > 0)
					{
						using (frmCob_InformeNotCompro Gestion = new frmCob_InformeNotCompro((int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idCobrador"].Value, 
										 e.Cell.Row.Cells["Cedula"].Value.ToString(), (int)e.Cell.Row.Cells["idEntidadFinanciera"].Value, 
										 e.Cell.Row.Cells["Numero_Documento"].Value.ToString(), e.Cell.Row.Cells["Cliente"].Value.ToString(), 
										 e.Cell.Row.Cells["Cobrador"].Value.ToString(), 1, (int)e.Cell.Row.Cells["idCbo_GestorDeCobranzas"].Value))
						{
							if (DialogResult.OK == Gestion.ShowDialog( ))
							{

	
							}
						}
					}
					else
					{
						MessageBox.Show("Seleccione un Cobrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
				else
				{
					MessageBox.Show("Seleccione un Cobrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}

			#endregion Cobrador

			#region Nota Credito

			if (e.Cell.Column.ToString() == "NC")
			{

				if (e.Cell.Row.Cells["idOperador"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idcompra"].Value > 0)
					{
						
						if ((decimal)e.Cell.Row.Cells["Valor_Retirado"].Value > 0)
						{
							idOrigen = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select idCompra from Compra where idOrigen = {0} and idTipoFactura = 5 and Estado not in (4)",
								e.Cell.Row.Cells["idCompra"].Value ));	

							using (frmDevolucionVenta NC = new frmDevolucionVenta(idOrigen))
							{
								if (DialogResult.OK == NC.ShowDialog( ))
								{

	
								}
							}
						}
					}
				}
				else
				{
					MessageBox.Show("Seleccione un Cobrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}

			#endregion Nota Credito

			#region Historico

			if (e.Cell.Column.ToString() == "Historico")
			{
			
				using (frmGestionCobranza Gestion = new frmGestionCobranza((int)e.Cell.Row.Cells["idCompra"].Value, false, false))
				{
					if (DialogResult.OK == Gestion.ShowDialog())
					{
								
					}
				}
			}
			#endregion Historico
		}

		private void optAsignados_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click (sender , e );
		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void cmbCobrador_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbGestionCredito.Focus();
		}

		private void cmbGestionCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbGestion.Focus();
		}

		private void cmbGestion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbEntidadFinanciera.Focus();
		}

		private void cmbEntidadFinanciera_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnVer_Click (sender, e);
		}

		bool bActualiza = false;

	
		private void uGridInformacion_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;
			if (e.Cell.Column.ToString() == "idCobrador")
			{		
				bActualiza = true;
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Vendedor?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE Cbo_GestorDeCobranzas SET idCobrador = {0} WHERE idCbo_GestorDeCobranzas = {1} ", 
						(int)	this.uGridInformacion.ActiveRow.Cells["idCobrador"].Value, (int)	this.uGridInformacion.ActiveRow.Cells["idCbo_GestorDeCobranzas"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);		
				
					string sSQLCb = string.Format("UPDATE Compra SET idCobrador = {0} WHERE idCompra = {1} ", 
						(int)	this.uGridInformacion.ActiveRow.Cells["idCobrador"].Value, (int)	this.uGridInformacion.ActiveRow.Cells["idCompra"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQLCb);	

				}
				else
				{						
					e.Cell.Row.Cells["idCobrador"].Value = (int)e.Cell.OriginalValue;
					//					return;
				}
				bActualiza = false;
			}
		}

		private void frmConsultaCbo_GestorDeCobranzas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			int idOperadorCobrador  = 0;

			if (this.cmbCobrador.ActiveRow != null) idOperadorCobrador = (int)this.cmbCobrador.Value;

			using (frmPorcentajes FpC = new frmPorcentajes (idOperadorCobrador, (DateTime)this.dtFecha.Value))
			{
				if (DialogResult.OK == FpC.ShowDialog())
				{
						
				}
			}
		}

		private void uGridInformacion_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if ((int)e.Row.Cells["IEmpleado"].Value > 0)
			{
				e.Row.Appearance.BackColor = Color.Bisque;
			}
		}

		private void optGestion_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int) this.optGestion.Value == 0 )
			{
				if ( iGrupo == 19 )
				{
					this.txtDMDesde.MinValue = iMinimo;
					this.txtDMDesde.MaxValue = iMaximo;
					this.txtDMHasta.MinValue = iMinimo;
					this.txtDMHasta.MaxValue = iMaximo;

					this.txtDMDesde.Value = iMinimo;
					this.txtDMHasta.Value = iMaximo;
				}
				if ( iGrupo == 31 )
				{
					this.txtDMDesde.Value = iMinimo;
					this.txtDMHasta.Value = iMaximo;

				}

				this.txtDMDesde.MinValue = iMinimo;
				this.txtDMDesde.MaxValue = iMaximo;
				
			}
			else
			{
				this.txtDMDesde.MinValue = 0;
				this.txtDMDesde.MaxValue = 5000;
				this.txtDMHasta.MinValue = 0;
				this.txtDMHasta.MaxValue = 5000;

				this.txtDMDesde.Value = 0;
				this.txtDMHasta.Value = 30;

			}

		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{
				int iMesM = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Month));
				int iYearM = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Year));

				DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(( Convert.ToDateTime(this.dtFecha.Value).Year), ( Convert.ToDateTime(this.dtFecha.Value).Month)).ToString() + "/" + iMesM + "/" + iYearM );
			  
				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("select Bodega, Nombre from dbo.bodega b where exists (select bodega from dbo.Cbo_GestorDeCobranzas where Bodega = b.Bodega and Periodo ='{0}' )order by nombre",Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")));

				this.cmbCobradores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("Exec ListadoCobranzas {0}, '{1}'", 2, Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")));
				this.cmbSupervisor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("SELECT n.idIngresoCobrador as idCobrador, (n.Nombre) as Nombre from dbo.Usuario u Inner join dbo.IngresoCobrador n on n.codigo = u.Nombre where u.idGrupo = 33"));
		    
				this.cmbOperadora.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("Exec ListadoCobranzas {0}, '{1}'", 1, Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")));

				this.cmbEstadosGestio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("Exec ListadoCobranzas {0}, '{1}'", 3, Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")));
				this.cmbResultadoGestion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("Exec ListadoCobranzas {0}, '{1}'", 4, Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")));
				this.cmbZona.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("Exec ListadoCobranzas {0}, '{1}'", 5, Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")));
			}
		}

		private void uGridInformacion_ClickCellButton_1(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			int idOrigen = 0;


			#region Operadora

			if (e.Cell.Column.ToString() == "Ir")
			{
				//				if (e.Cell.Row.Cells["Cartera_Asignada"].Value != System.DBNull.Value)
				//				{
				//					if ((int)e.Cell.Row.Cells["Cartera_Asignada"].Value > 0)
				//					{
				//						MessageBox.Show(string.Format("El cliente fue asignado a una gestion Externa "), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				////						return;
				//          
				//					}
				//				}


				if (e.Cell.Row.Cells["idOperador"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idOperador"].Value > 0)
					{
						using (frmCob_InformeNotCompro Gestion = new frmCob_InformeNotCompro((int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idOperador"].Value, 
										 e.Cell.Row.Cells["Cedula"].Value.ToString(), (int)e.Cell.Row.Cells["idEntidadFinanciera"].Value, 
										 e.Cell.Row.Cells["Numero_Documento"].Value.ToString(), e.Cell.Row.Cells["Cliente"].Value.ToString(), 
										 "", 0, (int)e.Cell.Row.Cells["idCbo_GestorDeCobranzas"].Value))
						{
							if (DialogResult.OK == Gestion.ShowDialog( ))
							{

	
							}
						}
					}
					else
					{
						MessageBox.Show("Seleccione un Operador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
				else
				{
					MessageBox.Show("Seleccione un Operador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}

			#endregion Operadora

			#region Cobrador

			if (e.Cell.Column.ToString() == "Cob")
			{
				//				if (e.Cell.Row.Cells["Cartera_Asignada"].Value != System.DBNull.Value)
				//				{
				//					if ((int)e.Cell.Row.Cells["Cartera_Asignada"].Value > 0)
				//					{
				//						MessageBox.Show(string.Format("El cliente fue asignado a una gestion Externa "), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				////						return;
				//          
				//					}
				//				}

				if (e.Cell.Row.Cells["idOperador"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idCobrador"].Value > 0)
					{
						using (frmCob_InformeNotCompro Gestion = new frmCob_InformeNotCompro((int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idCobrador"].Value, 
										 e.Cell.Row.Cells["Cedula"].Value.ToString(), (int)e.Cell.Row.Cells["idEntidadFinanciera"].Value, 
										 e.Cell.Row.Cells["Numero_Documento"].Value.ToString(), e.Cell.Row.Cells["Cliente"].Value.ToString(), 
										 "", 1, (int)e.Cell.Row.Cells["idCbo_GestorDeCobranzas"].Value))
						{
							if (DialogResult.OK == Gestion.ShowDialog( ))
							{

	
							}
						}
					}
					else
					{
						MessageBox.Show("Seleccione un Cobrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
				else
				{
					MessageBox.Show("Seleccione un Cobrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}

			#endregion Cobrador

			#region Nota Credito

			if (e.Cell.Column.ToString() == "NC")
			{

				if (e.Cell.Row.Cells["idOperador"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idcompra"].Value > 0)
					{
						
						if ((decimal)e.Cell.Row.Cells["Valor_Retirado"].Value > 0)
						{
							idOrigen = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select idCompra from Compra where idOrigen = {0} and idTipoFactura = 5 and Estado not in (4)",
								e.Cell.Row.Cells["idCompra"].Value ));	

							using (frmDevolucionVenta NC = new frmDevolucionVenta(idOrigen))
							{
								if (DialogResult.OK == NC.ShowDialog( ))
								{

	
								}
							}
						}
					}
				}
				else
				{
					MessageBox.Show("Seleccione un Cobrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}

			#endregion Nota Credito

			#region Historico

			if (e.Cell.Column.ToString() == "Historico")
			{
			
				using (frmGestionCobranza Gestion = new frmGestionCobranza((int)e.Cell.Row.Cells["idCompra"].Value, false, false))
				{
					if (DialogResult.OK == Gestion.ShowDialog())
					{
								
					}
				}
			}
			#endregion Historico

			#region Historico

			if (e.Cell.Column.ToString() == "btnEquifax")
			{
				int iEquifax = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(1) from coordenadasequifax where Identificacion ='{0}'",e.Cell.Row.Cells["Cedula"].Value), true);
				if( iEquifax > 0)
				{
					string sUrlNube = Funcion.sEscalarSQL(cdsSeteoF,string.Format("EXEC [RetornaURLCoordanadasEquifax] '{0}','{1}'",e.Cell.Row.Cells["Cedula"].Value, e.Cell.Row.Cells["Cliente"].Value)); 

					if (sUrlNube.Length > 10)
					{
						System.Diagnostics.Process.Start("chrome.exe", sUrlNube);				
					}
				}
				else
				{
					MessageBox.Show("Cliente no cuenta con Información.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
						return;
				}
			}
			#endregion Historico
		}

		private void uGridInformacion_AfterCellUpdate_1(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;
			if (e.Cell.Column.ToString() == "idCobrador")
			{		
				bActualiza = true;
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Vendedor?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE Cbo_GestorDeCobranzas SET idCobrador = {0} WHERE idCbo_GestorDeCobranzas = {1} ", 
						(int)	this.uGridInformacion.ActiveRow.Cells["idCobrador"].Value, (int)	this.uGridInformacion.ActiveRow.Cells["idCbo_GestorDeCobranzas"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);		
				
					string sSQLCb = string.Format("UPDATE Compra SET idCobrador = {0} WHERE idCompra = {1} ", 
						(int)	this.uGridInformacion.ActiveRow.Cells["idCobrador"].Value, (int)	this.uGridInformacion.ActiveRow.Cells["idCompra"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQLCb);	

				}
				else
				{						
					e.Cell.Row.Cells["idCobrador"].Value = (int)e.Cell.OriginalValue;
					//					return;
				}
				bActualiza = false;
			}
		}

		private void uGridInformacion_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			string sSQLPa = string.Format("Exec ConsultaGeneralBarriosAsignacion {0}, {1}, {2}, {3}, {4}" ,0, 0,0, 0,0 );
			this.ultgconsulta.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);

			FuncionesProcedimientos.ExportaExcel(this.ultgconsulta);

		}
		public static string GetValueOrDefault(DataRow row, string columnName)
		{
			return row[columnName] != DBNull.Value ? row[columnName].ToString() : "";
		}

		public static int GetIntValueOrDefault(DataRow row, string columnName)
		{
			return row[columnName] != DBNull.Value ? Convert.ToInt32(row[columnName]) : 0;
		}

		public static decimal GetDecimalValueOrDefault(DataRow row, string columnName)
		{
			return row[columnName] != DBNull.Value ? Convert.ToDecimal(row[columnName]) : 0.00m;
		}
	

		private void btnImportar_Click(object sender, System.EventArgs e)
		{
			int iMesM = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Month));
			int iYearM = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Year));

			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(( Convert.ToDateTime(this.dtFecha.Value).Year), ( Convert.ToDateTime(this.dtFecha.Value).Month)).ToString() + "/" + iMesM + "/" + iYearM );
			
			if (DialogResult.Yes == MessageBox.Show(string.Format("¿Esta seguro de actualizar la cartera {0}?",dtFinalPeriodo), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{	
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";
			int iValida = 0;
				string Scobrador ="";

				if (sArchivo.ShowDialog() == DialogResult.OK && sArchivo.FileName != "")
				{
					try
					{
						this.Cursor = Cursors.WaitCursor;

						DataTable politicasTable = FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT idCompra1,Cobrador,Zona FROM [POLITICAS$]", "POLITICAS").Tables[0];
																				
						foreach (DataRow row in politicasTable.Rows)
						{
							int idCompra1 = GetIntValueOrDefault(row, "idCompra1");
							string Cobrador = GetValueOrDefault(row, "Cobrador");
							int Zona = GetIntValueOrDefault(row, "Zona");
							
							#region Guardar

							using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
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

									#region Graba DetalleMaestro
								
									string  sSQLDet = string.Format("Exec [GeneraCarteraZona] {0}, '{1}', {2}, '{3}'" ,
										idCompra1, Cobrador, Zona, Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")
										
										);
									oCmdActualiza.CommandText = sSQLDet;
									
									iValida = (int)oCmdActualiza.ExecuteScalar();
									if( iValida == 0 )	Scobrador = Scobrador + " , "+Cobrador;
								
									#endregion Graba DetalleMaestro
					
									oTransaction.Commit();

									Cursor = Cursors.Default;	

									
						
								}
								catch(Exception ex)
								{
									try
									{ 
										oTransaction.Rollback();

										MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
									}
									catch (Exception ex2)
									{
										MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
									}
								}
								finally
								{
									oConexion.Close();
								}
							}
				
							#endregion Guardar
						}

						MessageBox.Show("Actualización de cartera generado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
						if( Scobrador.Length > 2 )
						{
             MessageBox.Show(string.Format("Los siguientes cobradores no existen: /n {0}", Scobrador ), "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
						}

					}
					catch (Exception ex)
					{
						MessageBox.Show("Se produjo un error inesperado: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					finally
					{
						this.Cursor = Cursors.Default;
						sArchivo.Dispose();
					}
				}
			}
		}

		private void btnImpOperador_Click(object sender, System.EventArgs e)
		{
			int iMesM = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Month));
			int iYearM = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Year));
			ArrayList cobradoresExistentes = new ArrayList();
			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(( Convert.ToDateTime(this.dtFecha.Value).Year), ( Convert.ToDateTime(this.dtFecha.Value).Month)).ToString() + "/" + iMesM + "/" + iYearM );
			int iRegistros = 0;
			int iregistrosValidos = 0;
			if (DialogResult.Yes == MessageBox.Show(string.Format("¿Esta seguro de actualizar la cartera {0}, Operadoras?",dtFinalPeriodo), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{	
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == DialogResult.OK && sArchivo.FileName != "")
				{
					try
					{
						this.Cursor = Cursors.WaitCursor;

						DataTable politicasTable = FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT idCompra1,Operador FROM [OPERADOR$]", "OPERADOR").Tables[0];

						foreach (DataRow row in politicasTable.Rows)
						{
							int idCompra1 = GetIntValueOrDefault(row, "idCompra1");
							string Operador = GetValueOrDefault(row, "Operador");
							
							#region Guardar

							using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
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

									#region Graba DetalleMaestro

									iRegistros ++;								  
									string  sSQLDet = string.Format("Exec [GeneraCarteraZonaOperador] {0}, '{1}', '{2}'" ,
										idCompra1, Operador,  Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")
										);

									oCmdActualiza.CommandText = sSQLDet;
									int iCount  = (int)oCmdActualiza.ExecuteScalar();
									if (iCount == 0) // Si iCount es igual a 0, es porque no se actualizaron registros
									{
										// Guardamos el Cobrador en el array solo si no se procesaron registros válidos (iCount == 0)
										if (!cobradoresExistentes.Contains(Operador))
										{
											cobradoresExistentes.Add(Operador);
										}
									}
									else
									{
										iregistrosValidos += iCount; // Si hay registros actualizados, incrementamos los registros válidos
									}

								
									#endregion Graba DetalleMaestro
					
									oTransaction.Commit();

									Cursor = Cursors.Default;	
						
								}
								catch(Exception ex)
								{
									try
									{ 
										oTransaction.Rollback();

										MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
									}
									catch (Exception ex2)
									{
										MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
									}
								}
								finally
								{
									oConexion.Close();
								}
							}
				
							#endregion Guardar
						}
						// Después del ciclo que procesa los datos:

						MessageBox.Show(string.Format(
							"Total Registros: '{0}'\nRegistros Actualizados: '{1}'\nCobradores No Procesados: {2}",
							iRegistros, 
							iregistrosValidos, 
							cobradoresExistentes.Count > 0 ? string.Join(", ", (string[])cobradoresExistentes.ToArray(typeof(string))) : "Ninguno"
							), "Validación", MessageBoxButtons.OK, MessageBoxIcon.Information);

					}
					catch (Exception ex)
					{
						MessageBox.Show("Se produjo un error inesperado: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					finally
					{
						this.Cursor = Cursors.Default;
						sArchivo.Dispose();
					}
				}
			}
		}
	}
}
