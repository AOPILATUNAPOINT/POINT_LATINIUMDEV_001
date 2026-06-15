using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteAnalisiEquiFax.
	/// </summary>
	public class frmReporteAnalisiEquiFax : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridConsulta;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVariableNinelIdent;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVariableOperacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesCopiarUltimoMes;
		private System.Windows.Forms.MenuItem mnuAccionesAgregarNuevoLocal;
		private System.Windows.Forms.MenuItem menuItem1;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteAnalisiEquiFax()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteAnalisiEquiFax));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_De_Consulta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplica_NoAplica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Otorgado");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_De_Consulta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aplica_NoAplica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Otorgado");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Identificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("est_civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fecha_nacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("canton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("parroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_cargas");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("marca_rechazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dependiente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipo_vivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("profesion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("codigo_act_nueva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nombre_act_nueva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sueldo");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("INGRESOSPERSONAL");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("INGRESOSNEGOCIO");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EGRESOSPERSONAL");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EGRESOSNEGOCIO");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ACTIVOS");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PASIVOS");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("antigüedad_micro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("antigüedad_Laboral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo_ahorro");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo_Prom_6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen_de_Ingresos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Apertura_de_Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Número_de_Codeudas_Otorgadas_Previas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Número_de_Garantias_Otorgadas_Previas");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Identificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("est_civil");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sexo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fecha_nacimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("provincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("canton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("parroquia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_cargas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("marca_rechazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dependiente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipo_vivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("educacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("profesion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("codigo_act_nueva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("nombre_act_nueva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("INGRESOSPERSONAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("INGRESOSNEGOCIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EGRESOSPERSONAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EGRESOSNEGOCIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ACTIVOS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PASIVOS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("antigüedad_micro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("antigüedad_Laboral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo_ahorro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo_Prom_6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen_de_Ingresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Apertura_de_Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Número_de_Codeudas_Otorgadas_Previas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Número_de_Garantias_Otorgadas_Previas");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Identificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("fecha_otorgamiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Num_Ope");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("marca_reestructuracion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipo_garantia_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VALOR_GARANTIA_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipo_garantia_2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VALOR_GARANTIA_2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipo_garantia_3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VALOR_GARANTIA3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("marca_excepción");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipo_credito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("tipo_producto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Plazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("destino_Credito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_de_Operación");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DEUDAINICIAL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NUMEROCUOTAS");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NUMEROCUOTAS_pag");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FRECUENCIAPAGO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n0");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n0");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n0");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n0");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n0");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n0");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n0");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n1");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n1");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n1");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n1");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n1");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n1");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n1");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n2");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n2");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n2");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n2");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n2");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n2");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n2");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n3");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n3");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n3");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n3");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n3");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n3");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n3");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n4");
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n4");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n4");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n4");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n4");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n4");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n4");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n5");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n5");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n5");
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n5");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n5");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n5");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n5");
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n6");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n6");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n6");
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n6");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n6");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n6");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n6");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n7");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n7");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n7");
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n7");
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n7");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n7");
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n7");
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n8");
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n8");
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n8");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n8");
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n8");
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n8");
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n8");
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n9");
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n9");
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n9");
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n9");
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n9");
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n9");
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n9");
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n10");
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n10");
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n10");
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n10");
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n10");
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n10");
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n10");
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n11");
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n11");
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n11");
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n11");
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n11");
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n11");
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n11");
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_deuda_n12");
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_xvencer_n12");
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_ndi_n12");
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_vencido_n12");
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_dema_n12");
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sal_cast_n12");
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("num_dven_n12");
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Identificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("fecha_otorgamiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Num_Ope");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("marca_reestructuracion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipo_garantia_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VALOR_GARANTIA_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipo_garantia_2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VALOR_GARANTIA_2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipo_garantia_3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VALOR_GARANTIA3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("marca_excepción");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipo_credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("tipo_producto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Plazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("destino_Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado_de_Operación");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DEUDAINICIAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NUMEROCUOTAS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NUMEROCUOTAS_pag");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FRECUENCIAPAGO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn98 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn99 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn100 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn101 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn102 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn103 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn104 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn105 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn106 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn107 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn108 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn109 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn110 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn111 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn112 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn113 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn114 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn115 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn116 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn117 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn118 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn119 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn120 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn121 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn122 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn123 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn124 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn125 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn126 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn127 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn128 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn129 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn130 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn131 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn132 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn133 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn134 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn135 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn136 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn137 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn138 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn139 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn140 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_deuda_n12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn141 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_xvencer_n12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn142 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_ndi_n12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn143 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_vencido_n12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn144 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_dema_n12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn145 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sal_cast_n12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn146 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("num_dven_n12");
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.uGridConsulta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnVer = new System.Windows.Forms.Button();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridVariableNinelIdent = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridVariableOperacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnComandos = new System.Windows.Forms.Button();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesCopiarUltimoMes = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAgregarNuevoLocal = new System.Windows.Forms.MenuItem();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVariableNinelIdent)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVariableOperacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(208, 48);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 322;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(24, 48);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 321;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 46);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 320;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(80, 46);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 319;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// optEstados
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance3;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance4;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Consulta Equifax";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Variables Nivel Identificación";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Variable pto Operación";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(24, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(448, 29);
			this.optEstados.TabIndex = 323;
			this.optEstados.Text = "Consulta Equifax";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// uGridConsulta
			// 
			this.uGridConsulta.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridConsulta.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridConsulta.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			this.uGridConsulta.DisplayLayout.Appearance = appearance5;
			ultraGridColumn1.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.Caption = "Cédula";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 150;
			ultraGridColumn2.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Fecha de Consulta";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 120;
			ultraGridColumn3.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "Aplica /No Aplica";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 110;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance6;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.uGridConsulta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridConsulta.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridConsulta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridConsulta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridConsulta.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridConsulta.DisplayLayout.Override.CellAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridConsulta.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.uGridConsulta.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConsulta.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridConsulta.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.uGridConsulta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridConsulta.Location = new System.Drawing.Point(8, 96);
			this.uGridConsulta.Name = "uGridConsulta";
			this.uGridConsulta.Size = new System.Drawing.Size(520, 376);
			this.uGridConsulta.TabIndex = 1062;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(-152, 80);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1578, 8);
			this.groupBox2.TabIndex = 1063;
			this.groupBox2.TabStop = false;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(392, 44);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(80, 24);
			this.btnVer.TabIndex = 1064;
			this.btnVer.Text = "&Consultar";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// uGridVariableNinelIdent
			// 
			this.uGridVariableNinelIdent.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVariableNinelIdent.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVariableNinelIdent.DataSource = this.ultraDataSource2;
			appearance12.BackColor = System.Drawing.Color.White;
			this.uGridVariableNinelIdent.DisplayLayout.Appearance = appearance12;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Identificación";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 60;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 60;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance13;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Width = 30;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 60;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Width = 40;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn15.Width = 110;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 12;
			ultraGridColumn18.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.Width = 60;
			ultraGridColumn19.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 14;
			ultraGridColumn19.Width = 60;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance14;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 15;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance15;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 16;
			ultraGridColumn21.Width = 80;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance16;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 17;
			ultraGridColumn22.Width = 80;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance17;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 18;
			ultraGridColumn23.Width = 80;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance18;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 19;
			ultraGridColumn24.Width = 80;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance19;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 20;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance20;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 21;
			ultraGridColumn27.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 22;
			ultraGridColumn28.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 23;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance21;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 24;
			ultraGridColumn30.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 25;
			ultraGridColumn31.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.Caption = "Origen de Ingresos";
			ultraGridColumn31.Header.VisiblePosition = 26;
			ultraGridColumn32.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.Caption = "Fecha Apertura de Cuenta";
			ultraGridColumn32.Header.VisiblePosition = 27;
			ultraGridColumn33.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.Caption = "Número de Codeudas Otorgadas_Previas";
			ultraGridColumn33.Header.VisiblePosition = 28;
			ultraGridColumn34.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn34.Header.Caption = "Número de Garantias Otorgadas_Previas";
			ultraGridColumn34.Header.VisiblePosition = 29;
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
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.uGridVariableNinelIdent.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridVariableNinelIdent.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVariableNinelIdent.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVariableNinelIdent.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance22.BackColor = System.Drawing.Color.Transparent;
			this.uGridVariableNinelIdent.DisplayLayout.Override.CardAreaAppearance = appearance22;
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVariableNinelIdent.DisplayLayout.Override.CellAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance24.FontData.BoldAsString = "True";
			appearance24.FontData.Name = "Arial";
			appearance24.FontData.SizeInPoints = 8F;
			appearance24.ForeColor = System.Drawing.Color.White;
			appearance24.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVariableNinelIdent.DisplayLayout.Override.HeaderAppearance = appearance24;
			this.uGridVariableNinelIdent.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVariableNinelIdent.DisplayLayout.Override.RowSelectorAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVariableNinelIdent.DisplayLayout.Override.SelectedRowAppearance = appearance26;
			this.uGridVariableNinelIdent.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVariableNinelIdent.Location = new System.Drawing.Point(8, 96);
			this.uGridVariableNinelIdent.Name = "uGridVariableNinelIdent";
			this.uGridVariableNinelIdent.Size = new System.Drawing.Size(1288, 376);
			this.uGridVariableNinelIdent.TabIndex = 1065;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn34});
			// 
			// uGridVariableOperacion
			// 
			this.uGridVariableOperacion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVariableOperacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVariableOperacion.DataSource = this.ultraDataSource3;
			appearance27.BackColor = System.Drawing.Color.White;
			this.uGridVariableOperacion.DisplayLayout.Appearance = appearance27;
			ultraGridColumn35.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn35.Header.Caption = "Identificación";
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Width = 100;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn37.Header.VisiblePosition = 2;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Header.VisiblePosition = 3;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn39.Header.VisiblePosition = 4;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn40.Header.VisiblePosition = 5;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn41.Header.VisiblePosition = 6;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn42.Header.VisiblePosition = 7;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn43.Header.VisiblePosition = 8;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn44.Header.VisiblePosition = 9;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn45.Header.VisiblePosition = 10;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn46.Header.VisiblePosition = 11;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn47.Header.VisiblePosition = 12;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn48.Header.VisiblePosition = 13;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn49.Header.VisiblePosition = 14;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn50.Header.VisiblePosition = 15;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn51.Header.VisiblePosition = 16;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn52.Header.VisiblePosition = 17;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn53.Header.VisiblePosition = 18;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn54.Header.VisiblePosition = 19;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn55.Header.VisiblePosition = 20;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance28;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.VisiblePosition = 21;
			ultraGridColumn56.Width = 80;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn57.CellAppearance = appearance29;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.VisiblePosition = 22;
			ultraGridColumn57.Width = 80;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance30;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.VisiblePosition = 23;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance31;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.VisiblePosition = 24;
			ultraGridColumn59.Width = 80;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance32;
			ultraGridColumn60.Format = "#,##0.00";
			ultraGridColumn60.Header.VisiblePosition = 25;
			ultraGridColumn60.Width = 80;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance33;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.VisiblePosition = 26;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn62.CellAppearance = appearance34;
			ultraGridColumn62.Header.VisiblePosition = 27;
			ultraGridColumn62.Width = 80;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance35;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.VisiblePosition = 28;
			ultraGridColumn63.Width = 80;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance36;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.VisiblePosition = 29;
			ultraGridColumn64.Width = 80;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance37;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.VisiblePosition = 30;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance38;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.VisiblePosition = 31;
			ultraGridColumn66.Width = 80;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance39;
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.VisiblePosition = 32;
			ultraGridColumn67.Width = 80;
			ultraGridColumn68.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance40;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.VisiblePosition = 33;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn69.CellAppearance = appearance41;
			ultraGridColumn69.Header.VisiblePosition = 34;
			ultraGridColumn69.Width = 80;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance42;
			ultraGridColumn70.Format = "#,##0.00";
			ultraGridColumn70.Header.VisiblePosition = 35;
			ultraGridColumn70.Width = 80;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance43;
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.VisiblePosition = 36;
			ultraGridColumn71.Width = 80;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance44;
			ultraGridColumn72.Format = "#,##0.00";
			ultraGridColumn72.Header.VisiblePosition = 37;
			ultraGridColumn73.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance45;
			ultraGridColumn73.Format = "#,##0.00";
			ultraGridColumn73.Header.VisiblePosition = 38;
			ultraGridColumn73.Width = 80;
			ultraGridColumn74.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance46;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.VisiblePosition = 39;
			ultraGridColumn74.Width = 80;
			ultraGridColumn75.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance47;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.VisiblePosition = 40;
			ultraGridColumn76.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn76.CellAppearance = appearance48;
			ultraGridColumn76.Header.VisiblePosition = 41;
			ultraGridColumn76.Width = 80;
			ultraGridColumn77.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance49;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.VisiblePosition = 42;
			ultraGridColumn77.Width = 80;
			ultraGridColumn78.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance50;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.VisiblePosition = 43;
			ultraGridColumn78.Width = 80;
			ultraGridColumn79.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance51;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.VisiblePosition = 44;
			ultraGridColumn80.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance52;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.VisiblePosition = 45;
			ultraGridColumn80.Width = 80;
			ultraGridColumn81.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn81.CellAppearance = appearance53;
			ultraGridColumn81.Format = "#,##0.00";
			ultraGridColumn81.Header.VisiblePosition = 46;
			ultraGridColumn81.Width = 80;
			ultraGridColumn82.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn82.CellAppearance = appearance54;
			ultraGridColumn82.Format = "#,##0.00";
			ultraGridColumn82.Header.VisiblePosition = 47;
			ultraGridColumn83.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn83.CellAppearance = appearance55;
			ultraGridColumn83.Header.VisiblePosition = 48;
			ultraGridColumn83.Width = 80;
			ultraGridColumn84.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn84.CellAppearance = appearance56;
			ultraGridColumn84.Format = "#,##0.00";
			ultraGridColumn84.Header.VisiblePosition = 49;
			ultraGridColumn84.Width = 80;
			ultraGridColumn85.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn85.CellAppearance = appearance57;
			ultraGridColumn85.Format = "#,##0.00";
			ultraGridColumn85.Header.VisiblePosition = 50;
			ultraGridColumn85.Width = 80;
			ultraGridColumn86.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn86.CellAppearance = appearance58;
			ultraGridColumn86.Format = "#,##0.00";
			ultraGridColumn86.Header.VisiblePosition = 51;
			ultraGridColumn87.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn87.CellAppearance = appearance59;
			ultraGridColumn87.Format = "#,##0.00";
			ultraGridColumn87.Header.VisiblePosition = 52;
			ultraGridColumn87.Width = 80;
			ultraGridColumn88.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn88.CellAppearance = appearance60;
			ultraGridColumn88.Format = "#,##0.00";
			ultraGridColumn88.Header.VisiblePosition = 53;
			ultraGridColumn88.Width = 80;
			ultraGridColumn89.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn89.CellAppearance = appearance61;
			ultraGridColumn89.Format = "#,##0.00";
			ultraGridColumn89.Header.VisiblePosition = 54;
			ultraGridColumn90.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn90.CellAppearance = appearance62;
			ultraGridColumn90.Header.VisiblePosition = 55;
			ultraGridColumn90.Width = 80;
			ultraGridColumn91.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn91.CellAppearance = appearance63;
			ultraGridColumn91.Format = "#,##0.00";
			ultraGridColumn91.Header.VisiblePosition = 56;
			ultraGridColumn91.Width = 80;
			ultraGridColumn92.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn92.CellAppearance = appearance64;
			ultraGridColumn92.Format = "#,##0.00";
			ultraGridColumn92.Header.VisiblePosition = 57;
			ultraGridColumn92.Width = 80;
			ultraGridColumn93.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn93.CellAppearance = appearance65;
			ultraGridColumn93.Format = "#,##0.00";
			ultraGridColumn93.Header.VisiblePosition = 58;
			ultraGridColumn94.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn94.CellAppearance = appearance66;
			ultraGridColumn94.Format = "#,##0.00";
			ultraGridColumn94.Header.VisiblePosition = 59;
			ultraGridColumn94.Width = 80;
			ultraGridColumn95.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn95.CellAppearance = appearance67;
			ultraGridColumn95.Format = "#,##0.00";
			ultraGridColumn95.Header.VisiblePosition = 60;
			ultraGridColumn95.Width = 80;
			ultraGridColumn96.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn96.CellAppearance = appearance68;
			ultraGridColumn96.Format = "#,##0.00";
			ultraGridColumn96.Header.VisiblePosition = 61;
			ultraGridColumn97.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance69.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn97.CellAppearance = appearance69;
			ultraGridColumn97.Header.VisiblePosition = 62;
			ultraGridColumn97.Width = 80;
			ultraGridColumn98.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn98.CellAppearance = appearance70;
			ultraGridColumn98.Format = "#,##0.00";
			ultraGridColumn98.Header.VisiblePosition = 63;
			ultraGridColumn98.Width = 80;
			ultraGridColumn99.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn99.CellAppearance = appearance71;
			ultraGridColumn99.Format = "#,##0.00";
			ultraGridColumn99.Header.VisiblePosition = 64;
			ultraGridColumn99.Width = 80;
			ultraGridColumn100.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn100.CellAppearance = appearance72;
			ultraGridColumn100.Format = "#,##0.00";
			ultraGridColumn100.Header.VisiblePosition = 65;
			ultraGridColumn101.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance73.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn101.CellAppearance = appearance73;
			ultraGridColumn101.Format = "#,##0.00";
			ultraGridColumn101.Header.VisiblePosition = 66;
			ultraGridColumn101.Width = 80;
			ultraGridColumn102.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn102.CellAppearance = appearance74;
			ultraGridColumn102.Format = "#,##0.00";
			ultraGridColumn102.Header.VisiblePosition = 67;
			ultraGridColumn102.Width = 80;
			ultraGridColumn103.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn103.CellAppearance = appearance75;
			ultraGridColumn103.Format = "#,##0.00";
			ultraGridColumn103.Header.VisiblePosition = 68;
			ultraGridColumn104.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn104.CellAppearance = appearance76;
			ultraGridColumn104.Header.VisiblePosition = 69;
			ultraGridColumn104.Width = 80;
			ultraGridColumn105.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn105.CellAppearance = appearance77;
			ultraGridColumn105.Format = "#,##0.00";
			ultraGridColumn105.Header.VisiblePosition = 70;
			ultraGridColumn105.Width = 80;
			ultraGridColumn106.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn106.CellAppearance = appearance78;
			ultraGridColumn106.Format = "#,##0.00";
			ultraGridColumn106.Header.VisiblePosition = 71;
			ultraGridColumn106.Width = 80;
			ultraGridColumn107.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn107.CellAppearance = appearance79;
			ultraGridColumn107.Format = "#,##0.00";
			ultraGridColumn107.Header.VisiblePosition = 72;
			ultraGridColumn108.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn108.CellAppearance = appearance80;
			ultraGridColumn108.Format = "#,##0.00";
			ultraGridColumn108.Header.VisiblePosition = 73;
			ultraGridColumn108.Width = 80;
			ultraGridColumn109.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn109.CellAppearance = appearance81;
			ultraGridColumn109.Format = "#,##0.00";
			ultraGridColumn109.Header.VisiblePosition = 74;
			ultraGridColumn109.Width = 80;
			ultraGridColumn110.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn110.CellAppearance = appearance82;
			ultraGridColumn110.Format = "#,##0.00";
			ultraGridColumn110.Header.VisiblePosition = 75;
			ultraGridColumn111.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn111.CellAppearance = appearance83;
			ultraGridColumn111.Header.VisiblePosition = 76;
			ultraGridColumn111.Width = 80;
			ultraGridColumn112.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn112.CellAppearance = appearance84;
			ultraGridColumn112.Format = "#,##0.00";
			ultraGridColumn112.Header.VisiblePosition = 77;
			ultraGridColumn112.Width = 80;
			ultraGridColumn113.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn113.CellAppearance = appearance85;
			ultraGridColumn113.Format = "#,##0.00";
			ultraGridColumn113.Header.VisiblePosition = 78;
			ultraGridColumn113.Width = 80;
			ultraGridColumn114.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance86.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn114.CellAppearance = appearance86;
			ultraGridColumn114.Format = "#,##0.00";
			ultraGridColumn114.Header.VisiblePosition = 79;
			ultraGridColumn115.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance87.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn115.CellAppearance = appearance87;
			ultraGridColumn115.Format = "#,##0.00";
			ultraGridColumn115.Header.VisiblePosition = 80;
			ultraGridColumn115.Width = 80;
			ultraGridColumn116.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance88.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn116.CellAppearance = appearance88;
			ultraGridColumn116.Format = "#,##0.00";
			ultraGridColumn116.Header.VisiblePosition = 81;
			ultraGridColumn116.Width = 80;
			ultraGridColumn117.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance89.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn117.CellAppearance = appearance89;
			ultraGridColumn117.Format = "#,##0.00";
			ultraGridColumn117.Header.VisiblePosition = 82;
			ultraGridColumn118.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance90.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn118.CellAppearance = appearance90;
			ultraGridColumn118.Header.VisiblePosition = 83;
			ultraGridColumn118.Width = 80;
			ultraGridColumn119.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance91.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn119.CellAppearance = appearance91;
			ultraGridColumn119.Format = "#,##0.00";
			ultraGridColumn119.Header.VisiblePosition = 84;
			ultraGridColumn119.Width = 80;
			ultraGridColumn120.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance92.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn120.CellAppearance = appearance92;
			ultraGridColumn120.Format = "#,##0.00";
			ultraGridColumn120.Header.VisiblePosition = 85;
			ultraGridColumn120.Width = 80;
			ultraGridColumn121.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance93.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn121.CellAppearance = appearance93;
			ultraGridColumn121.Format = "#,##0.00";
			ultraGridColumn121.Header.VisiblePosition = 86;
			ultraGridColumn122.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn122.CellAppearance = appearance94;
			ultraGridColumn122.Format = "#,##0.00";
			ultraGridColumn122.Header.VisiblePosition = 87;
			ultraGridColumn122.Width = 80;
			ultraGridColumn123.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn123.CellAppearance = appearance95;
			ultraGridColumn123.Format = "#,##0.00";
			ultraGridColumn123.Header.VisiblePosition = 88;
			ultraGridColumn123.Width = 80;
			ultraGridColumn124.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance96.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn124.CellAppearance = appearance96;
			ultraGridColumn124.Format = "#,##0.00";
			ultraGridColumn124.Header.VisiblePosition = 89;
			ultraGridColumn125.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance97.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn125.CellAppearance = appearance97;
			ultraGridColumn125.Header.VisiblePosition = 90;
			ultraGridColumn125.Width = 80;
			ultraGridColumn126.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance98.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn126.CellAppearance = appearance98;
			ultraGridColumn126.Format = "#,##0.00";
			ultraGridColumn126.Header.VisiblePosition = 91;
			ultraGridColumn126.Width = 80;
			ultraGridColumn127.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance99.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn127.CellAppearance = appearance99;
			ultraGridColumn127.Format = "#,##0.00";
			ultraGridColumn127.Header.VisiblePosition = 92;
			ultraGridColumn127.Width = 80;
			ultraGridColumn128.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance100.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn128.CellAppearance = appearance100;
			ultraGridColumn128.Format = "#,##0.00";
			ultraGridColumn128.Header.VisiblePosition = 93;
			ultraGridColumn129.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance101.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn129.CellAppearance = appearance101;
			ultraGridColumn129.Format = "#,##0.00";
			ultraGridColumn129.Header.VisiblePosition = 94;
			ultraGridColumn129.Width = 80;
			ultraGridColumn130.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance102.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn130.CellAppearance = appearance102;
			ultraGridColumn130.Format = "#,##0.00";
			ultraGridColumn130.Header.VisiblePosition = 95;
			ultraGridColumn130.Width = 80;
			ultraGridColumn131.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance103.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn131.CellAppearance = appearance103;
			ultraGridColumn131.Format = "#,##0.00";
			ultraGridColumn131.Header.VisiblePosition = 96;
			ultraGridColumn131.Width = 80;
			ultraGridColumn132.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance104.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn132.CellAppearance = appearance104;
			ultraGridColumn132.Header.VisiblePosition = 97;
			ultraGridColumn132.Width = 80;
			ultraGridColumn133.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance105.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn133.CellAppearance = appearance105;
			ultraGridColumn133.Format = "#,##0.00";
			ultraGridColumn133.Header.VisiblePosition = 98;
			ultraGridColumn133.Width = 80;
			ultraGridColumn134.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance106.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn134.CellAppearance = appearance106;
			ultraGridColumn134.Format = "#,##0.00";
			ultraGridColumn134.Header.VisiblePosition = 99;
			ultraGridColumn134.Width = 80;
			ultraGridColumn135.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance107.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn135.CellAppearance = appearance107;
			ultraGridColumn135.Format = "#,##0.00";
			ultraGridColumn135.Header.VisiblePosition = 100;
			ultraGridColumn136.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance108.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn136.CellAppearance = appearance108;
			ultraGridColumn136.Format = "#,##0.00";
			ultraGridColumn136.Header.VisiblePosition = 101;
			ultraGridColumn136.Width = 80;
			ultraGridColumn137.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance109.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn137.CellAppearance = appearance109;
			ultraGridColumn137.Format = "#,##0.00";
			ultraGridColumn137.Header.VisiblePosition = 102;
			ultraGridColumn137.Width = 80;
			ultraGridColumn138.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn138.CellAppearance = appearance110;
			ultraGridColumn138.Format = "#,##0.00";
			ultraGridColumn138.Header.VisiblePosition = 103;
			ultraGridColumn138.Width = 80;
			ultraGridColumn139.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance111.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn139.CellAppearance = appearance111;
			ultraGridColumn139.Header.VisiblePosition = 104;
			ultraGridColumn139.Width = 80;
			ultraGridColumn140.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance112.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn140.CellAppearance = appearance112;
			ultraGridColumn140.Format = "#,##0.00";
			ultraGridColumn140.Header.VisiblePosition = 105;
			ultraGridColumn140.Width = 80;
			ultraGridColumn141.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance113.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn141.CellAppearance = appearance113;
			ultraGridColumn141.Format = "#,##0.00";
			ultraGridColumn141.Header.VisiblePosition = 106;
			ultraGridColumn141.Width = 80;
			ultraGridColumn142.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance114.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn142.CellAppearance = appearance114;
			ultraGridColumn142.Format = "#,##0.00";
			ultraGridColumn142.Header.VisiblePosition = 107;
			ultraGridColumn143.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance115.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn143.CellAppearance = appearance115;
			ultraGridColumn143.Format = "#,##0.00";
			ultraGridColumn143.Header.VisiblePosition = 108;
			ultraGridColumn143.Width = 80;
			ultraGridColumn144.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance116.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn144.CellAppearance = appearance116;
			ultraGridColumn144.Format = "#,##0.00";
			ultraGridColumn144.Header.VisiblePosition = 109;
			ultraGridColumn144.Width = 80;
			ultraGridColumn145.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance117.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn145.CellAppearance = appearance117;
			ultraGridColumn145.Format = "#,##0.00";
			ultraGridColumn145.Header.VisiblePosition = 110;
			ultraGridColumn145.Width = 80;
			ultraGridColumn146.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance118.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn146.CellAppearance = appearance118;
			ultraGridColumn146.Header.VisiblePosition = 111;
			ultraGridColumn146.Width = 80;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn99,
																										 ultraGridColumn100,
																										 ultraGridColumn101,
																										 ultraGridColumn102,
																										 ultraGridColumn103,
																										 ultraGridColumn104,
																										 ultraGridColumn105,
																										 ultraGridColumn106,
																										 ultraGridColumn107,
																										 ultraGridColumn108,
																										 ultraGridColumn109,
																										 ultraGridColumn110,
																										 ultraGridColumn111,
																										 ultraGridColumn112,
																										 ultraGridColumn113,
																										 ultraGridColumn114,
																										 ultraGridColumn115,
																										 ultraGridColumn116,
																										 ultraGridColumn117,
																										 ultraGridColumn118,
																										 ultraGridColumn119,
																										 ultraGridColumn120,
																										 ultraGridColumn121,
																										 ultraGridColumn122,
																										 ultraGridColumn123,
																										 ultraGridColumn124,
																										 ultraGridColumn125,
																										 ultraGridColumn126,
																										 ultraGridColumn127,
																										 ultraGridColumn128,
																										 ultraGridColumn129,
																										 ultraGridColumn130,
																										 ultraGridColumn131,
																										 ultraGridColumn132,
																										 ultraGridColumn133,
																										 ultraGridColumn134,
																										 ultraGridColumn135,
																										 ultraGridColumn136,
																										 ultraGridColumn137,
																										 ultraGridColumn138,
																										 ultraGridColumn139,
																										 ultraGridColumn140,
																										 ultraGridColumn141,
																										 ultraGridColumn142,
																										 ultraGridColumn143,
																										 ultraGridColumn144,
																										 ultraGridColumn145,
																										 ultraGridColumn146});
			this.uGridVariableOperacion.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridVariableOperacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVariableOperacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVariableOperacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance119.BackColor = System.Drawing.Color.Transparent;
			this.uGridVariableOperacion.DisplayLayout.Override.CardAreaAppearance = appearance119;
			appearance120.ForeColor = System.Drawing.Color.Black;
			appearance120.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVariableOperacion.DisplayLayout.Override.CellAppearance = appearance120;
			appearance121.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance121.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance121.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance121.FontData.BoldAsString = "True";
			appearance121.FontData.Name = "Arial";
			appearance121.FontData.SizeInPoints = 8F;
			appearance121.ForeColor = System.Drawing.Color.White;
			appearance121.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVariableOperacion.DisplayLayout.Override.HeaderAppearance = appearance121;
			this.uGridVariableOperacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance122.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance122.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance122.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVariableOperacion.DisplayLayout.Override.RowSelectorAppearance = appearance122;
			appearance123.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance123.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance123.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVariableOperacion.DisplayLayout.Override.SelectedRowAppearance = appearance123;
			this.uGridVariableOperacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVariableOperacion.Location = new System.Drawing.Point(8, 96);
			this.uGridVariableOperacion.Name = "uGridVariableOperacion";
			this.uGridVariableOperacion.Size = new System.Drawing.Size(1288, 376);
			this.uGridVariableOperacion.TabIndex = 1066;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn36.DataType = typeof(System.DateTime);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn50.DataType = typeof(System.Decimal);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn53.DataType = typeof(int);
			ultraDataColumn54.DataType = typeof(int);
			ultraDataColumn55.DataType = typeof(int);
			ultraDataColumn56.DataType = typeof(System.Decimal);
			ultraDataColumn57.DataType = typeof(System.Decimal);
			ultraDataColumn58.DataType = typeof(System.Decimal);
			ultraDataColumn59.DataType = typeof(System.Decimal);
			ultraDataColumn60.DataType = typeof(System.Decimal);
			ultraDataColumn61.DataType = typeof(System.Decimal);
			ultraDataColumn62.DataType = typeof(int);
			ultraDataColumn63.DataType = typeof(System.Decimal);
			ultraDataColumn64.DataType = typeof(System.Decimal);
			ultraDataColumn65.DataType = typeof(System.Decimal);
			ultraDataColumn66.DataType = typeof(System.Decimal);
			ultraDataColumn67.DataType = typeof(System.Decimal);
			ultraDataColumn68.DataType = typeof(System.Decimal);
			ultraDataColumn69.DataType = typeof(int);
			ultraDataColumn70.DataType = typeof(System.Decimal);
			ultraDataColumn71.DataType = typeof(System.Decimal);
			ultraDataColumn72.DataType = typeof(System.Decimal);
			ultraDataColumn73.DataType = typeof(System.Decimal);
			ultraDataColumn74.DataType = typeof(System.Decimal);
			ultraDataColumn75.DataType = typeof(System.Decimal);
			ultraDataColumn76.DataType = typeof(int);
			ultraDataColumn77.DataType = typeof(System.Decimal);
			ultraDataColumn78.DataType = typeof(System.Decimal);
			ultraDataColumn79.DataType = typeof(System.Decimal);
			ultraDataColumn80.DataType = typeof(System.Decimal);
			ultraDataColumn81.DataType = typeof(System.Decimal);
			ultraDataColumn82.DataType = typeof(System.Decimal);
			ultraDataColumn83.DataType = typeof(int);
			ultraDataColumn84.DataType = typeof(System.Decimal);
			ultraDataColumn85.DataType = typeof(System.Decimal);
			ultraDataColumn86.DataType = typeof(System.Decimal);
			ultraDataColumn87.DataType = typeof(System.Decimal);
			ultraDataColumn88.DataType = typeof(System.Decimal);
			ultraDataColumn89.DataType = typeof(System.Decimal);
			ultraDataColumn90.DataType = typeof(int);
			ultraDataColumn91.DataType = typeof(System.Decimal);
			ultraDataColumn92.DataType = typeof(System.Decimal);
			ultraDataColumn93.DataType = typeof(System.Decimal);
			ultraDataColumn94.DataType = typeof(System.Decimal);
			ultraDataColumn95.DataType = typeof(System.Decimal);
			ultraDataColumn96.DataType = typeof(System.Decimal);
			ultraDataColumn97.DataType = typeof(int);
			ultraDataColumn98.DataType = typeof(System.Decimal);
			ultraDataColumn99.DataType = typeof(System.Decimal);
			ultraDataColumn100.DataType = typeof(System.Decimal);
			ultraDataColumn101.DataType = typeof(System.Decimal);
			ultraDataColumn102.DataType = typeof(System.Decimal);
			ultraDataColumn103.DataType = typeof(System.Decimal);
			ultraDataColumn104.DataType = typeof(int);
			ultraDataColumn105.DataType = typeof(System.Decimal);
			ultraDataColumn106.DataType = typeof(System.Decimal);
			ultraDataColumn107.DataType = typeof(System.Decimal);
			ultraDataColumn108.DataType = typeof(System.Decimal);
			ultraDataColumn109.DataType = typeof(System.Decimal);
			ultraDataColumn110.DataType = typeof(System.Decimal);
			ultraDataColumn111.DataType = typeof(int);
			ultraDataColumn112.DataType = typeof(System.Decimal);
			ultraDataColumn113.DataType = typeof(System.Decimal);
			ultraDataColumn114.DataType = typeof(System.Decimal);
			ultraDataColumn115.DataType = typeof(System.Decimal);
			ultraDataColumn116.DataType = typeof(System.Decimal);
			ultraDataColumn117.DataType = typeof(System.Decimal);
			ultraDataColumn118.DataType = typeof(int);
			ultraDataColumn119.DataType = typeof(System.Decimal);
			ultraDataColumn120.DataType = typeof(System.Decimal);
			ultraDataColumn121.DataType = typeof(System.Decimal);
			ultraDataColumn122.DataType = typeof(System.Decimal);
			ultraDataColumn123.DataType = typeof(System.Decimal);
			ultraDataColumn124.DataType = typeof(System.Decimal);
			ultraDataColumn125.DataType = typeof(int);
			ultraDataColumn126.DataType = typeof(System.Decimal);
			ultraDataColumn127.DataType = typeof(System.Decimal);
			ultraDataColumn128.DataType = typeof(System.Decimal);
			ultraDataColumn129.DataType = typeof(System.Decimal);
			ultraDataColumn130.DataType = typeof(System.Decimal);
			ultraDataColumn131.DataType = typeof(System.Decimal);
			ultraDataColumn132.DataType = typeof(int);
			ultraDataColumn133.DataType = typeof(System.Decimal);
			ultraDataColumn134.DataType = typeof(System.Decimal);
			ultraDataColumn135.DataType = typeof(System.Decimal);
			ultraDataColumn136.DataType = typeof(System.Decimal);
			ultraDataColumn137.DataType = typeof(System.Decimal);
			ultraDataColumn138.DataType = typeof(System.Decimal);
			ultraDataColumn139.DataType = typeof(int);
			ultraDataColumn140.DataType = typeof(System.Decimal);
			ultraDataColumn141.DataType = typeof(System.Decimal);
			ultraDataColumn142.DataType = typeof(System.Decimal);
			ultraDataColumn143.DataType = typeof(System.Decimal);
			ultraDataColumn144.DataType = typeof(System.Decimal);
			ultraDataColumn145.DataType = typeof(System.Decimal);
			ultraDataColumn146.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn94,
																																 ultraDataColumn95,
																																 ultraDataColumn96,
																																 ultraDataColumn97,
																																 ultraDataColumn98,
																																 ultraDataColumn99,
																																 ultraDataColumn100,
																																 ultraDataColumn101,
																																 ultraDataColumn102,
																																 ultraDataColumn103,
																																 ultraDataColumn104,
																																 ultraDataColumn105,
																																 ultraDataColumn106,
																																 ultraDataColumn107,
																																 ultraDataColumn108,
																																 ultraDataColumn109,
																																 ultraDataColumn110,
																																 ultraDataColumn111,
																																 ultraDataColumn112,
																																 ultraDataColumn113,
																																 ultraDataColumn114,
																																 ultraDataColumn115,
																																 ultraDataColumn116,
																																 ultraDataColumn117,
																																 ultraDataColumn118,
																																 ultraDataColumn119,
																																 ultraDataColumn120,
																																 ultraDataColumn121,
																																 ultraDataColumn122,
																																 ultraDataColumn123,
																																 ultraDataColumn124,
																																 ultraDataColumn125,
																																 ultraDataColumn126,
																																 ultraDataColumn127,
																																 ultraDataColumn128,
																																 ultraDataColumn129,
																																 ultraDataColumn130,
																																 ultraDataColumn131,
																																 ultraDataColumn132,
																																 ultraDataColumn133,
																																 ultraDataColumn134,
																																 ultraDataColumn135,
																																 ultraDataColumn136,
																																 ultraDataColumn137,
																																 ultraDataColumn138,
																																 ultraDataColumn139,
																																 ultraDataColumn140,
																																 ultraDataColumn141,
																																 ultraDataColumn142,
																																 ultraDataColumn143,
																																 ultraDataColumn144,
																																 ultraDataColumn145,
																																 ultraDataColumn146});
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(496, 48);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 1067;
			// 
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Enabled = false;
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(1216, 8);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.TabIndex = 1068;
			this.btnComandos.Text = "Exportar";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click);
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesCopiarUltimoMes,
																																								this.mnuAccionesAgregarNuevoLocal,
																																								this.menuItem1});
			// 
			// mnuAccionesCopiarUltimoMes
			// 
			this.mnuAccionesCopiarUltimoMes.Index = 0;
			this.mnuAccionesCopiarUltimoMes.Text = "Consulta EquiFax";
			this.mnuAccionesCopiarUltimoMes.Click += new System.EventHandler(this.mnuAccionesCopiarUltimoMes_Click);
			// 
			// mnuAccionesAgregarNuevoLocal
			// 
			this.mnuAccionesAgregarNuevoLocal.Index = 1;
			this.mnuAccionesAgregarNuevoLocal.Text = "Variables Nivel Identificación ";
			this.mnuAccionesAgregarNuevoLocal.Click += new System.EventHandler(this.mnuAccionesAgregarNuevoLocal_Click);
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 2;
			this.menuItem1.Text = "Variable pto Operación";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
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
			// frmReporteAnalisiEquiFax
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1304, 478);
			this.Controls.Add(this.btnComandos);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.uGridVariableOperacion);
			this.Controls.Add(this.uGridVariableNinelIdent);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.uGridConsulta);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Name = "frmReporteAnalisiEquiFax";
			this.Text = "Originación POINT ( EquiFax)";
			this.Load += new System.EventHandler(this.frmReporteAnalisiEquiFax_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVariableNinelIdent)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVariableOperacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}


		private void frmReporteAnalisiEquiFax_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200608");

			if (!Funcion.Permiso("1322", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar al Reporte ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar) this.btnComandos.Enabled =  true;

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2016");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2016");

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Today.AddDays(-30);
			this.dtHasta.Value = DateTime.Today;

			this.Estado ();

		}
		public void Estado ()
		{

			#region Estado
			if ( (int) this.optEstados.Value == 0)
			{
				this.uGridConsulta.Visible =  true;
				this.uGridVariableNinelIdent.Visible =  false;
				this.uGridVariableOperacion.Visible =  false;

			}
			if ( (int) this.optEstados.Value == 1)
			{
				this.uGridConsulta.Visible =  false;
				this.uGridVariableNinelIdent.Visible =  true;
				this.uGridVariableOperacion.Visible =  false;

			}
			if ( (int) this.optEstados.Value == 2)
			{
				this.uGridConsulta.Visible =  true;
				this.uGridVariableNinelIdent.Visible =  false;
				this.uGridVariableOperacion.Visible =  true;

			}
     #endregion Estado

		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
      this.Estado ();

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

			this.lblContador.Text ="";

			if ( (int) this.optEstados.Value == 0)
			{
				string sSQLr = string.Format("Exec [ReporteConsultaEquifaxCedula] '{0}', '{1}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
				FuncionesProcedimientos.dsGeneralModulo(sSQLr, this.uGridConsulta);

						this.lblContador.Text = this.uGridConsulta.Rows.Count + " CONSULTAS ENCONTRADAS";
			}
			if ( (int) this.optEstados.Value == 1)
			{
				string sSQLi = string.Format("Exec [ReporteIdentificacíonEquiFaxOtorgadas] '{0}', '{1}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
				FuncionesProcedimientos.dsGeneralModulo(sSQLi, this.uGridVariableNinelIdent);
							this.lblContador.Text = this.uGridVariableNinelIdent.Rows.Count + " CONSULTAS ENCONTRADAS";
			}
			if ( (int) this.optEstados.Value == 2)
			{
				string sSQLe = string.Format("Exec [ReporteVariableOperacionEquiFax] '{0}', '{1}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
				FuncionesProcedimientos.dsGeneralModulo(sSQLe, this.uGridVariableOperacion);

							this.lblContador.Text = this.uGridVariableOperacion.Rows.Count + " CONSULTAS ENCONTRADAS";
			}
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			this.Estado ();
			this.btnVer_Click (sender, e);
		}

		private void mnuAccionesCopiarUltimoMes_Click(object sender, System.EventArgs e)
		{
					FuncionesProcedimientos.ExportaExcel(this.uGridConsulta);
		}

		private void mnuAccionesAgregarNuevoLocal_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.uGridVariableNinelIdent);
		}

		private void btnComandos_Click(object sender, System.EventArgs e)
		{
				this.mnuAcciones.Show(this.btnComandos, new Point(1, 25));
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
		 FuncionesProcedimientos.ExportaExcel(this.uGridVariableOperacion);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if ( (int) this.optEstados.Value == 2)
			{
				if (this.dtDesde.Value != System.DBNull.Value)
				{
					this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/07/2017");
					this.dtHasta.CalendarInfo.MaxDate = DateTime.Parse("31/12/2030");

					if (Convert.ToDateTime(this.dtDesde.Value).Month == DateTime.Today.Month) 
					{
						this.dtHasta.Value = DateTime.Today;
						this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;
					}
					else
					{
						this.dtHasta.Value = DateTime.DaysInMonth(Convert.ToDateTime(this.dtDesde.Value).Year, Convert.ToDateTime(this.dtDesde.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Year.ToString();
						this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
					}

					this.dtHasta.CalendarInfo.MinDate = Convert.ToDateTime(this.dtDesde.Value);						
				}
				else this.dtHasta.Value = System.DBNull.Value;
			}
			
		}
	}
}
