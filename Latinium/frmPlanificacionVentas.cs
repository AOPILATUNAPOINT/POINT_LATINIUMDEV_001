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
	/// Descripción breve de frmPlanificacionVentas.
	/// </summary>
	public class frmPlanificacionVentas : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnAprobarSemana6;
		private System.Windows.Forms.Button btnAprobarSemana5;
		private System.Windows.Forms.Button btnAprobarSemana4;
		private System.Windows.Forms.Button btnAprobarSemana3;
		private System.Windows.Forms.Button btnAprobarSemana2;
		private System.Windows.Forms.Button btnAprobarSemana1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.DateTimePicker dtpMes;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana4;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana3;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana2;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbLocales;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridSemana6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombres;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdPlanificacionVentas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPlanificacionVentas()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPlanificacionVentas));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Ubicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ubicacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnHorarios", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnLunes", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMartes", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMiercoles", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJueves", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnViernes", 5);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSabado", 6);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDomingo", 7);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lunes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Martes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Miercoles");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jueves");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Viernes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sabado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domingo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnHorarios", 0);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnLunes", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMartes", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMiercoles", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJueves", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnViernes", 5);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSabado", 6);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDomingo", 7);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnHorarios", 0);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnLunes", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMartes", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMiercoles", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJueves", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnViernes", 5);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSabado", 6);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDomingo", 7);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnHorarios", 0);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnLunes", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMartes", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMiercoles", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJueves", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnViernes", 5);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSabado", 6);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDomingo", 7);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnHorarios", 0);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnLunes", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMartes", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMiercoles", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJueves", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnViernes", 5);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSabado", 6);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDomingo", 7);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetallePlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlanificacionVentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lunes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Martes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Miercoles", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jueves", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Viernes", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sabado", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domingo", -1, "cmbLocales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnHorarios", 0);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnLunes", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMartes", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnMiercoles", 3);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnJueves", 4);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnViernes", 5);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSabado", 6);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDomingo", 7);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnAprobarSemana6 = new System.Windows.Forms.Button();
			this.btnAprobarSemana5 = new System.Windows.Forms.Button();
			this.btnAprobarSemana4 = new System.Windows.Forms.Button();
			this.btnAprobarSemana3 = new System.Windows.Forms.Button();
			this.btnAprobarSemana2 = new System.Windows.Forms.Button();
			this.btnAprobarSemana1 = new System.Windows.Forms.Button();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdPlanificacionVentas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtpMes = new System.Windows.Forms.DateTimePicker();
			this.lblBodega = new System.Windows.Forms.Label();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSemana5 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSemana4 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridSemana3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridSemana2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridSemana1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbNombres = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbLocales = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.uGridSemana6 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPlanificacionVentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana6)).BeginInit();
			this.SuspendLayout();
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
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
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
																																 ultraDataColumn10});
			// 
			// btnAprobarSemana6
			// 
			this.btnAprobarSemana6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnAprobarSemana6.CausesValidation = false;
			this.btnAprobarSemana6.Enabled = false;
			this.btnAprobarSemana6.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnAprobarSemana6.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnAprobarSemana6.Image = ((System.Drawing.Image)(resources.GetObject("btnAprobarSemana6.Image")));
			this.btnAprobarSemana6.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAprobarSemana6.Location = new System.Drawing.Point(0, 552);
			this.btnAprobarSemana6.Name = "btnAprobarSemana6";
			this.btnAprobarSemana6.TabIndex = 790;
			this.btnAprobarSemana6.Text = "Aprobar";
			this.btnAprobarSemana6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana6.Click += new System.EventHandler(this.btnAprobarSemana6_Click);
			// 
			// btnAprobarSemana5
			// 
			this.btnAprobarSemana5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnAprobarSemana5.CausesValidation = false;
			this.btnAprobarSemana5.Enabled = false;
			this.btnAprobarSemana5.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnAprobarSemana5.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnAprobarSemana5.Image = ((System.Drawing.Image)(resources.GetObject("btnAprobarSemana5.Image")));
			this.btnAprobarSemana5.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAprobarSemana5.Location = new System.Drawing.Point(0, 448);
			this.btnAprobarSemana5.Name = "btnAprobarSemana5";
			this.btnAprobarSemana5.TabIndex = 789;
			this.btnAprobarSemana5.Text = "Aprobar";
			this.btnAprobarSemana5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana5.Click += new System.EventHandler(this.btnAprobarSemana5_Click);
			// 
			// btnAprobarSemana4
			// 
			this.btnAprobarSemana4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnAprobarSemana4.CausesValidation = false;
			this.btnAprobarSemana4.Enabled = false;
			this.btnAprobarSemana4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnAprobarSemana4.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnAprobarSemana4.Image = ((System.Drawing.Image)(resources.GetObject("btnAprobarSemana4.Image")));
			this.btnAprobarSemana4.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAprobarSemana4.Location = new System.Drawing.Point(0, 344);
			this.btnAprobarSemana4.Name = "btnAprobarSemana4";
			this.btnAprobarSemana4.TabIndex = 788;
			this.btnAprobarSemana4.Text = "Aprobar";
			this.btnAprobarSemana4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana4.Click += new System.EventHandler(this.btnAprobarSemana4_Click);
			// 
			// btnAprobarSemana3
			// 
			this.btnAprobarSemana3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnAprobarSemana3.CausesValidation = false;
			this.btnAprobarSemana3.Enabled = false;
			this.btnAprobarSemana3.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnAprobarSemana3.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnAprobarSemana3.Image = ((System.Drawing.Image)(resources.GetObject("btnAprobarSemana3.Image")));
			this.btnAprobarSemana3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAprobarSemana3.Location = new System.Drawing.Point(0, 240);
			this.btnAprobarSemana3.Name = "btnAprobarSemana3";
			this.btnAprobarSemana3.TabIndex = 787;
			this.btnAprobarSemana3.Text = "Aprobar";
			this.btnAprobarSemana3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana3.Click += new System.EventHandler(this.btnAprobarSemana3_Click);
			// 
			// btnAprobarSemana2
			// 
			this.btnAprobarSemana2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnAprobarSemana2.CausesValidation = false;
			this.btnAprobarSemana2.Enabled = false;
			this.btnAprobarSemana2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnAprobarSemana2.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnAprobarSemana2.Image = ((System.Drawing.Image)(resources.GetObject("btnAprobarSemana2.Image")));
			this.btnAprobarSemana2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAprobarSemana2.Location = new System.Drawing.Point(0, 136);
			this.btnAprobarSemana2.Name = "btnAprobarSemana2";
			this.btnAprobarSemana2.TabIndex = 786;
			this.btnAprobarSemana2.Text = "Aprobar";
			this.btnAprobarSemana2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana2.Click += new System.EventHandler(this.btnAprobarSemana2_Click);
			// 
			// btnAprobarSemana1
			// 
			this.btnAprobarSemana1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnAprobarSemana1.CausesValidation = false;
			this.btnAprobarSemana1.Enabled = false;
			this.btnAprobarSemana1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnAprobarSemana1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnAprobarSemana1.Image = ((System.Drawing.Image)(resources.GetObject("btnAprobarSemana1.Image")));
			this.btnAprobarSemana1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAprobarSemana1.Location = new System.Drawing.Point(0, 32);
			this.btnAprobarSemana1.Name = "btnAprobarSemana1";
			this.btnAprobarSemana1.TabIndex = 785;
			this.btnAprobarSemana1.Text = "Aprobar";
			this.btnAprobarSemana1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobarSemana1.Click += new System.EventHandler(this.btnAprobarSemana1_Click);
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(432, 426);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 783;
			this.txtEstado.Visible = false;
			// 
			// txtIdPlanificacionVentas
			// 
			this.txtIdPlanificacionVentas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdPlanificacionVentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdPlanificacionVentas.Enabled = false;
			this.txtIdPlanificacionVentas.Location = new System.Drawing.Point(408, 426);
			this.txtIdPlanificacionVentas.Name = "txtIdPlanificacionVentas";
			this.txtIdPlanificacionVentas.PromptChar = ' ';
			this.txtIdPlanificacionVentas.Size = new System.Drawing.Size(16, 21);
			this.txtIdPlanificacionVentas.TabIndex = 781;
			this.txtIdPlanificacionVentas.Visible = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(504, 7);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 775;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(579, 7);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 776;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(660, 7);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 778;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(739, 7);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 777;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(818, 7);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 779;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(893, 7);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 780;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// dtpMes
			// 
			this.dtpMes.Checked = false;
			this.dtpMes.CustomFormat = "MM/yyyy";
			this.dtpMes.Enabled = false;
			this.dtpMes.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpMes.Location = new System.Drawing.Point(312, 8);
			this.dtpMes.Name = "dtpMes";
			this.dtpMes.ShowUpDown = true;
			this.dtpMes.Size = new System.Drawing.Size(88, 20);
			this.dtpMes.TabIndex = 768;
			this.dtpMes.Value = new System.DateTime(2017, 1, 1, 0, 0, 0, 0);
			this.dtpMes.ValueChanged += new System.EventHandler(this.dtpMes_ValueChanged);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(44, 16);
			this.lblBodega.TabIndex = 762;
			this.lblBodega.Text = "Nombre";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// uGridSemana5
			// 
			this.uGridSemana5.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana5.DataSource = this.ultraDataSource7;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana5.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 25;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 89;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 46;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn4.Width = 120;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn5.Width = 120;
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn6.Width = 120;
			ultraGridColumn7.Header.VisiblePosition = 10;
			ultraGridColumn7.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn7.Width = 120;
			ultraGridColumn8.Header.VisiblePosition = 12;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn8.Width = 120;
			ultraGridColumn9.Header.VisiblePosition = 14;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.Header.VisiblePosition = 16;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn10.Width = 120;
			ultraGridColumn11.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			appearance2.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance2.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn11.CellButtonAppearance = appearance2;
			ultraGridColumn11.Header.Caption = "...";
			ultraGridColumn11.Header.VisiblePosition = 3;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn11.Width = 25;
			ultraGridColumn12.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn12.Header.Caption = "...";
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn12.Width = 16;
			ultraGridColumn13.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn13.Header.Caption = "...";
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn13.Width = 16;
			ultraGridColumn14.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn14.Header.Caption = "...";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn14.Width = 16;
			ultraGridColumn15.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn15.Header.Caption = "...";
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn15.Width = 16;
			ultraGridColumn16.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn16.Header.Caption = "...";
			ultraGridColumn16.Header.VisiblePosition = 13;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn16.Width = 16;
			ultraGridColumn17.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn17.Header.Caption = "...";
			ultraGridColumn17.Header.VisiblePosition = 15;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn17.Width = 16;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn18.Header.Caption = "...";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 16;
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
																										 ultraGridColumn18});
			ultraGridBand1.Header.Caption = "Semana 5";
			ultraGridBand1.HeaderVisible = true;
			this.uGridSemana5.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana5.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.uGridSemana5.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana5.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana5.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana5.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 7.5F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana5.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.uGridSemana5.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana5.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana5.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana5.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridSemana5.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana5.Location = new System.Drawing.Point(0, 448);
			this.uGridSemana5.Name = "uGridSemana5";
			this.uGridSemana5.Size = new System.Drawing.Size(1018, 104);
			this.uGridSemana5.TabIndex = 767;
			this.uGridSemana5.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana5_AfterRowInsert);
			this.uGridSemana5.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana5_KeyDown);
			this.uGridSemana5.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana5_ClickCellButton);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42});
			// 
			// ultraDataSource6
			// 
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(int);
			ultraDataColumn50.DataType = typeof(int);
			ultraDataColumn51.DataType = typeof(int);
			ultraDataColumn52.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52});
			// 
			// ultraDataSource4
			// 
			ultraDataColumn53.DataType = typeof(int);
			ultraDataColumn54.DataType = typeof(int);
			ultraDataColumn55.DataType = typeof(int);
			ultraDataColumn56.DataType = typeof(int);
			ultraDataColumn57.DataType = typeof(int);
			ultraDataColumn58.DataType = typeof(int);
			ultraDataColumn59.DataType = typeof(int);
			ultraDataColumn60.DataType = typeof(int);
			ultraDataColumn61.DataType = typeof(int);
			ultraDataColumn62.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn53,
																																 ultraDataColumn54,
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58,
																																 ultraDataColumn59,
																																 ultraDataColumn60,
																																 ultraDataColumn61,
																																 ultraDataColumn62});
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(272, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(26, 16);
			this.label4.TabIndex = 760;
			this.label4.Text = "Mes";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn63.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn63,
																																 ultraDataColumn64,
																																 ultraDataColumn65});
			// 
			// uGridSemana4
			// 
			this.uGridSemana4.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana4.DataSource = this.ultraDataSource6;
			appearance9.BackColor = System.Drawing.Color.White;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana4.DisplayLayout.Appearance = appearance9;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 116;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 89;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 46;
			ultraGridColumn22.Header.VisiblePosition = 4;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn22.Width = 120;
			ultraGridColumn23.Header.VisiblePosition = 6;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn23.Width = 120;
			ultraGridColumn24.Header.VisiblePosition = 8;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn24.Width = 120;
			ultraGridColumn25.Header.VisiblePosition = 10;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn25.Width = 120;
			ultraGridColumn26.Header.VisiblePosition = 12;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn26.Width = 120;
			ultraGridColumn27.Header.VisiblePosition = 14;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn27.Width = 120;
			ultraGridColumn28.Header.VisiblePosition = 16;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn28.Width = 120;
			ultraGridColumn29.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			appearance10.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance10.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn29.CellButtonAppearance = appearance10;
			ultraGridColumn29.Header.Caption = "...";
			ultraGridColumn29.Header.VisiblePosition = 3;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn29.Width = 25;
			ultraGridColumn30.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn30.Header.Caption = "...";
			ultraGridColumn30.Header.VisiblePosition = 5;
			ultraGridColumn30.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn30.Width = 16;
			ultraGridColumn31.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn31.Header.Caption = "...";
			ultraGridColumn31.Header.VisiblePosition = 7;
			ultraGridColumn31.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn31.Width = 16;
			ultraGridColumn32.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn32.Header.Caption = "...";
			ultraGridColumn32.Header.VisiblePosition = 9;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn32.Width = 16;
			ultraGridColumn33.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn33.Header.Caption = "...";
			ultraGridColumn33.Header.VisiblePosition = 11;
			ultraGridColumn33.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn33.Width = 16;
			ultraGridColumn34.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn34.Header.Caption = "...";
			ultraGridColumn34.Header.VisiblePosition = 13;
			ultraGridColumn34.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn34.Width = 16;
			ultraGridColumn35.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn35.Header.Caption = "...";
			ultraGridColumn35.Header.VisiblePosition = 15;
			ultraGridColumn35.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn35.Width = 16;
			ultraGridColumn36.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn36.Header.Caption = "...";
			ultraGridColumn36.Header.VisiblePosition = 17;
			ultraGridColumn36.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn36.Width = 16;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn36});
			ultraGridBand2.Header.Caption = "Semana 4";
			ultraGridBand2.HeaderVisible = true;
			this.uGridSemana4.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana4.DisplayLayout.Override.ActiveRowAppearance = appearance11;
			this.uGridSemana4.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana4.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana4.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana4.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 7.5F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana4.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.uGridSemana4.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana4.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana4.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana4.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.uGridSemana4.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana4.Location = new System.Drawing.Point(0, 344);
			this.uGridSemana4.Name = "uGridSemana4";
			this.uGridSemana4.Size = new System.Drawing.Size(1018, 104);
			this.uGridSemana4.TabIndex = 766;
			this.uGridSemana4.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana4_AfterRowInsert);
			this.uGridSemana4.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana4_KeyDown);
			this.uGridSemana4.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana4_ClickCellButton);
			// 
			// uGridSemana3
			// 
			this.uGridSemana3.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana3.DataSource = this.ultraDataSource5;
			appearance17.BackColor = System.Drawing.Color.White;
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana3.DisplayLayout.Appearance = appearance17;
			ultraGridColumn37.Header.VisiblePosition = 0;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 116;
			ultraGridColumn38.Header.VisiblePosition = 1;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 89;
			ultraGridColumn39.Header.VisiblePosition = 2;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 46;
			ultraGridColumn40.Header.VisiblePosition = 4;
			ultraGridColumn40.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn40.Width = 120;
			ultraGridColumn41.Header.VisiblePosition = 6;
			ultraGridColumn41.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn41.Width = 120;
			ultraGridColumn42.Header.VisiblePosition = 8;
			ultraGridColumn42.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn42.Width = 120;
			ultraGridColumn43.Header.VisiblePosition = 10;
			ultraGridColumn43.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn43.Width = 120;
			ultraGridColumn44.Header.VisiblePosition = 12;
			ultraGridColumn44.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn44.Width = 120;
			ultraGridColumn45.Header.VisiblePosition = 14;
			ultraGridColumn45.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn45.Width = 120;
			ultraGridColumn46.Header.VisiblePosition = 16;
			ultraGridColumn46.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn46.Width = 120;
			ultraGridColumn47.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance18.Image = ((object)(resources.GetObject("appearance18.Image")));
			appearance18.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance18.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn47.CellButtonAppearance = appearance18;
			ultraGridColumn47.Header.Caption = "...";
			ultraGridColumn47.Header.VisiblePosition = 3;
			ultraGridColumn47.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn47.Width = 25;
			ultraGridColumn48.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn48.Header.Caption = "...";
			ultraGridColumn48.Header.VisiblePosition = 5;
			ultraGridColumn48.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn48.Width = 16;
			ultraGridColumn49.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn49.Header.Caption = "...";
			ultraGridColumn49.Header.VisiblePosition = 7;
			ultraGridColumn49.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn49.Width = 16;
			ultraGridColumn50.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn50.Header.Caption = "...";
			ultraGridColumn50.Header.VisiblePosition = 9;
			ultraGridColumn50.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn50.Width = 16;
			ultraGridColumn51.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn51.Header.Caption = "...";
			ultraGridColumn51.Header.VisiblePosition = 11;
			ultraGridColumn51.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn51.Width = 16;
			ultraGridColumn52.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn52.Header.Caption = "...";
			ultraGridColumn52.Header.VisiblePosition = 13;
			ultraGridColumn52.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn52.Width = 16;
			ultraGridColumn53.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn53.Header.Caption = "...";
			ultraGridColumn53.Header.VisiblePosition = 15;
			ultraGridColumn53.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn53.Width = 16;
			ultraGridColumn54.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn54.Header.Caption = "...";
			ultraGridColumn54.Header.VisiblePosition = 17;
			ultraGridColumn54.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn54.Width = 16;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn54});
			ultraGridBand3.Header.Caption = "Semana 3";
			ultraGridBand3.HeaderVisible = true;
			this.uGridSemana3.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana3.DisplayLayout.Override.ActiveRowAppearance = appearance19;
			this.uGridSemana3.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance20.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana3.DisplayLayout.Override.CardAreaAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance21.FontData.BoldAsString = "True";
			appearance21.FontData.Name = "Arial";
			appearance21.FontData.SizeInPoints = 7.5F;
			appearance21.ForeColor = System.Drawing.Color.White;
			appearance21.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana3.DisplayLayout.Override.HeaderAppearance = appearance21;
			this.uGridSemana3.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance22.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance22.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana3.DisplayLayout.Override.RowAlternateAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana3.DisplayLayout.Override.RowSelectorAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana3.DisplayLayout.Override.SelectedRowAppearance = appearance24;
			this.uGridSemana3.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana3.Location = new System.Drawing.Point(0, 240);
			this.uGridSemana3.Name = "uGridSemana3";
			this.uGridSemana3.Size = new System.Drawing.Size(1018, 104);
			this.uGridSemana3.TabIndex = 765;
			this.uGridSemana3.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana3_AfterRowInsert);
			this.uGridSemana3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana3_KeyDown);
			this.uGridSemana3.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana3_ClickCellButton);
			// 
			// uGridSemana2
			// 
			this.uGridSemana2.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana2.DataSource = this.ultraDataSource4;
			appearance25.BackColor = System.Drawing.Color.White;
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana2.DisplayLayout.Appearance = appearance25;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 116;
			ultraGridColumn56.Header.VisiblePosition = 1;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 89;
			ultraGridColumn57.Header.VisiblePosition = 2;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 46;
			ultraGridColumn58.Header.VisiblePosition = 4;
			ultraGridColumn58.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn58.Width = 120;
			ultraGridColumn59.Header.VisiblePosition = 6;
			ultraGridColumn59.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn59.Width = 120;
			ultraGridColumn60.Header.VisiblePosition = 8;
			ultraGridColumn60.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn60.Width = 120;
			ultraGridColumn61.Header.VisiblePosition = 10;
			ultraGridColumn61.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn61.Width = 120;
			ultraGridColumn62.Header.VisiblePosition = 12;
			ultraGridColumn62.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn62.Width = 120;
			ultraGridColumn63.Header.VisiblePosition = 14;
			ultraGridColumn63.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn63.Width = 120;
			ultraGridColumn64.Header.VisiblePosition = 16;
			ultraGridColumn64.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn64.Width = 120;
			ultraGridColumn65.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance26.Image = ((object)(resources.GetObject("appearance26.Image")));
			appearance26.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance26.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn65.CellButtonAppearance = appearance26;
			ultraGridColumn65.Header.Caption = "...";
			ultraGridColumn65.Header.VisiblePosition = 3;
			ultraGridColumn65.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn65.Width = 25;
			ultraGridColumn66.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn66.Header.Caption = "...";
			ultraGridColumn66.Header.VisiblePosition = 5;
			ultraGridColumn66.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn66.Width = 16;
			ultraGridColumn67.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn67.Header.Caption = "...";
			ultraGridColumn67.Header.VisiblePosition = 7;
			ultraGridColumn67.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn67.Width = 16;
			ultraGridColumn68.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn68.Header.Caption = "...";
			ultraGridColumn68.Header.VisiblePosition = 9;
			ultraGridColumn68.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn68.Width = 16;
			ultraGridColumn69.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn69.Header.Caption = "...";
			ultraGridColumn69.Header.VisiblePosition = 13;
			ultraGridColumn69.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn69.Width = 16;
			ultraGridColumn70.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn70.Header.Caption = "...";
			ultraGridColumn70.Header.VisiblePosition = 11;
			ultraGridColumn70.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn70.Width = 16;
			ultraGridColumn71.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn71.Header.Caption = "...";
			ultraGridColumn71.Header.VisiblePosition = 15;
			ultraGridColumn71.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn71.Width = 16;
			ultraGridColumn72.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn72.Header.Caption = "...";
			ultraGridColumn72.Header.VisiblePosition = 17;
			ultraGridColumn72.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn72.Width = 16;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn72});
			ultraGridBand4.Header.Caption = "Semana 2";
			ultraGridBand4.HeaderVisible = true;
			this.uGridSemana2.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana2.DisplayLayout.Override.ActiveRowAppearance = appearance27;
			this.uGridSemana2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance28.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana2.DisplayLayout.Override.CardAreaAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance29.FontData.BoldAsString = "True";
			appearance29.FontData.Name = "Arial";
			appearance29.FontData.SizeInPoints = 7.5F;
			appearance29.ForeColor = System.Drawing.Color.White;
			appearance29.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana2.DisplayLayout.Override.HeaderAppearance = appearance29;
			this.uGridSemana2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance30.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance30.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana2.DisplayLayout.Override.RowAlternateAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana2.DisplayLayout.Override.RowSelectorAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana2.DisplayLayout.Override.SelectedRowAppearance = appearance32;
			this.uGridSemana2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana2.Location = new System.Drawing.Point(0, 136);
			this.uGridSemana2.Name = "uGridSemana2";
			this.uGridSemana2.Size = new System.Drawing.Size(1018, 104);
			this.uGridSemana2.TabIndex = 764;
			this.uGridSemana2.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana2_AfterRowInsert);
			this.uGridSemana2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana2_KeyDown);
			this.uGridSemana2.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana2_ClickCellButton);
			this.uGridSemana2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSemana2_InitializeLayout);
			// 
			// uGridSemana1
			// 
			this.uGridSemana1.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana1.DataSource = this.ultraDataSource1;
			appearance33.BackColor = System.Drawing.Color.White;
			appearance33.ForeColor = System.Drawing.Color.Black;
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana1.DisplayLayout.Appearance = appearance33;
			ultraGridColumn73.Header.VisiblePosition = 0;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn73.Width = 116;
			ultraGridColumn74.Header.VisiblePosition = 1;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn74.Width = 89;
			ultraGridColumn75.Header.VisiblePosition = 2;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 46;
			ultraGridColumn76.Header.VisiblePosition = 4;
			ultraGridColumn76.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn76.Width = 120;
			ultraGridColumn77.Header.VisiblePosition = 6;
			ultraGridColumn77.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn77.Width = 120;
			ultraGridColumn78.Header.VisiblePosition = 8;
			ultraGridColumn78.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn78.Width = 120;
			ultraGridColumn79.Header.VisiblePosition = 10;
			ultraGridColumn79.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn79.Width = 120;
			ultraGridColumn80.Header.VisiblePosition = 12;
			ultraGridColumn80.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn80.Width = 120;
			ultraGridColumn81.Header.VisiblePosition = 14;
			ultraGridColumn81.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn81.Width = 120;
			ultraGridColumn82.Header.VisiblePosition = 16;
			ultraGridColumn82.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn82.Width = 120;
			ultraGridColumn83.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance34.Image = ((object)(resources.GetObject("appearance34.Image")));
			appearance34.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance34.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn83.CellButtonAppearance = appearance34;
			ultraGridColumn83.Header.Caption = "...";
			ultraGridColumn83.Header.VisiblePosition = 3;
			ultraGridColumn83.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn83.Width = 25;
			ultraGridColumn84.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn84.Header.Caption = "...";
			ultraGridColumn84.Header.VisiblePosition = 5;
			ultraGridColumn84.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn84.Width = 16;
			ultraGridColumn85.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn85.Header.Caption = "...";
			ultraGridColumn85.Header.VisiblePosition = 7;
			ultraGridColumn85.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn85.Width = 16;
			ultraGridColumn86.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn86.Header.Caption = "...";
			ultraGridColumn86.Header.VisiblePosition = 9;
			ultraGridColumn86.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn86.Width = 16;
			ultraGridColumn87.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn87.Header.Caption = "...";
			ultraGridColumn87.Header.VisiblePosition = 11;
			ultraGridColumn87.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn87.Width = 16;
			ultraGridColumn88.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn88.Header.Caption = "...";
			ultraGridColumn88.Header.VisiblePosition = 13;
			ultraGridColumn88.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn88.Width = 16;
			ultraGridColumn89.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn89.Header.Caption = "...";
			ultraGridColumn89.Header.VisiblePosition = 15;
			ultraGridColumn89.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn89.Width = 16;
			ultraGridColumn90.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn90.Header.Caption = "...";
			ultraGridColumn90.Header.VisiblePosition = 17;
			ultraGridColumn90.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn90.Width = 16;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn90});
			ultraGridBand5.Header.Caption = "Semana 1";
			ultraGridBand5.HeaderVisible = true;
			this.uGridSemana1.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance35.ForeColor = System.Drawing.Color.Black;
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana1.DisplayLayout.Override.ActiveRowAppearance = appearance35;
			this.uGridSemana1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance36.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana1.DisplayLayout.Override.CardAreaAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance37.FontData.BoldAsString = "True";
			appearance37.FontData.Name = "Arial";
			appearance37.FontData.SizeInPoints = 7.5F;
			appearance37.ForeColor = System.Drawing.Color.White;
			appearance37.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana1.DisplayLayout.Override.HeaderAppearance = appearance37;
			this.uGridSemana1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance38.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance38.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana1.DisplayLayout.Override.RowAlternateAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana1.DisplayLayout.Override.RowSelectorAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana1.DisplayLayout.Override.SelectedRowAppearance = appearance40;
			this.uGridSemana1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana1.Location = new System.Drawing.Point(0, 32);
			this.uGridSemana1.Name = "uGridSemana1";
			this.uGridSemana1.Size = new System.Drawing.Size(1018, 104);
			this.uGridSemana1.TabIndex = 763;
			this.uGridSemana1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana1_AfterRowInsert);
			this.uGridSemana1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana1_KeyDown);
			this.uGridSemana1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana1_ClickCellButton);
			this.uGridSemana1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSemana1_InitializeLayout);
			// 
			// cmbNombres
			// 
			this.cmbNombres.CausesValidation = false;
			this.cmbNombres.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombres.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombres.DataSource = this.ultraDataSource9;
			ultraGridColumn91.Header.VisiblePosition = 0;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 1;
			ultraGridColumn93.Header.VisiblePosition = 2;
			ultraGridColumn93.Width = 180;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn91,
																										 ultraGridColumn92,
																										 ultraGridColumn93});
			this.cmbNombres.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbNombres.DisplayMember = "Nombre";
			this.cmbNombres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombres.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombres.Enabled = false;
			this.cmbNombres.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombres.Location = new System.Drawing.Point(64, 8);
			this.cmbNombres.MaxDropDownItems = 30;
			this.cmbNombres.Name = "cmbNombres";
			this.cmbNombres.Size = new System.Drawing.Size(200, 21);
			this.cmbNombres.TabIndex = 761;
			this.cmbNombres.ValueMember = "idPersonal";
			this.cmbNombres.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNombres_InitializeLayout);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn66.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn66,
																																 ultraDataColumn67,
																																 ultraDataColumn68});
			// 
			// cmbLocales
			// 
			this.cmbLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLocales.DataSource = this.ultraDataSource3;
			ultraGridColumn94.Header.VisiblePosition = 0;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Width = 57;
			ultraGridColumn95.Header.VisiblePosition = 1;
			ultraGridColumn95.Width = 100;
			ultraGridColumn96.Header.VisiblePosition = 2;
			ultraGridColumn96.Width = 180;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn94,
																										 ultraGridColumn95,
																										 ultraGridColumn96});
			this.cmbLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbLocales.DisplayMember = "Nombre";
			this.cmbLocales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbLocales.Location = new System.Drawing.Point(377, 307);
			this.cmbLocales.Name = "cmbLocales";
			this.cmbLocales.Size = new System.Drawing.Size(303, 112);
			this.cmbLocales.TabIndex = 770;
			this.cmbLocales.ValueMember = "Bodega";
			this.cmbLocales.Visible = false;
			// 
			// uGridSemana6
			// 
			this.uGridSemana6.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemana6.DataSource = this.ultraDataSource8;
			appearance41.BackColor = System.Drawing.Color.White;
			appearance41.ForeColor = System.Drawing.Color.Black;
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana6.DisplayLayout.Appearance = appearance41;
			ultraGridColumn97.Header.VisiblePosition = 0;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 116;
			ultraGridColumn98.Header.VisiblePosition = 1;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn98.Width = 89;
			ultraGridColumn99.Header.VisiblePosition = 2;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn99.Width = 46;
			ultraGridColumn100.Header.VisiblePosition = 4;
			ultraGridColumn100.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn100.Width = 120;
			ultraGridColumn101.Header.VisiblePosition = 6;
			ultraGridColumn101.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn101.Width = 120;
			ultraGridColumn102.Header.VisiblePosition = 8;
			ultraGridColumn102.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn102.Width = 120;
			ultraGridColumn103.Header.VisiblePosition = 10;
			ultraGridColumn103.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn103.Width = 120;
			ultraGridColumn104.Header.VisiblePosition = 12;
			ultraGridColumn104.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn104.Width = 120;
			ultraGridColumn105.Header.VisiblePosition = 14;
			ultraGridColumn105.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn105.Width = 120;
			ultraGridColumn106.Header.VisiblePosition = 16;
			ultraGridColumn106.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn106.Width = 120;
			ultraGridColumn107.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance42.Image = ((object)(resources.GetObject("appearance42.Image")));
			appearance42.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance42.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn107.CellButtonAppearance = appearance42;
			ultraGridColumn107.Header.Caption = "...";
			ultraGridColumn107.Header.VisiblePosition = 3;
			ultraGridColumn107.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn107.Width = 25;
			ultraGridColumn108.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn108.Header.Caption = "...";
			ultraGridColumn108.Header.VisiblePosition = 5;
			ultraGridColumn108.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn108.Width = 16;
			ultraGridColumn109.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn109.Header.Caption = "...";
			ultraGridColumn109.Header.VisiblePosition = 7;
			ultraGridColumn109.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn109.Width = 16;
			ultraGridColumn110.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn110.Header.Caption = "...";
			ultraGridColumn110.Header.VisiblePosition = 9;
			ultraGridColumn110.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn110.Width = 16;
			ultraGridColumn111.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn111.Header.Caption = "...";
			ultraGridColumn111.Header.VisiblePosition = 11;
			ultraGridColumn111.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn111.Width = 16;
			ultraGridColumn112.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn112.Header.Caption = "...";
			ultraGridColumn112.Header.VisiblePosition = 13;
			ultraGridColumn112.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn112.Width = 16;
			ultraGridColumn113.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn113.Header.Caption = "...";
			ultraGridColumn113.Header.VisiblePosition = 15;
			ultraGridColumn113.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn113.Width = 16;
			ultraGridColumn114.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn114.Header.Caption = "...";
			ultraGridColumn114.Header.VisiblePosition = 17;
			ultraGridColumn114.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn114.Width = 16;
			ultraGridBand8.Columns.AddRange(new object[] {
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
																										 ultraGridColumn114});
			ultraGridBand8.Header.Caption = "Semana 6";
			ultraGridBand8.HeaderVisible = true;
			this.uGridSemana6.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			appearance43.ForeColor = System.Drawing.Color.Black;
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemana6.DisplayLayout.Override.ActiveRowAppearance = appearance43;
			this.uGridSemana6.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridSemana6.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSemana6.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance44.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemana6.DisplayLayout.Override.CardAreaAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance45.FontData.BoldAsString = "True";
			appearance45.FontData.Name = "Arial";
			appearance45.FontData.SizeInPoints = 7.5F;
			appearance45.ForeColor = System.Drawing.Color.White;
			appearance45.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemana6.DisplayLayout.Override.HeaderAppearance = appearance45;
			this.uGridSemana6.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance46.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance46.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana6.DisplayLayout.Override.RowAlternateAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana6.DisplayLayout.Override.RowSelectorAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemana6.DisplayLayout.Override.SelectedRowAppearance = appearance48;
			this.uGridSemana6.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemana6.Location = new System.Drawing.Point(0, 552);
			this.uGridSemana6.Name = "uGridSemana6";
			this.uGridSemana6.Size = new System.Drawing.Size(1018, 104);
			this.uGridSemana6.TabIndex = 782;
			this.uGridSemana6.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridSemana6_AfterRowInsert);
			this.uGridSemana6.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemana6_KeyDown);
			this.uGridSemana6.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemana6_ClickCellButton);
			// 
			// frmPlanificacionVentas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(947, 509);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.txtIdPlanificacionVentas);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.dtpMes);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.cmbNombres);
			this.Controls.Add(this.cmbLocales);
			this.Controls.Add(this.btnAprobarSemana6);
			this.Controls.Add(this.btnAprobarSemana5);
			this.Controls.Add(this.btnAprobarSemana4);
			this.Controls.Add(this.btnAprobarSemana3);
			this.Controls.Add(this.btnAprobarSemana2);
			this.Controls.Add(this.btnAprobarSemana1);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.uGridSemana5);
			this.Controls.Add(this.uGridSemana4);
			this.Controls.Add(this.uGridSemana3);
			this.Controls.Add(this.uGridSemana2);
			this.Controls.Add(this.uGridSemana1);
			this.Controls.Add(this.uGridSemana6);
			this.Name = "frmPlanificacionVentas";
			this.Text = "Planificación de Ventas";
			this.Load += new System.EventHandler(this.frmPlanificacionVentas_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPlanificacionVentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemana6)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		int idBloqueaTransacciones = 0;
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmPlanificacionVentas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0729");
			
			if (!Funcion.Permiso("724", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Planificación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));		
				return;
			}
			
			DateTime dtFecha = DateTime.Parse(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
				
			this.cmbLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 43", MenuLatinium.IdUsuario));
			this.cmbNombres.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idPersonal, (ApellidoPaterno + ' ' + ApellidoMaterno + ' ' + PrimerNombre + ' ' + SegundoNombre) As Nombre From Nomina n Inner Join RolCargo rc On rc.idCargo = n.idCargo Where rc.Com_IdCargo In (8, 9, 12, 18) Order By Nombre"));

			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana1);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana2);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana3);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana4);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana5);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana6);			
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 6", idBloqueaTransacciones, (int)this.txtIdPlanificacionVentas.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			this.cmbNombres.Enabled = false;
			this.dtpMes.Enabled = false;
			this.dtpMes.Value = DateTime.Parse("01/01/2017");

			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana1, 0, 1, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana2, 0, 2, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana3, 0, 3, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana4, 0, 4, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana5, 0, 5, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana6, 0, 6, 2, cdsSeteoF);

			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana1);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana2);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana3);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana4);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana5);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana6);

			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana1, false);
			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana2, false);
			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana3, false);
			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana4, false);
			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana5, false);
			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana6, false);

			bNuevo = false;
			bEdicion = false;

			this.txtIdPlanificacionVentas.Value = 0;
			this.cmbNombres.Value = System.DBNull.Value;
			this.txtEstado.Value = 0;
			idBloqueaTransacciones = 0;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;

			this.btnAprobarSemana1.Enabled = false;
			this.btnAprobarSemana2.Enabled = false;
			this.btnAprobarSemana3.Enabled = false;
			this.btnAprobarSemana4.Enabled = false;
			this.btnAprobarSemana5.Enabled = false;
			this.btnAprobarSemana6.Enabled = false;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			string sSQLVal = string.Format("Exec ValidaHorariosDeCarpa {0}, '{1}', 2", (int)this.cmbNombres.Value, Convert.ToDateTime(this.dtpMes.Value).ToString("yyyyMMdd"));
			if (Funcion.bEscalarSQL(cdsSeteoF, sSQLVal, true) && bNuevo)
			{
				MessageBox.Show("Ya existe una planificacion de carpa del local y periodo seleccionado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);        				
				this.dtpMes.Focus();
				return;
			}

			string sSQL = string.Format("Exec GuardaHorariosCarpa {0}, {1}, '{2}', '', 2", 
				(int)this.txtIdPlanificacionVentas.Value, (int)this.cmbNombres.Value, Convert.ToDateTime(this.dtpMes.Value).ToString("yyyyMMdd"));
			this.txtIdPlanificacionVentas.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (this.uGridSemana1.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana1, (int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF);			
			if (this.uGridSemana2.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana2, (int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF);
			if (this.uGridSemana3.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana3, (int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF);
			if (this.uGridSemana4.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana4, (int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF);
			if (this.uGridSemana5.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana5, (int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF);
			if (this.uGridSemana6.Rows.Count > 0) FuncionesProcedimientos.GuardaDetalleHorariosDeCarpa(this.uGridSemana6, (int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF);

			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana1, (int)this.txtIdPlanificacionVentas.Value, 1, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana2, (int)this.txtIdPlanificacionVentas.Value, 2, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana3, (int)this.txtIdPlanificacionVentas.Value, 3, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana4, (int)this.txtIdPlanificacionVentas.Value, 4, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana5, (int)this.txtIdPlanificacionVentas.Value, 5, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana6, (int)this.txtIdPlanificacionVentas.Value, 6, 2, cdsSeteoF);

			if (idBloqueaTransacciones > 0)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 6", 
					idBloqueaTransacciones, (int)this.txtIdPlanificacionVentas.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			this.cmbNombres.Enabled = false;
			this.dtpMes.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana1);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana2);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana3);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana4);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana5);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana6);

			bNuevo = false;
			bEdicion = false;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;			
			this.btnGuardar.Enabled = false;

			if ((int)this.txtEstado.Value == 0) 
				this.btnEditar.Enabled = true;
			
			this.btnAprobarSemana1.Enabled = true;
			this.btnAprobarSemana2.Enabled = true;
			this.btnAprobarSemana3.Enabled = true;
			this.btnAprobarSemana4.Enabled = true;
			this.btnAprobarSemana5.Enabled = true;
			this.btnAprobarSemana6.Enabled = true;			

			this.btnCancelar.Enabled = true;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			
			this.btnAprobarSemana1.Enabled = false;
			this.btnAprobarSemana2.Enabled = false;
			this.btnAprobarSemana3.Enabled = false;
			this.btnAprobarSemana4.Enabled = false;
			this.btnAprobarSemana5.Enabled = false;
			this.btnAprobarSemana6.Enabled = false;

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 2", 1, (int)this.txtIdPlanificacionVentas.Value)) > 0)
			{
				if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana1);	
			}
			else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana1);
			
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 2", 2, (int)this.txtIdPlanificacionVentas.Value)) > 0)
			{
				if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana2);	
			}
			else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana2);

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 2", 3, (int)this.txtIdPlanificacionVentas.Value)) > 0)
			{
				if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana3);
			}
			else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana3);

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 2", 4, (int)this.txtIdPlanificacionVentas.Value)) > 0)
			{
				if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana4);
			}
			else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana4);

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 2", 5, (int)this.txtIdPlanificacionVentas.Value)) > 0)
			{
				if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana5);	
			}
			else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana5);

			if (this.uGridSemana6.Rows.Count > 0) 
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdASHC {0}, {1}, 2", 6, (int)this.txtIdPlanificacionVentas.Value)) > 0)
				{
					if (miAcceso.BAprobarCarpa) FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana6);	
				}
				else FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana6);
			}

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 6", (int)this.txtIdPlanificacionVentas.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			bNuevo = false;
			bEdicion = true;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana1);
			FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana2);
			FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana3);
			FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana4);
			FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana5);			

			this.dtpMes.Enabled = true;			

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

			this.btnAprobarSemana1.Enabled = false;
			this.btnAprobarSemana2.Enabled = false;
			this.btnAprobarSemana3.Enabled = false;
			this.btnAprobarSemana4.Enabled = false;
			this.btnAprobarSemana5.Enabled = false;
			this.btnAprobarSemana6.Enabled = false;

			//this.cmbNombres.Enabled = true;
			this.cmbNombres.Value = 
				Funcion.iEscalarSQL(cdsSeteoF, 
				string.Format("If Exists(Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0})) Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario where idUsuario = {0}) Else Select 0", 
				MenuLatinium.IdUsuario));
						
			this.dtpMes.Value = DateTime.Today;//DateTime.Parse("30/04/2017");//DateTime.Today;

			bNuevo = true;
			bEdicion = true;

			this.dtpMes.Focus();
		}

		private void Consultar(int idPlanificacionVentas)
		{			
			DateTime dtMes = DateTime.Today;

			string sSQL = string.Format("Exec ConsultaHorariosCarpaIndivisual {0}, 2", idPlanificacionVentas);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdPlanificacionVentas.Value = idPlanificacionVentas;
				this.cmbNombres.Value = dr.GetInt32(1);
				dtMes = dr.GetDateTime(2);
			}
			dr.Close();
		
			this.dtpMes.Value = dtMes; 

			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana1, idPlanificacionVentas, 1, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana2, idPlanificacionVentas, 2, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana3, idPlanificacionVentas, 3, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana4, idPlanificacionVentas, 4, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana5, idPlanificacionVentas, 5, 2, cdsSeteoF);
			FuncionesProcedimientos.ConsultaDetalleHorariosDeCarpa(this.uGridSemana6, idPlanificacionVentas, 6, 2, cdsSeteoF);

			this.uGridSemana1.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana1.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF, 1);
			this.uGridSemana2.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana2.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF, 2);
			this.uGridSemana3.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana3.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF, 3);
			this.uGridSemana4.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana4.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF, 4);
			this.uGridSemana5.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana5.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF, 5);
			this.uGridSemana6.DisplayLayout.Bands[0].Header.Caption = this.uGridSemana6.DisplayLayout.Bands[0].Header.Caption + FuncionesProcedimientos.sMensajeAprobado((int)this.txtIdPlanificacionVentas.Value, 2, cdsSeteoF, 6);

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			
			this.btnAprobarSemana1.Enabled = true;
			this.btnAprobarSemana2.Enabled = true;
			this.btnAprobarSemana3.Enabled = true;
			this.btnAprobarSemana4.Enabled = true;
			this.btnAprobarSemana5.Enabled = true;
			this.btnAprobarSemana6.Enabled = true;
			
			this.btnCancelar.Enabled = true;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaPlanificacionVentas Busqueda = new frmBuscaPlanificacionVentas())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consultar((int)Busqueda.uGridHorarios.ActiveRow.Cells["idPlanificacionVentas"].Value);
				}
			}
		}

		private void OcultaCeldasSemana1()
		{
			DateTime dtInicio = Convert.ToDateTime("01/" + Convert.ToDateTime(dtpMes.Value).Month.ToString() + "/" + Convert.ToDateTime(dtpMes.Value).Year.ToString());
			int iNumDiaSemana = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaNumDiaSemana '{0}'", dtInicio.ToString("yyyyMMdd"))) - 1; 
			if (iNumDiaSemana == 0) iNumDiaSemana = 7;
			iNumDiaSemana = iNumDiaSemana + 3;

			int iContador = 3;

			string sNombreCol = "";

			for (int j = iContador; j <= 9; j++)
			{
				sNombreCol  = this.uGridSemana1.DisplayLayout.Bands[0].Columns[j].Header.Caption;

				if (j < iNumDiaSemana - 1) 
				{
					this.uGridSemana1.DisplayLayout.Bands[0].Columns[sNombreCol].Hidden = true;
					this.uGridSemana1.DisplayLayout.Bands[0].Columns["btn" + sNombreCol].Hidden = true;
				}
			}
		}

		private void OcultaCeldasSemana5()
		{
			DateTime dtFecha = (DateTime)this.dtpMes.Value;
			int UDia = DateTime.DaysInMonth(dtFecha.Year, dtFecha.Month);
			int iMes = dtFecha.Month;
			int iAño = dtFecha.Year;
			int iSemanas = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CalculaSemanasXMes '{0}'", dtFecha.ToString("yyyyMMdd")));
			DateTime dtFin = Convert.ToDateTime(UDia.ToString() + "/" + iMes.ToString() + "/" + iAño.ToString());
			
			int iNumDiaSemana = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaNumDiaSemana '{0}'", dtFin.ToString("yyyyMMdd"))) - 1; 
			if (iNumDiaSemana == 0) 
			{
				iNumDiaSemana = 7;				
			}	

			iNumDiaSemana = iNumDiaSemana + 3;

			int iContador = 3;		

			if (iSemanas == 5) 
			{
				for (int j = iContador; j <= 9; j++)
				{
					if ((iNumDiaSemana - 1) < j) 
					{
						this.uGridSemana5.DisplayLayout.Bands[0].Columns[j].Hidden = true;						
					}
				}
			}

			if (iSemanas == 6) 
			{
				for (int j = iContador; j <= 9; j++)
				{
					if ((iNumDiaSemana - 1) < j) 
					{
						this.uGridSemana6.DisplayLayout.Bands[0].Columns[j].Hidden = true;						
					}
				}
			}
		}

		private void DiasMesCeldas ()
		{
			FuncionesProcedimientos.EstadoGrids(false, this.uGridSemana6);

			DateTime dtInicio = Convert.ToDateTime("01/" + Convert.ToDateTime(dtpMes.Value).Month.ToString() + "/" + Convert.ToDateTime(dtpMes.Value).Year.ToString());
			int UDia = DateTime.DaysInMonth(dtInicio.Year, dtInicio.Month);
			int iNumDiaSemana = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaNumDiaSemana '{0}'", dtInicio.ToString("yyyyMMdd"))) - 1; 
			if (iNumDiaSemana == 0) iNumDiaSemana = 7;
			iNumDiaSemana = iNumDiaSemana + 3;
			
			int iDia = 1;		
			int iContador = 3;	
			int iSemanas = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CalculaSemanasXMes '{0}'", dtInicio.ToString("yyyyMMdd")));

			for (int j = iContador; j <= 9; j++)
			{				
				if (!this.uGridSemana1.DisplayLayout.Bands[0].Columns[j].Hidden)
				{
					this.uGridSemana1.DisplayLayout.Bands[0].Columns[j].Header.Caption = this.uGridSemana1.DisplayLayout.Bands[0].Columns[j].Header.Caption + "-" + iDia.ToString();
					iDia++;
				}
			}

			iContador = 3;
			for (int j = iContador; j <= 9; j++)
			{				
				if (!this.uGridSemana2.DisplayLayout.Bands[0].Columns[j].Hidden)
				{
					this.uGridSemana2.DisplayLayout.Bands[0].Columns[j].Header.Caption = this.uGridSemana2.DisplayLayout.Bands[0].Columns[j].Header.Caption + "-" + iDia.ToString();
					iDia++;
				}
			}

			iContador = 3;
			for (int j = iContador; j <= 9; j++)
			{	
				if (!this.uGridSemana3.DisplayLayout.Bands[0].Columns[j].Hidden)
				{
					this.uGridSemana3.DisplayLayout.Bands[0].Columns[j].Header.Caption = this.uGridSemana3.DisplayLayout.Bands[0].Columns[j].Header.Caption + "-" + iDia.ToString();
					iDia++;
				}
			}

			iContador = 3;
			for (int j = iContador; j <= 9; j++)
			{
				if (!this.uGridSemana4.DisplayLayout.Bands[0].Columns[j].Hidden)
				{
					if (iDia <= UDia)
						this.uGridSemana4.DisplayLayout.Bands[0].Columns[j].Header.Caption = this.uGridSemana4.DisplayLayout.Bands[0].Columns[j].Header.Caption + "-" + iDia.ToString();
					iDia++;
				}
			}
			
			iContador = 3;
			for (int j = iContador; j <= 9; j++)
			{
				if (!this.uGridSemana5.DisplayLayout.Bands[0].Columns[j].Hidden)
				{
					if (iDia <= UDia)					
						this.uGridSemana5.DisplayLayout.Bands[0].Columns[j].Header.Caption = this.uGridSemana5.DisplayLayout.Bands[0].Columns[j].Header.Caption + "-" + iDia.ToString();
					iDia++;
				}
			}

			if (iSemanas == 6)
			{
				FuncionesProcedimientos.EstadoGrids(true, this.uGridSemana6);

				iContador = 3;
				for (int j = iContador; j <= 9; j++)
				{
					if (!this.uGridSemana6.DisplayLayout.Bands[0].Columns[j].Hidden)
					{
						if (iDia <= UDia)
							this.uGridSemana6.DisplayLayout.Bands[0].Columns[j].Header.Caption = this.uGridSemana6.DisplayLayout.Bands[0].Columns[j].Header.Caption + "-" + iDia.ToString();
						iDia++;
					}
				}
			}
		}

		private void dtpMes_ValueChanged(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime)this.dtpMes.Value;
			int UDia = DateTime.DaysInMonth(dtFecha.Year, dtFecha.Month);
			int iMes = dtFecha.Month;
			int iAño = dtFecha.Year;

			DateTime dtInicio = Convert.ToDateTime("01/" + iMes.ToString() + "/" + iAño.ToString());
			DateTime dtFin = Convert.ToDateTime(UDia.ToString() + "/" + iMes.ToString() + "/" + iAño.ToString());
			
			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana1, false);
			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana2, false);
			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana3, false);
			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana4, false);
			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana5, false);
			FuncionesProcedimientos.MostrarOcultar(this.uGridSemana6, false);
			
			OcultaCeldasSemana1();
			OcultaCeldasSemana5();
			DiasMesCeldas ();			
		}

		private void GuardaAprobacion(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int iSemana)
		{
			if (uGrid.Rows.Count == 0)
			{
				MessageBox.Show(string.Format("La semana {0} no tiene registros", iSemana), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			//			string sSQL = string.Format("Exec RetornaCantidadPersonalPorLocal {0}, '{1}'", (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtpMes.Value).ToString("yyyyMMdd"));
			//			int iPersonal = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			//
			//			if (uGrid.Rows.Count < iPersonal)
			//			{
			//				MessageBox.Show(string.Format("El numero de registros ingresados en la semana {0} es menor al numero de personal asignado al local '{1}' en el periodo '{2}'", 
			//					iSemana, this.cmbBodega.Text.ToString(), Convert.ToDateTime(this.dtpMes.Value).ToString("MMMM-yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				return;
			//			}

			int idAprobacionSemanalHorariosDeCarpa = Funcion.iEscalarSQL(cdsSeteoF, 
				string.Format("Exec RetornaIdASHC {0}, {1}, 2", iSemana, (int)this.txtIdPlanificacionVentas.Value));

			if (idAprobacionSemanalHorariosDeCarpa == 0 && !miAcceso.BAprobarCarpa)
			{
				MessageBox.Show(string.Format("La semana {0} de carpa no ha sido Aprobada\n\nConsulte con el Jefe Nacional De Ventas", iSemana), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 6", (int)this.txtIdPlanificacionVentas.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			using (frmNotasAprobacionCarpa Notas = new frmNotasAprobacionCarpa((int)this.txtIdPlanificacionVentas.Value, iSemana, miAcceso.BAprobarCarpa, 2))
			{
				if (DialogResult.OK == Notas.ShowDialog())
				{
					try
					{							
						string sSQL = string.Format("Exec GuardaAprobacionSemanalHorariosDeCarpa {0}, {1}, {2}, '{3}', 2", 
							idAprobacionSemanalHorariosDeCarpa, (int)this.txtIdPlanificacionVentas.Value, iSemana, Notas.txtNotas.Text.ToString());
						idAprobacionSemanalHorariosDeCarpa = Funcion.iEscalarSQL(cdsSeteoF, sSQL);						
					}
					catch(Exception Exc)
					{
						MessageBox.Show(string.Format("Error al Aprobar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							
					}
				}
			}

			sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 6", idBloqueaTransacciones, (int)this.txtIdPlanificacionVentas.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
		}

		private void btnAprobarSemana1_Click(object sender, System.EventArgs e)
		{
			GuardaAprobacion(this.uGridSemana1, 1);
		}

		private void btnAprobarSemana2_Click(object sender, System.EventArgs e)
		{
			GuardaAprobacion(this.uGridSemana2, 2);
		}

		private void btnAprobarSemana3_Click(object sender, System.EventArgs e)
		{
			GuardaAprobacion(this.uGridSemana3, 3);
		}

		private void btnAprobarSemana4_Click(object sender, System.EventArgs e)
		{
			GuardaAprobacion(this.uGridSemana4, 4);
		}

		private void btnAprobarSemana5_Click(object sender, System.EventArgs e)
		{
			GuardaAprobacion(this.uGridSemana5, 5);
		}

		private void btnAprobarSemana6_Click(object sender, System.EventArgs e)
		{
			GuardaAprobacion(this.uGridSemana6, 6);
		}

		private void ResumenDiario(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if ((int)this.txtIdPlanificacionVentas.Value == 0) return;

			if (e.Cell.Row.Cells["idDetallePlanificacionVentas"].Value == System.DBNull.Value) return;

			if ((int)e.Cell.Row.Cells["idDetallePlanificacionVentas"].Value == 0) return;

			int iBodega = 0;
			string sNomCol = "";
			int iDia = 0;

			if (e.Cell.Column.ToString() == "btnLunes")
			{
				if ((int)e.Cell.Row.Cells["Lunes"].Value == 0) return;

				iBodega = (int)e.Cell.Row.Cells["Lunes"].Value;

				sNomCol = e.Cell.Row.Cells["Lunes"].Column.Header.Caption;

				int iLength = 1;
				if (sNomCol.Length == 8) iLength = 2;

				iDia = Convert.ToInt32(sNomCol.Substring(6, iLength));				
			}

			if (e.Cell.Column.ToString() == "btnMartes")
			{
				if ((int)e.Cell.Row.Cells["Martes"].Value == 0) return;

				iBodega = (int)e.Cell.Row.Cells["Martes"].Value;

				sNomCol = e.Cell.Row.Cells["Martes"].Column.Header.Caption;

				int iLength = 1;
				if (sNomCol.Length == 9) iLength = 2;

				iDia = Convert.ToInt32(sNomCol.Substring(7, iLength));
			}

			if (e.Cell.Column.ToString() == "btnMiercoles")
			{
				if ((int)e.Cell.Row.Cells["Miercoles"].Value == 0) return;

				iBodega = (int)e.Cell.Row.Cells["Miercoles"].Value;

				sNomCol = e.Cell.Row.Cells["Miercoles"].Column.Header.Caption;

				int iLength = 1;
				if (sNomCol.Length == 12) iLength = 2;
        
				iDia = Convert.ToInt32(sNomCol.Substring(10, iLength));
			}

			if (e.Cell.Column.ToString() == "btnJueves")
			{
				if ((int)e.Cell.Row.Cells["Jueves"].Value == 0) return;

				iBodega = (int)e.Cell.Row.Cells["Jueves"].Value;

				sNomCol = e.Cell.Row.Cells["Jueves"].Column.Header.Caption;

				int iLength = 1;
				if (sNomCol.Length == 9) iLength = 2;
				
				iDia = Convert.ToInt32(sNomCol.Substring(7, iLength));
			}

			if (e.Cell.Column.ToString() == "btnViernes")
			{
				if ((int)e.Cell.Row.Cells["Viernes"].Value == 0) return;

				iBodega = (int)e.Cell.Row.Cells["Viernes"].Value;

				sNomCol = e.Cell.Row.Cells["Viernes"].Column.Header.Caption;

				int iLength = 1;
				if (sNomCol.Length == 10) iLength = 2;

				iDia = Convert.ToInt32(sNomCol.Substring(8, iLength));
			}

			if (e.Cell.Column.ToString() == "btnSabado")
			{
				if ((int)e.Cell.Row.Cells["Sabado"].Value == 0) return;

				iBodega = (int)e.Cell.Row.Cells["Sabado"].Value;

				sNomCol = e.Cell.Row.Cells["Sabado"].Column.Header.Caption;

				int iLength = 1;
				if (sNomCol.Length == 9) iLength = 2;

				iDia = Convert.ToInt32(sNomCol.Substring(7, iLength));
			}

			if (e.Cell.Column.ToString() == "btnDomingo")
			{
				if ((int)e.Cell.Row.Cells["Domingo"].Value == 0) return;

				iBodega = (int)e.Cell.Row.Cells["Domingo"].Value;

				sNomCol = e.Cell.Row.Cells["Domingo"].Column.Header.Caption;

				int iLength = 1;
				if (sNomCol.Length == 10) iLength = 2;

				iDia = Convert.ToInt32(sNomCol.Substring(8, iLength));
			}
			
			DateTime dtFecha = DateTime.Parse(iDia + "/" + Convert.ToDateTime(this.dtpMes.Value).Date.Month + "/" + Convert.ToDateTime(this.dtpMes.Value).Date.Year);

			using (frmPlanificacionVentasVistaDiaria Vista = new frmPlanificacionVentasVistaDiaria(dtFecha, (int)e.Cell.Row.Cells["idDetallePlanificacionVentas"].Value, iBodega))
			{
				if (DialogResult.OK == Vista.ShowDialog())
				{

				}
			}
		}

		private void GuardaHorarios(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if ((int)this.txtIdPlanificacionVentas.Value == 0) return;

			if (e.Cell.Row.Cells["idDetallePlanificacionVentas"].Value == System.DBNull.Value) return;

			if ((int)e.Cell.Row.Cells["idDetallePlanificacionVentas"].Value == 0) return;

			using (frmFechaPlanificacionVentas ESCarpa = new frmFechaPlanificacionVentas(Funcion.dtEscalarSQL(cdsSeteoF, "Exec RetornaFechaServidor"), (int)e.Cell.Row.Cells["idDetallePlanificacionVentas"].Value))
			{
				if (DialogResult.OK == ESCarpa.ShowDialog())
				{
					try
					{
						DateTime dtEntrada = DateTime.Today;
						DateTime dtSalida = DateTime.Today;
						bool bHasta = false;

						if (ESCarpa.dtHoraSalida.Value != System.DBNull.Value) 
						{
							dtSalida = Convert.ToDateTime(ESCarpa.dtHoraSalida.Value);
							bHasta = true;
						}

						string sSQL = string.Format("Exec GuardaHorarioPlanificacionVentas {0}, {1}, {2}, '{3}', '{4}', '{5}', '{6}', {7}, {8}", 
							ESCarpa.idHorarioPlanificacionVentas, (int)e.Cell.Row.Cells["idDetallePlanificacionVentas"].Value, ESCarpa.iDia, 
							Convert.ToDateTime(ESCarpa.dtHoraEntrada.Value).ToString("yyyyMMdd HH:mm"), ESCarpa.txtObservacionesEntrada.Text.ToString(), 
							dtSalida.ToString("yyyyMMdd HH:mm"), ESCarpa.txtObservacionesSalida.Text.ToString(), 
							(int)ESCarpa.txtEstado.Value, bHasta);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
					}
					catch(Exception Exc)
					{
						MessageBox.Show(string.Format("Error al Aprobar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							
					}
				}
			}
		}

		private void uGridSemana1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana1, 1, 2);
		}

		private void uGridSemana2_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana2, 2, 2);
		}

		private void uGridSemana3_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana3, 3, 2);
		}

		private void uGridSemana4_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana4, 4, 2);
		}

		private void uGridSemana5_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana5, 5, 2);
		}

		private void uGridSemana6_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			FuncionesProcedimientos.IniciaValoresColumnasDetalleHorariosDeCarpa(sender, e, uGridSemana6, 6, 2);
		}

		private void uGridSemana1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridSemana1);
		}

		private void uGridSemana2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridSemana2);
		}

		private void uGridSemana3_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridSemana3);
		}

		private void uGridSemana4_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridSemana4);
		}

		private void uGridSemana5_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridSemana5);
		}

		private void uGridSemana6_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridSemana6);
		}

		private void uGridSemana1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnHorarios")
				GuardaHorarios(sender, e);

			if (e.Cell.Column.ToString() != "btnHorarios")
				this.ResumenDiario(sender, e);
		}

		private void uGridSemana2_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnHorarios")
				GuardaHorarios(sender, e);

			if (e.Cell.Column.ToString() != "btnHorarios")
				this.ResumenDiario(sender, e);
		}

		private void uGridSemana3_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnHorarios")
				GuardaHorarios(sender, e);

			if (e.Cell.Column.ToString() != "btnHorarios")
				this.ResumenDiario(sender, e);
		}

		private void uGridSemana4_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnHorarios")
				GuardaHorarios(sender, e);

			if (e.Cell.Column.ToString() != "btnHorarios")
				this.ResumenDiario(sender, e);
		}

		private void uGridSemana5_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnHorarios")
				GuardaHorarios(sender, e);

			if (e.Cell.Column.ToString() != "btnHorarios")
				this.ResumenDiario(sender, e);
		}

		private void uGridSemana6_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnHorarios")
				GuardaHorarios(sender, e);

			if (e.Cell.Column.ToString() != "btnHorarios")
				this.ResumenDiario(sender, e);
		}

		private void uGridSemana1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridSemana2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbNombres_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
